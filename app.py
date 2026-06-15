#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF ↔ Word 双向转换 Web 应用程序
基于 Flask + Microsoft Word COM 自动化引擎
提供 PDF 转 Word、Word 转 PDF 双向高保真转换
"""

import os
import sys
import uuid
import time
import threading
import traceback
import subprocess
from pathlib import Path
from typing import Tuple, Optional, Dict, Any
from datetime import datetime, timedelta

from flask import (
    Flask, render_template, request, jsonify,
    send_from_directory, redirect, url_for
)
from werkzeug.utils import secure_filename

# ------------------------------------------------------------
# 配置
# ------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
CONVERSION_DIR = BASE_DIR / "conversions"
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"

ALLOWED_EXTENSIONS_PDF = {".pdf"}
ALLOWED_EXTENSIONS_WORD = {".doc", ".docx"}
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB
MAX_FILE_AGE = timedelta(hours=1)  # 超过 1 小时的旧文件自动清理
MAX_TEMP_FILES = 10  # 每个临时文件夹最多保留 10 个文件，超出自动删除最旧的
CONVERSION_TIMEOUT = 120  # 转换超时时间（秒）

# 确保目录存在
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
CONVERSION_DIR.mkdir(parents=True, exist_ok=True)

# 创建 Flask 应用
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE
app.config["TEMPLATES_AUTO_RELOAD"] = True  # 确保模板实时更新

# 任务状态存储
tasks_lock = threading.Lock()
tasks = {}

# 结果文件名 → 原始显示名映射表
result_name_map_lock = threading.Lock()
result_name_map = {}

def register_result_name(stored_name: str, display_name: str):
    """注册存储文件名与显示文件名的映射"""
    with result_name_map_lock:
        result_name_map[stored_name] = display_name

def get_display_name(stored_name: str) -> str:
    """获取存储文件对应的显示文件名"""
    with result_name_map_lock:
        return result_name_map.get(stored_name, stored_name)


# ------------------------------------------------------------
# 文件安全 & 验证
# ------------------------------------------------------------
def is_allowed_file(filename: str, allowed_exts: set) -> bool:
    """检查文件扩展名是否在允许范围内"""
    ext = Path(filename).suffix.lower()
    return ext in allowed_exts


def validate_file_header(filepath: Path, expected_type: str) -> bool:
    """
    通过文件头部魔数验证文件格式
    PDF 文件头: %PDF
    DOCX 文件头: PK (ZIP格式，50 4B)
    DOC 文件头:  D0 CF 11 E0 A1 B1 1A E1 (OLE2)
    """
    try:
        with open(filepath, "rb") as f:
            header = f.read(8)

        if expected_type == "pdf":
            return header.startswith(b"%PDF")
        elif expected_type == "docx":
            return header[:2] == b"PK"
        elif expected_type == "doc":
            return header[:8] == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
        return False
    except Exception:
        return False


def secure_save(uploaded_file, upload_dir: Path) -> Tuple[str, str]:
    """
    安全保存上传的文件，返回 (存储文件名, 原始文件名)
    使用 UUID 重命名以防止路径穿越和文件名冲突
    """
    original_name = secure_filename(uploaded_file.filename or "unnamed")
    ext = Path(original_name).suffix.lower()
    stored_name = uuid.uuid4().hex + ext
    filepath = upload_dir / stored_name
    uploaded_file.save(str(filepath))
    return stored_name, original_name


def cleanup_old_files():
    """
    清理临时文件：
    1. 删除超过 MAX_FILE_AGE 的旧文件
    2. 每个目录最多保留 MAX_TEMP_FILES 个文件，超出则删除最旧的
    """
    now = datetime.now()
    for d in [UPLOAD_DIR, CONVERSION_DIR]:
        # 第一步：删除过期文件
        for f in list(d.iterdir()):
            if f.is_file():
                mtime = datetime.fromtimestamp(f.stat().st_mtime)
                if now - mtime > MAX_FILE_AGE:
                    try:
                        f.unlink()
                    except Exception:
                        pass

        # 第二步：如果文件数量超过 MAX_TEMP_FILES，删除最旧的文件
        files = sorted(
            [f for f in d.iterdir() if f.is_file()],
            key=lambda f: f.stat().st_mtime
        )
        while len(files) > MAX_TEMP_FILES:
            try:
                files[0].unlink()
                files.pop(0)
            except Exception:
                pass


def create_task(task_type: str, source_name: str) -> str:
    """创建新任务并返回 task_id"""
    task_id = uuid.uuid4().hex
    with tasks_lock:
        tasks[task_id] = {
            "task_id": task_id,
            "type": task_type,
            "source_name": source_name,
            "progress": 0,
            "status": "queued",       # queued → processing → done / error
            "result_file": None,
            "error": None,
            "created_at": time.time(),
        }
    return task_id


def update_task(task_id: str, **kwargs):
    """更新任务状态"""
    with tasks_lock:
        if task_id in tasks:
            tasks[task_id].update(kwargs)


def get_task(task_id: str) -> dict:
    """获取任务状态快照"""
    with tasks_lock:
        return dict(tasks.get(task_id, {}))


def clean_expired_tasks():
    """清理过期任务（超过1小时）"""
    now = time.time()
    with tasks_lock:
        expired = [
            tid for tid, t in tasks.items()
            if now - t.get("created_at", 0) > 3600
        ]
        for tid in expired:
            del tasks[tid]


# ------------------------------------------------------------
# 转换功能
# ------------------------------------------------------------
def convert_pdf_to_word(source_path: Path, output_path: Path, task_id: str):
    """
    PDF → Word 1:1 精确转换
    使用 Microsoft Word 原生引擎，通过 PowerShell COM 自动化
    利用 Word 的 PDF 重排引擎实现最高保真度
    """
    import subprocess

    update_task(task_id, status="processing", progress=5)

    try:
        src_abs = str(source_path.absolute()).replace('\\', '/')
        out_abs = str(output_path.absolute()).replace('\\', '/')

        output_path.parent.mkdir(parents=True, exist_ok=True)

        # ── Step 1: 预扫描 PDF 中的下划线/横线 ──
        update_task(task_id, progress=10)
        ul_info = _scan_pdf_for_underlines(source_path)

        # ── Step 2: Word COM 引擎转换 ──
        update_task(task_id, progress=20)

        ps_exe = os.path.join(
            os.environ.get('SystemRoot', 'C:\\Windows'),
            'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe'
        )

        ps_script = (
            '$ErrorActionPreference = "Stop"; '
            '$word = New-Object -ComObject Word.Application; '
            '$word.Visible = $false; '
            '$word.DisplayAlerts = 0; '
            'try { '
            f"$doc = $word.Documents.Open('{src_abs}'); "
            f"$doc.SaveAs('{out_abs}', 16); "
            '$doc.Close(); '
            '} finally { '
            '$word.Quit(); '
            '}'
        )

        update_task(task_id, progress=40)

        result = subprocess.run(
            [ps_exe, "-NoProfile", "-Command", ps_script],
            capture_output=True, text=True, timeout=CONVERSION_TIMEOUT
        )

        if result.returncode != 0:
            err = result.stderr.strip() or result.stdout.strip() or "未知错误"
            raise Exception(f"Word COM 转换失败 (code={result.returncode}): {err}")

        if not output_path.exists():
            raise Exception("Word COM 转换完成但未生成 DOCX 文件")

        # ── Step 3: 后处理 — 恢复丢失的下划线/横线 ──
        update_task(task_id, progress=80)
        _restore_underlines_in_docx(output_path, ul_info)

        update_task(task_id, progress=100, status="done",
                    result_file=output_path.name)

    except Exception as e:
        error_msg = f"PDF 转 Word 失败: {str(e)}"
        traceback.print_exc()
        update_task(task_id, status="error", error=error_msg)
        raise


def _scan_pdf_for_underlines(pdf_path: Path) -> list:
    """
    用 PyMuPDF 扫描 PDF，检测所有下划线和横线文本
    返回: [{"text": "...", "underline": True/False, "page": N, ...}, ...]
    检测三种类型：
      A) 字体下划线（span flags 含 underline）
      B) 连续下划线字符（_____）
      C) 向量横线（图形线段）
    """
    import fitz
    results = []
    try:
        doc = fitz.open(str(pdf_path))
    except Exception:
        return results

    for page_num in range(len(doc)):
        page = doc[page_num]
        # 方法A：从文本 spans 检测下划线格式
        try:
            text_dict = page.get_text("dict")
            for block in text_dict.get("blocks", []):
                if block.get("type") != 0:  # 0=text
                    continue
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "").strip()
                        if not text:
                            continue
                        # 下划线检测：flags=4 表示 underline
                        flags = span.get("flags", 0)
                        is_underlined = bool(flags & 4)  # bit 2 = underline
                        if is_underlined:
                            results.append({
                                "text": text, "type": "underline",
                                "page": page_num, "y": span.get("bbox", [0,0,0,0])[1]
                            })
                        # 检测连续下划线字符
                        if "_" in text:
                            stripped = text.strip()
                            if len(stripped) >= 2 and all(c == "_" for c in stripped):
                                results.append({
                                    "text": text, "type": "underscore_chars",
                                    "page": page_num, "y": span.get("bbox", [0,0,0,0])[1]
                                })
                            # 也检测混合文本含下划线（如"____ 姓名 ____"）
                            elif "_" in stripped:
                                results.append({
                                    "text": text, "type": "mixed_underscore",
                                    "page": page_num, "y": span.get("bbox", [0,0,0,0])[1]
                                })
        except Exception:
            pass

        # 方法B：向量横线检测
        try:
            for path in page.get_drawings():
                for item in path.get("items", []):
                    if item[0] == "l":
                        x1, y1, x2, y2 = float(item[1]), float(item[2]), float(item[3]), float(item[4])
                        if abs(y1 - y2) < 4 and abs(x2 - x1) > 80:
                            y_mid = (y1 + y2) / 2
                            results.append({
                                "type": "vector_line", "page": page_num,
                                "y": y_mid, "x0": x1, "x1": x2
                            })
        except Exception:
            pass

    doc.close()
    return results


def _restore_underlines_in_docx(docx_path: Path, ul_data: list):
    """
    在 docx 中恢复/修正下划线

    Word COM 的 PDF 重排引擎会把填空线区域转为「空格 + underline 格式」，
    但也会把相邻文字错误地继承下划线（如「计算机与信息技术学院」被误加下划线）。

    本函数的核心策略：
    1. 已有 underline 格式的 run → 精细化处理
       a. 纯空格/制表符 → 保持空格不变，确保有 underline 格式
          （Word 中「空格+underline」渲染为连续实心横线 = PDF 原版样式）
       b. 混合文字+空格/_ → 拆分 run，仅填空线部分保留下划线，文字部分移除
       c. 纯文字/标点 → 移除 underline（Word COM 误加）
    2. 无 underline 但含 _ 字符的 run → 加下划线（兜底逻辑）
    """
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    try:
        doc = Document(str(docx_path))
        modified = False

        for para in doc.paragraphs:
            runs_to_check = list(para.runs)
            for run in runs_to_check:
                text = run.text
                if not text:
                    continue

                has_ul = (run.font.underline is True)

                if has_ul:
                    # ════════════════════════════════════════
                    # 分支 A：run 已有下划线格式
                    # ════════════════════════════════════════

                    # 检查是否只由填空线字符组成（_、空格、制表符）
                    core = text.strip('_ \t')
                    if not core:
                        # A1：纯填空线（空格/_/制表符）
                        # ★ 保持原文本（空格），只确保有 underline 格式
                        #   Word 中「空格+underline」渲染为连续实心横线，与原版 PDF 一致
                        #   不要替换为 _ 字符（有间隙，和原版实线不一致）
                        _set_underline_on_element(run._element, True)
                        modified = True
                    elif ' ' in text or '\t' in text or '_' in text:
                        # A2：混合文字+填空线 → 拆分，只对填空线部分保留 underline
                        _split_and_underline_run(para, run, ul_chars='_ \t')
                        modified = True
                    else:
                        # A3：纯文字/标点 → Word COM 误加下划线 → 移除
                        _remove_run_underline(run)
                        modified = True
                else:
                    # ════════════════════════════════════════
                    # 分支 B：run 无下划线格式 → _ 字符检测（兜底）
                    # ════════════════════════════════════════
                    if '_' not in text:
                        continue
                    stripped = text.strip()
                    if not stripped:
                        continue
                    if all(c == '_' for c in stripped) and len(stripped) >= 2:
                        run.font.underline = True
                        modified = True
                    else:
                        _split_and_underline_run(para, run, ul_chars='_')
                        modified = True

        if modified:
            doc.save(str(docx_path))
    except Exception as e:
        print(f"[_restore_underlines_in_docx] ERROR: {e}", flush=True)
        import traceback
        traceback.print_exc()


def _remove_run_underline(run):
    """移除 run 的下划线格式（通过 XML 直接删除 w:u 元素）"""
    from docx.oxml.ns import qn
    rPr = run._element.find(qn('w:rPr'))
    if rPr is None:
        return
    u = rPr.find(qn('w:u'))
    if u is not None:
        rPr.remove(u)


def _set_underline_on_element(element, enable: bool):
    """
    通过 XML 操作设置/移除 w:r 元素的下划线格式
    enable=True  → 确保存在 w:u w:val="single"
    enable=False → 确保移除 w:u 元素
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    rPr = element.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        element.insert(0, rPr)
    u = rPr.find(qn('w:u'))
    if enable:
        if u is None:
            u = OxmlElement('w:u')
            rPr.append(u)
        u.set(qn('w:val'), 'single')
    else:
        if u is not None:
            rPr.remove(u)


def _split_and_underline_run(para, run, ul_chars='_'):
    """
    拆分段落中的 run，仅对 ul_chars 中的字符部分加下划线

    ul_chars: 哪些字符被视为「填空线标记」
      默认 '_'             → 处理显式 _ 字符
      可传 '_ \t' 或 ' \t' → 同时处理空格和制表符（Word COM 转换产物）

    对于 ul_chars 中的空格/制表符，保持原文本不变（Word 中空格+underline
    渲染为连续实心横线，与 PDF 原版一致）。非填空线部分如果有 underline
    格式则移除（修正 Word COM 误加）。
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import copy

    text = run.text
    if not text:
        return

    # ── 分段：连续 ul_chars 字符 vs 连续非 ul_chars 字符 ──
    segments = []
    buf = ""
    in_ul = False
    for ch in text:
        is_ul = (ch in ul_chars)
        if is_ul == in_ul:
            buf += ch
        else:
            if buf:
                segments.append((buf, in_ul))
            buf = ch
            in_ul = is_ul
    if buf:
        segments.append((buf, in_ul))

    # ── 只有一段 → 不需要拆分 ──
    if len(segments) <= 1:
        if segments and segments[0][1]:
            # 全是填空线字符 → 确保有 underline（保持空格不变）
            _set_underline_on_element(run._element, True)
        else:
            # 全是非填空线字符 + 有 underline → 移除误加
            if run.font.underline is True:
                _remove_run_underline(run)
        return

    # ── 多段 → 拆分 run ──
    parent = run._element.getparent()
    first = True
    insert_pos = list(parent).index(run._element)
    pos = insert_pos + 1  # 从原 run 后面开始逐个插入

    for seg_text, is_underline in segments:
        if first:
            # 保留原 run
            run.text = seg_text
            _set_underline_on_element(run._element, is_underline)
            first = False
        else:
            # 创建新 run，复制原 run 的格式属性
            new_r = copy.deepcopy(run._element)
            for child in list(new_r):
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 't':
                    child.text = seg_text
            _set_underline_on_element(new_r, is_underline)
            parent.insert(pos, new_r)
            pos += 1


def convert_word_to_pdf(source_path: Path, output_path: Path, task_id: str):
    """
    Word → PDF 1:1 精确转换
    使用 Microsoft Word 原生引擎，通过 PowerShell COM 自动化
    保证格式（字体/字号/颜色/间距/对齐/边距）100% 还原
    """
    import subprocess

    update_task(task_id, status="processing", progress=10)

    try:
        # 确保路径使用绝对路径并规范化（PowerShell 兼容）
        src_abs = str(source_path.absolute()).replace('\\', '/')
        out_abs = str(output_path.absolute()).replace('\\', '/')

        # 确保输出目录存在
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # PowerShell 路径
        ps_exe = os.path.join(
            os.environ.get('SystemRoot', 'C:\\Windows'),
            'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe'
        )

        # 使用单引号包裹路径防止转义问题
        ps_script = (
            '$ErrorActionPreference = "Stop"; '
            '$word = New-Object -ComObject Word.Application; '
            '$word.Visible = $false; '
            '$word.DisplayAlerts = 0; '
            'try { '
            f"$doc = $word.Documents.Open('{src_abs}'); "
            f"$doc.SaveAs('{out_abs}', 17); "
            '$doc.Close(); '
            '} finally { '
            '$word.Quit(); '
            '}'
        )

        update_task(task_id, progress=30)

        result = subprocess.run(
            [ps_exe, "-NoProfile", "-Command", ps_script],
            capture_output=True, text=True, timeout=CONVERSION_TIMEOUT
        )

        if result.returncode != 0:
            err = result.stderr.strip() or result.stdout.strip() or "未知错误"
            raise Exception(f"Word COM 转换失败 (code={result.returncode}): {err}")

        if not output_path.exists():
            raise Exception("Word COM 转换完成但未生成 PDF 文件")

        update_task(task_id, progress=100, status="done",
                    result_file=output_path.name)

    except Exception as e:
        error_msg = f"Word 转 PDF 失败: {str(e)}"
        traceback.print_exc()
        update_task(task_id, status="error", error=error_msg)
        raise


# ------------------------------------------------------------
# 后台任务执行器
# ------------------------------------------------------------
def run_conversion(task_id: str, source_path: Path, output_path: Path,
                   conv_type: str):
    """在后台线程中执行转换任务"""
    try:
        if conv_type == "pdf2word":
            convert_pdf_to_word(source_path, output_path, task_id)
        else:
            convert_word_to_pdf(source_path, output_path, task_id)
    except Exception as e:
        # 错误已在子函数中记录，这里只做清理
        pass
    finally:
        # 清理上传的源文件
        try:
            if source_path.exists():
                source_path.unlink()
        except Exception:
            pass


# ------------------------------------------------------------
# Flask 路由
# ------------------------------------------------------------
@app.route("/")
def index():
    """渲染主页面"""
    cleanup_old_files()
    clean_expired_tasks()
    return render_template("index.html")


@app.route("/upload/pdf2word", methods=["POST"])
def upload_pdf2word():
    """上传 PDF 并触发转换"""
    return handle_upload("pdf2word", ALLOWED_EXTENSIONS_PDF, "pdf")


@app.route("/upload/word2pdf", methods=["POST"])
def upload_word2pdf():
    """上传 Word 并触发转换"""
    return handle_upload("word2pdf", ALLOWED_EXTENSIONS_WORD, "docx")


def handle_upload(conv_type: str, allowed_exts: set, expected_header: str):
    """通用文件上传处理"""
    if "file" not in request.files:
        return jsonify({"error": "未选择文件"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "文件名为空"}), 400

    # 验证扩展名
    if not is_allowed_file(file.filename, allowed_exts):
        suffix = Path(file.filename).suffix.lower()
        if conv_type == "pdf2word":
            return jsonify({"error": f"不支持的文件格式 '{suffix}'，请上传 PDF 文件"}), 400
        else:
            return jsonify({"error": f"不支持的文件格式 '{suffix}'，请上传 Word 文件 (.doc/.docx)"}), 400

    # 检查文件大小（Content-Length 可能不准确，使用流式读取）
    file.seek(0, os.SEEK_END)
    file_size = file.tell()
    file.seek(0)
    if file_size > MAX_FILE_SIZE:
        max_mb = MAX_FILE_SIZE // (1024 * 1024)
        return jsonify({"error": f"文件大小不能超过 {max_mb}MB"}), 400

    # 检查 file_size 是否为 0
    if file_size == 0:
        return jsonify({"error": "上传的文件为空"}), 400

    # 安全保存文件
    try:
        stored_name, original_name = secure_save(file, UPLOAD_DIR)
    except Exception as e:
        return jsonify({"error": f"文件保存失败: {str(e)}"}), 500

    # 验证文件头部魔数
    filepath = UPLOAD_DIR / stored_name
    if not validate_file_header(filepath, expected_header):
        filepath.unlink(missing_ok=True)
        if conv_type == "pdf2word":
            return jsonify({"error": "文件格式验证失败，不是有效的 PDF 文件"}), 400
        else:
            return jsonify({"error": "文件格式验证失败，不是有效的 Word 文件"}), 400

    # 创建任务
    task_id = create_task(conv_type, original_name)

    # 确定输出文件名 — 使用原始文件名保持一致性
    base_name = Path(original_name).stem  # 去掉扩展名
    if conv_type == "pdf2word":
        output_name = base_name + ".docx"
    else:
        output_name = base_name + ".pdf"

    # 如果同名文件已存在，添加数字后缀
    output_path = CONVERSION_DIR / output_name
    counter = 1
    while output_path.exists():
        if conv_type == "pdf2word":
            output_name = f"{base_name}_{counter}.docx"
        else:
            output_name = f"{base_name}_{counter}.pdf"
        output_path = CONVERSION_DIR / output_name
        counter += 1

    # 注册文件名映射（使用存储名 → 显示名）
    register_result_name(output_name, output_name)

    # 启动后台转换线程
    t = threading.Thread(
        target=run_conversion,
        args=(task_id, filepath, output_path, conv_type),
        daemon=True
    )
    t.start()

    return jsonify({
        "task_id": task_id,
        "message": "文件上传成功，开始转换",
        "source_name": original_name,
    })


@app.route("/status/<task_id>")
def get_status(task_id: str):
    """查询转换进度"""
    task = get_task(task_id)
    if not task:
        return jsonify({"error": "任务不存在"}), 404
    return jsonify(task)


@app.route("/download/<filename>")
def download_file(filename: str):
    """下载转换后的文件"""
    # 安全检查：防止路径穿越
    if ".." in filename or "/" in filename or "\\" in filename:
        return jsonify({"error": "非法的文件名"}), 400

    filepath = CONVERSION_DIR / filename
    if not filepath.exists():
        return jsonify({"error": "文件不存在或已过期"}), 404

    # 获取显示文件名
    display_name = get_display_name(filename)

    return send_from_directory(
        str(CONVERSION_DIR),
        filename,
        as_attachment=True,
        download_name=display_name
    )


@app.route("/api/open-temp-folder")
def open_temp_folder():
    """打开临时文件夹（在文件资源管理器中）"""
    try:
        subprocess.Popen(['explorer', str(CONVERSION_DIR.absolute())])
        return jsonify({"status": "ok", "message": "已打开临时文件夹"})
    except Exception as e:
        return jsonify({"status": "error", "message": f"无法打开文件夹: {str(e)}"}), 500


@app.route("/api/clear-temp-files", methods=["POST"])
def clear_temp_files():
    """手动清理所有临时文件"""
    try:
        deleted_count = 0
        for d in [UPLOAD_DIR, CONVERSION_DIR]:
            for f in list(d.iterdir()):
                if f.is_file():
                    try:
                        f.unlink()
                        deleted_count += 1
                    except Exception:
                        pass
        # 清理过期任务
        clean_expired_tasks()
        return jsonify({
            "status": "ok",
            "message": f"已清理 {deleted_count} 个临时文件",
            "count": deleted_count
        })
    except Exception as e:
        return jsonify({"status": "error", "message": f"清理失败: {str(e)}"}), 500


@app.route("/api/temp-file-count")
def temp_file_count():
    """获取临时文件数量"""
    count = 0
    for d in [UPLOAD_DIR, CONVERSION_DIR]:
        count += sum(1 for f in d.iterdir() if f.is_file())
    return jsonify({"count": count, "max": MAX_TEMP_FILES * 2})


@app.route("/temp-files")
def temp_files():
    """查看临时文件夹中的文件列表"""
    cleanup_old_files()
    files_info = []
    for d_name, d_path in [("uploads", UPLOAD_DIR), ("conversions", CONVERSION_DIR)]:
        for f in sorted(d_path.iterdir(), key=lambda x: x.stat().st_mtime, reverse=True):
            if f.is_file():
                mtime = datetime.fromtimestamp(f.stat().st_mtime)
                files_info.append({
                    "name": f.name,
                    "dir": d_name,
                    "size": f.stat().st_size,
                    "modified": mtime.strftime("%Y-%m-%d %H:%M:%S"),
                    "age_minutes": int((datetime.now() - mtime).total_seconds() / 60)
                })
    return jsonify(files_info)


@app.errorhandler(413)
def request_entity_too_large(error):
    """文件超过大小限制"""
    max_mb = MAX_FILE_SIZE // (1024 * 1024)
    return jsonify({"error": f"文件大小不能超过 {max_mb}MB"}), 413


@app.errorhandler(500)
def internal_error(error):
    """服务器内部错误"""
    return jsonify({"error": "服务器内部错误，请重试"}), 500


# ------------------------------------------------------------
# 入口
# ------------------------------------------------------------
if __name__ == "__main__":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    print("=" * 60)
    print("  PDF ↔ Word 双向转换工具")
    print(f"  本地地址: http://127.0.0.1:5000")
    print(f"  局域网地址: http://0.0.0.0:5000")
    print("  按 Ctrl+C 停止服务")
    print("=" * 60)

    # debug=True 开发模式，生产环境建议设为 False
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
