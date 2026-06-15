# -*- coding: utf-8 -*-
"""
HTML 模板辅助生成脚本
仅用于开发阶段快速生成/更新模板，日常使用无需运行
"""

import os
from pathlib import Path

# 自动定位到脚本所在目录
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "index.html"


def build_template():
    """生成 HTML 模板文件（开发用）"""
    print(f"模板文件路径: {TEMPLATE_PATH}")
    if TEMPLATE_PATH.exists():
        print(f"模板已存在，大小: {TEMPLATE_PATH.stat().st_size} 字节")
        print("如需重新生成，请先删除现有模板文件")
    else:
        print("模板文件不存在，请从完整项目拷贝")
        print("或直接编辑 templates/index.html 文件")


if __name__ == "__main__":
    build_template()
