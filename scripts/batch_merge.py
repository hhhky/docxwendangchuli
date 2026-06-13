"""
批量合并 .docx 文件

用法：
  python batch_merge.py file1.docx file2.docx file3.docx -o merged.docx
  python batch_merge.py *.docx -o merged.docx --page-break
"""

import argparse
import glob
import os
import sys

try:
    from docx import Document
    from docx.shared import Inches, Pt, Cm, RGBColor
    from docx.enum.text import WD_BREAK
    from docx.oxml.ns import qn
except ImportError:
    print("需要安装 python-docx: pip install python-docx")
    sys.exit(1)


def merge_docx(input_files: list, output_path: str, separator: str = "page_break"):
    """将多个 docx 文件合并为一个。

    Args:
        input_files: 输入文件路径列表
        output_path: 输出文件路径
        separator: 分隔符类型 — "page_break"（分页）或 "none"（无分隔）
    """
    if not input_files:
        print("错误：未指定输入文件")
        sys.exit(1)

    if len(input_files) == 1:
        print("警告：只有一个输入文件，无需合并。直接复制。")
        import shutil
        shutil.copy2(input_files[0], output_path)
        print(f"已复制到: {output_path}")
        return

    # 验证所有文件存在
    missing = [f for f in input_files if not os.path.exists(f)]
    if missing:
        print(f"错误：以下文件不存在: {missing}")
        sys.exit(1)

    print(f"合并 {len(input_files)} 个文件...")

    # 以第一个文件为基础
    merged = Document(input_files[0])

    for i, filepath in enumerate(input_files[1:], 1):
        print(f"  [{i+1}/{len(input_files)}] {os.path.basename(filepath)}")
        src = Document(filepath)

        if separator == "page_break":
            # 在合并前加一个分页符
            merged.add_page_break()

        # 复制源文档的 body 元素到目标文档
        for element in src.element.body:
            # 跳过 sectPr（节属性），保留目标文档的节属性
            if element.tag == qn('w:sectPr'):
                continue
            merged.element.body.append(element)

    merged.save(output_path)
    print(f"合并完成: {output_path}  ({len(input_files)} 个文件)")


def expand_globs(patterns: list) -> list:
    """展开 glob 模式为文件列表，按名称排序。"""
    files = []
    for pattern in patterns:
        matched = glob.glob(pattern)
        if not matched:
            print(f"警告: '{pattern}' 没有匹配到任何文件")
        files.extend(matched)
    # 去重并排序
    seen = set()
    result = []
    for f in sorted(files):
        abs_path = os.path.abspath(f)
        if abs_path not in seen:
            seen.add(abs_path)
            result.append(f)
    return result


def main():
    parser = argparse.ArgumentParser(
        description="批量合并多个 .docx 文件为一个文件"
    )
    parser.add_argument(
        "inputs", nargs="+",
        help="输入的 .docx 文件（支持通配符，如 *.docx）"
    )
    parser.add_argument(
        "-o", "--output", required=True,
        help="输出文件路径"
    )
    parser.add_argument(
        "--separator", choices=["page_break", "none"], default="page_break",
        help="文档之间的分隔方式（默认: page_break）"
    )

    args = parser.parse_args()

    # 展开通配符
    files = expand_globs(args.inputs)

    # 过滤出 .docx 文件
    docx_files = [f for f in files if f.lower().endswith('.docx')]
    non_docx = [f for f in files if not f.lower().endswith('.docx')]
    if non_docx:
        print(f"跳过非 docx 文件: {non_docx}")

    if not docx_files:
        print("错误：没有找到有效的 .docx 文件")
        sys.exit(1)

    merge_docx(docx_files, args.output, args.separator)


if __name__ == "__main__":
    main()
