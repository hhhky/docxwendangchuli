"""
批量格式转换

用法：
  python batch_convert.py --from docx --to pdf file1.docx file2.docx
  python batch_convert.py --from docx --to md *.docx -o output/
  python batch_convert.py --from doc --to docx *.doc
  python batch_convert.py --from docx --to pdf ./documents/ -o ./pdfs/

支持的方向：
  docx → pdf, md, txt
  doc  → docx
  pdf  → docx（效果有限）

依赖：LibreOffice（soffice）, pandoc
"""

import argparse
import glob
import os
import subprocess
import sys
import time


def find_libreoffice():
    """查找 LibreOffice 可执行文件。"""
    # 先检查脚本目录下的 soffice.py 是否可用
    script_dir = os.path.dirname(os.path.abspath(__file__))
    soffice_py = os.path.join(script_dir, "office", "soffice.py")
    if os.path.exists(soffice_py):
        return soffice_py

    # try system soffice
    for name in ["soffice", "libreoffice", "soffice.exe", "libreoffice.exe"]:
        try:
            subprocess.run([name, "--version"], capture_output=True, timeout=5)
            return name
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    return None


def check_pandoc():
    """检查 pandoc 是否可用。"""
    try:
        subprocess.run(["pandoc", "--version"], capture_output=True, timeout=5)
        return "pandoc"
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return None


def resolve_files(inputs: list) -> list:
    """解析输入：展开通配符，识别目录，返回文件列表。"""
    files = []
    for item in inputs:
        if os.path.isdir(item):
            for f in os.listdir(item):
                full = os.path.join(item, f)
                if os.path.isfile(full):
                    files.append(full)
        elif '*' in item or '?' in item:
            files.extend(glob.glob(item))
        elif os.path.isfile(item):
            files.append(item)
        else:
            print(f"警告: 跳过无效输入 '{item}'")
    return sorted(set(os.path.abspath(f) for f in files))


def filter_by_ext(files: list, ext: str) -> list:
    """按扩展名过滤文件（不区分大小写）。"""
    ext = ext.lower().lstrip('.')
    return [f for f in files if f.lower().endswith(f'.{ext}')]


def convert_via_libreoffice(files: list, from_ext: str, to_ext: str, output_dir: str, soffice_cmd, timeout: int = 300):
    """使用 LibreOffice 批量转换。"""
    results = {"ok": [], "failed": []}

    for i, f in enumerate(files):
        basename = os.path.splitext(os.path.basename(f))[0]
        out_file = os.path.join(output_dir, f"{basename}.{to_ext}")
        print(f"  [{i+1}/{len(files)}] {os.path.basename(f)} → {basename}.{to_ext}")

        try:
            if soffice_cmd.endswith('.py'):
                cmd = [
                    sys.executable, soffice_cmd,
                    "--headless",
                    "--convert-to", to_ext,
                    "--outdir", output_dir,
                    f
                ]
            else:
                cmd = [
                    soffice_cmd,
                    "--headless",
                    "--convert-to", to_ext,
                    "--outdir", output_dir,
                    f
                ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
            if result.returncode != 0:
                print(f"    错误: {result.stderr.strip()}")
                results["failed"].append(f)
            elif os.path.exists(out_file):
                results["ok"].append(out_file)
            else:
                # 尝试找一下生成的文件（LibreOffice 有时用不同命名）
                found = glob.glob(os.path.join(output_dir, f"{basename}*"))
                matched = [x for x in found if x.endswith(f".{to_ext}")]
                if matched:
                    results["ok"].append(matched[0])
                else:
                    print(f"    警告: 未找到输出文件")
                    results["failed"].append(f)
        except subprocess.TimeoutExpired:
            print(f"    超时 (>{timeout}s)")
            results["failed"].append(f)
        except Exception as e:
            print(f"    异常: {e}")
            results["failed"].append(f)

    return results


def convert_docx_to_md(files: list, output_dir: str):
    """使用 pandoc 将 docx 转为 markdown。"""
    pandoc = check_pandoc()
    if not pandoc:
        print("错误: 需要安装 pandoc (https://pandoc.org/)")
        return {"ok": [], "failed": list(files)}

    results = {"ok": [], "failed": []}
    for i, f in enumerate(files):
        basename = os.path.splitext(os.path.basename(f))[0]
        out_file = os.path.join(output_dir, f"{basename}.md")
        print(f"  [{i+1}/{len(files)}] {os.path.basename(f)} → {basename}.md")

        try:
            result = subprocess.run(
                [pandoc, f, "-t", "markdown", "-o", out_file],
                capture_output=True, text=True, timeout=120
            )
            if result.returncode == 0:
                results["ok"].append(out_file)
            else:
                print(f"    错误: {result.stderr.strip()}")
                results["failed"].append(f)
        except Exception as e:
            print(f"    异常: {e}")
            results["failed"].append(f)

    return results


def convert_docx_to_txt(files: list, output_dir: str):
    """使用 pandoc 将 docx 转为纯文本。"""
    pandoc = check_pandoc()
    if not pandoc:
        print("错误: 需要安装 pandoc")
        return {"ok": [], "failed": list(files)}

    results = {"ok": [], "failed": []}
    for i, f in enumerate(files):
        basename = os.path.splitext(os.path.basename(f))[0]
        out_file = os.path.join(output_dir, f"{basename}.txt")
        print(f"  [{i+1}/{len(files)}] {os.path.basename(f)} → {basename}.txt")

        try:
            result = subprocess.run(
                [pandoc, f, "-t", "plain", "-o", out_file],
                capture_output=True, text=True, timeout=120
            )
            if result.returncode == 0:
                results["ok"].append(out_file)
            else:
                print(f"    错误: {result.stderr.strip()}")
                results["failed"].append(f)
        except Exception as e:
            print(f"    异常: {e}")
            results["failed"].append(f)

    return results


def main():
    parser = argparse.ArgumentParser(
        description="批量文档格式转换"
    )
    parser.add_argument(
        "inputs", nargs="+",
        help="输入文件/目录/通配符"
    )
    parser.add_argument(
        "--from", dest="from_fmt", required=True,
        choices=["docx", "doc", "pdf"],
        help="源格式"
    )
    parser.add_argument(
        "--to", dest="to_fmt", required=True,
        choices=["pdf", "md", "txt", "docx"],
        help="目标格式"
    )
    parser.add_argument(
        "-o", "--output", default=".",
        help="输出目录（默认: 当前目录）"
    )
    parser.add_argument(
        "--timeout", type=int, default=300,
        help="单文件超时秒数（默认: 300）"
    )

    args = parser.parse_args()

    from_fmt = args.from_fmt
    to_fmt = args.to_fmt

    if from_fmt == to_fmt:
        print("错误: 源格式和目标格式相同")
        sys.exit(1)

    # 创建输出目录
    output_dir = os.path.abspath(args.output)
    os.makedirs(output_dir, exist_ok=True)

    # 解析输入文件
    all_files = resolve_files(args.inputs)
    files = filter_by_ext(all_files, from_fmt)

    print(f"\n找到 {len(files)} 个 .{from_fmt} 文件")

    skipped = [f for f in all_files if f not in files]
    if skipped:
        print(f"跳过 {len(skipped)} 个非 .{from_fmt} 文件")

    if not files:
        print("没有可转换的文件")
        sys.exit(0)

    print(f"开始转换: .{from_fmt} → .{to_fmt}\n")

    # 选择转换方法
    if to_fmt == "md" and from_fmt == "docx":
        results = convert_docx_to_md(files, output_dir)
    elif to_fmt == "txt" and from_fmt == "docx":
        results = convert_docx_to_txt(files, output_dir)
    elif from_fmt in ("docx", "doc") and to_fmt == "pdf":
        soffice = find_libreoffice()
        if not soffice:
            print("错误: 转换 docx→PDF 需要 LibreOffice")
            print("  安装: https://www.libreoffice.org/download/")
            sys.exit(1)
        results = convert_via_libreoffice(files, from_fmt, to_fmt, output_dir, soffice, args.timeout)
    elif from_fmt == "doc" and to_fmt == "docx":
        soffice = find_libreoffice()
        if not soffice:
            print("错误: 转换 doc→docx 需要 LibreOffice")
            sys.exit(1)
        results = convert_via_libreoffice(files, from_fmt, to_fmt, output_dir, soffice, args.timeout)
    elif from_fmt == "pdf" and to_fmt == "docx":
        soffice = find_libreoffice()
        if not soffice:
            print("错误: 转换 pdf→docx 需要 LibreOffice")
            sys.exit(1)
        results = convert_via_libreoffice(files, from_fmt, to_fmt, output_dir, soffice, args.timeout)
    else:
        print(f"错误: 不支持的转换方向 .{from_fmt} → .{to_fmt}")
        sys.exit(1)

    # 汇总
    print(f"\n{'='*50}")
    print(f"成功: {len(results['ok'])}")
    print(f"失败: {len(results['failed'])}")
    if results['failed']:
        print("失败的文件:")
        for f in results['failed']:
            print(f"  - {f}")
    if results['ok']:
        print(f"输出目录: {output_dir}")


if __name__ == "__main__":
    main()
