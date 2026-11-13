"""生成「原始 PDF 页面」与「转换后幻灯片」并排对比图。

用于人工核查还原效果，尤其是底图上是否残留了未被抹除的文字。

用法::

    python scripts/compare_source.py tests/materials/测试1.pdf outputs/测试1/测试1.pptx \
        -o outputs/测试1/compare
"""

from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

import pymupdf
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from scripts.render_preview import render_pptx  # noqa: E402


def render_pdf_pages(pdf_path: Path, output_dir: Path, zoom: float = 1.2) -> list[Path]:
    """把 PDF 的每一页渲染为 PNG。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    document = pymupdf.open(str(pdf_path))
    outputs: list[Path] = []
    for index, page in enumerate(document, start=1):
        pixmap = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        target = output_dir / f"source_{index:02d}.png"
        pixmap.save(str(target))
        outputs.append(target)
    document.close()
    return outputs


def make_sheet(left: Path, right: Path, target: Path) -> None:
    """把两张图左右拼接。"""
    with Image.open(left) as first, Image.open(right) as second:
        width = max(first.width, second.width)
        height = first.height + second.height + 24
        sheet = Image.new("RGB", (width, height), (245, 245, 245))
        sheet.paste(first, (0, 0))
        sheet.paste(second, (0, first.height + 24))
        sheet.save(target, format="PNG")


def main() -> None:
    parser = argparse.ArgumentParser(description="生成原图与转换结果的对比图")
    parser.add_argument("pdf", help="原始 PDF")
    parser.add_argument("pptx", help="转换后的 PPTX")
    parser.add_argument("-o", "--output", default="compare", help="输出目录")
    parser.add_argument("-z", "--zoom", type=float, default=1.2, help="缩放倍数")
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    pptx_path = Path(args.pptx)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    source_dir = output_dir / "source"
    sources = render_pdf_pages(pdf_path, source_dir, args.zoom)

    # 转换结果渲染只作为拼接素材，用完即删，避免与 preview/ 下的正式预览重复
    with tempfile.TemporaryDirectory(prefix="pdf2pptx_preview_") as temp_dir:
        previews = render_pptx(pptx_path, Path(temp_dir), args.zoom)
        for index, (source, preview) in enumerate(zip(sources, previews), start=1):
            target = output_dir / f"compare_{index:02d}.png"
            make_sheet(source, preview, target)
            print(f"已生成 {target}")


if __name__ == "__main__":
    main()