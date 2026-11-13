"""压缩验证材料中的 PNG 图片，减小仓库体积。

验证材料（预览图 / 对比图 / 原图）全部是渲染输出，属于「可以重新生成」的产物，
但仓库里仍需保留一份作为证据。直接由 PyMuPDF / Pillow 输出的 PNG 未做压缩，
体积偏大；本脚本用「调色板量化 + 无损压缩」在不改变文件格式与路径的前提下
显著缩小体积，对文字与线条类渲染图几乎看不出差别。

用法::

    python scripts/optimize_images.py                 # 压缩 outputs/ 下全部 PNG
    python scripts/optimize_images.py --dry-run       # 只统计，不写回
    python scripts/optimize_images.py --colors 96     # 调色板颜色上限（2~256）
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[1]


def optimize_one(path: Path, colors: int, dry_run: bool) -> tuple[int, int]:
    """压缩单个 PNG，返回 (原大小, 新大小)。"""
    original = path.stat().st_size
    with Image.open(path) as image:
        image.load()
        if image.mode in ("RGBA", "LA", "P"):
            # 带透明通道的图量化后需保留 alpha
            quantized = image.convert("RGBA").quantize(colors=colors, method=Image.FASTOCTREE)
        else:
            quantized = image.convert("RGB").quantize(colors=colors, method=Image.MEDIANCUT)
        if dry_run:
            return original, original
        tmp = path.with_suffix(".tmp.png")
        quantized.save(tmp, format="PNG", optimize=True)
    new_size = tmp.stat().st_size
    if new_size < original:
        tmp.replace(path)
        return original, new_size
    tmp.unlink(missing_ok=True)
    return original, original


def main() -> int:
    parser = argparse.ArgumentParser(description="压缩验证材料 PNG，减小仓库体积")
    parser.add_argument(
        "--root",
        default=str(ROOT / "outputs"),
        help="待压缩目录，默认 outputs/",
    )
    parser.add_argument("--colors", type=int, default=256, help="调色板颜色上限，默认 256")
    parser.add_argument("--dry-run", action="store_true", help="只统计不写回")
    args = parser.parse_args()

    root = Path(args.root)
    if not root.exists():
        print(f"目录不存在：{root}")
        return 1

    png_files = sorted(root.rglob("*.png"))
    if not png_files:
        print(f"{root} 下没有 PNG 文件")
        return 0

    total_before = total_after = 0
    changed = 0
    for path in png_files:
        before, after = optimize_one(path, args.colors, args.dry_run)
        total_before += before
        total_after += after
        if after < before:
            changed += 1
        ratio = (before - after) / before * 100 if before else 0
        print(f"{path.relative_to(root)}  {before / 1024:8.1f} KB -> {after / 1024:8.1f} KB  -{ratio:4.1f}%")

    saved = total_before - total_after
    print()
    print(f"共 {len(png_files)} 个文件，其中 {changed} 个被压缩")
    print(
        f"总体积 {total_before / 1024 / 1024:.1f} MB -> {total_after / 1024 / 1024:.1f} MB"
        f"（节省 {saved / 1024 / 1024:.1f} MB）"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
