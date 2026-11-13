"""PPTX 可编辑性验证脚本。

读取生成的 PPTX，逐页列出其中的对象类型（文本框 / 图片 / 形状 / 表格 /
连接符），用于证明「页面内容不是一张整页图片，而是可独立编辑的原生对象」。

用法::

    python scripts/verify_pptx.py outputs/测试1/测试1.pptx
    python scripts/verify_pptx.py outputs/*/*.pptx -o verification/objects.txt
"""

from __future__ import annotations

import argparse
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def shape_kind(shape) -> str:
    if shape.has_table:
        return "表格"
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "图片"
    if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
        return "文本框"
    if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
        return "形状"
    if shape.shape_type == MSO_SHAPE_TYPE.LINE:
        return "连接符"
    return str(shape.shape_type)


def describe(shape) -> str:
    kind = shape_kind(shape)
    if kind == "文本框":
        text = shape.text_frame.text.replace("\n", " ").strip()
        if not text:
            return f"{kind}（空）"
        size, spacing = _run_style(shape)
        extra = f"{size:.1f}pt" if size else "-"
        if spacing:
            extra += f" 字距{spacing:.1f}pt"
        return f"{kind}「{text[:34]}」 {extra}"
    if kind == "表格":
        table = shape.table
        return f"{kind} {len(table.rows)} 行 × {len(table.columns)} 列"
    if kind == "图片":
        return f"{kind} {shape.width / 914400:.2f}in × {shape.height / 914400:.2f}in"
    return f"{kind} {shape.width / 914400:.2f}in × {shape.height / 914400:.2f}in"


def _run_style(shape) -> tuple[float | None, float]:
    """取首个 run 的字号（磅）与字符间距（磅）。"""
    try:
        paragraph = shape.text_frame.paragraphs[0]
        if not paragraph.runs:
            return None, 0.0
        run = paragraph.runs[0]
        size = run.font.size.pt if run.font.size else None
        spacing = 0.0
        rPr = run._r.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}rPr"
        )
        if rPr is not None and rPr.get("spc"):
            spacing = int(rPr.get("spc")) / 100.0
        return size, spacing
    except Exception:
        return None, 0.0


def inspect(path: Path) -> tuple[list[str], dict[str, int], dict]:
    """返回（文本行, 对象类型总计, 附加统计）。"""
    lines: list[str] = []
    presentation = Presentation(str(path))

    lines.append(f"文件：{path}")
    lines.append(
        f"页面尺寸：{presentation.slide_width / 914400:.2f}in × "
        f"{presentation.slide_height / 914400:.2f}in"
    )
    lines.append(f"页数：{len(presentation.slides)}")
    lines.append("-" * 78)

    totals: dict[str, int] = {}
    text_chars = 0
    for index, slide in enumerate(presentation.slides, start=1):
        counts: dict[str, int] = {}
        lines.append(f"第 {index} 页，共 {len(slide.shapes)} 个对象：")
        for shape in slide.shapes:
            kind = shape_kind(shape)
            counts[kind] = counts.get(kind, 0) + 1
            totals[kind] = totals.get(kind, 0) + 1
            if kind == "文本框":
                text_chars += len(shape.text_frame.text.strip())
            position = f"({shape.left / 914400:.2f}, {shape.top / 914400:.2f})"
            lines.append(f"  - {describe(shape):<44} 位置 {position}")
        lines.append(f"  小计：{counts}")
        lines.append("")

    lines.append("-" * 78)
    lines.append(f"对象总计：{totals}")
    editable = sum(count for kind, count in totals.items() if kind != "图片")
    lines.append(
        f"非图片对象数：{editable}（这些对象均可在 PowerPoint 中独立选中并编辑）"
    )
    lines.append(f"可编辑文字字符数：{text_chars}")

    stats = {
        "slides": len(presentation.slides),
        "objects": totals,
        "editable_objects": editable,
        "text_chars": text_chars,
    }
    return lines, totals, stats


def main() -> None:
    parser = argparse.ArgumentParser(description="PPTX 可编辑性验证")
    parser.add_argument("pptx", nargs="+", help="PPTX 文件路径（可多个）")
    parser.add_argument("-o", "--output", help="把结果写入文件（默认打印到控制台）")
    args = parser.parse_args()

    output: list[str] = []
    for item in args.pptx:
        path = Path(item)
        if not path.exists():
            output.append(f"文件不存在：{path}")
            continue
        lines, _, _ = inspect(path)
        output.extend(lines)
        output.append("")

    text = "\n".join(output)
    if args.output:
        target = Path(args.output)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(text, encoding="utf-8")
        print(f"已写入 {target}")
    else:
        print(text)


if __name__ == "__main__":
    main()
