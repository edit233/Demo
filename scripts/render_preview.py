"""把生成的 PPTX 渲染成 PNG，用于人工比对还原效果。

本脚本用 PyMuPDF 近似重绘 PPTX 中的对象（文本框 / 图片 / 形状 / 连接符 /
表格），输出每一页的 PNG 图片。它是「可编辑性验证材料」的一部分：既能
看到页面长什么样，也能证明页面由多个独立对象组成，而不是一张整页图片。

用法::

    python scripts/render_preview.py outputs/<file_id>/sample.pptx -o preview
"""

from __future__ import annotations

import argparse
import io
import sys
from pathlib import Path

import pymupdf
from PIL import Image
from pptx import Presentation
from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE

EMU_PER_POINT = 12700
DEFAULT_FONT = "china-s"
LATIN_FONT = "helv"
RPR_TAG = "{http://schemas.openxmlformats.org/drawingml/2006/main}rPr"


def _pick_font(text: str) -> str:
    """纯西文用 Helvetica 度量，含中文时用 china-s（内置简体中文字体）。"""
    if text.isascii():
        return LATIN_FONT
    return DEFAULT_FONT


def _char_spacing(shape) -> float:
    """读取首个 run 的字符间距（磅）。"""
    try:
        run = shape.text_frame.paragraphs[0].runs[0]
        rPr = run._r.find(RPR_TAG)
        if rPr is not None and rPr.get("spc"):
            return int(rPr.get("spc")) / 100.0
    except Exception:
        pass
    return 0.0


def _pt(value: int) -> float:
    """EMU → 点。"""
    return value / EMU_PER_POINT


def _rgb(color) -> tuple[float, float, float]:
    value = tuple(color)
    return (value[0] / 255.0, value[1] / 255.0, value[2] / 255.0)


def _fill_color(shape) -> tuple[float, float, float] | None:
    try:
        if shape.fill.type == MSO_FILL.SOLID:
            return _rgb(shape.fill.fore_color.rgb)
    except Exception:
        pass
    return None


def _line_color(shape) -> tuple[float, float, float] | None:
    try:
        if shape.line.fill.type == MSO_FILL.BACKGROUND:
            return None
        return _rgb(shape.line.color.rgb)
    except Exception:
        return None


def _line_width(shape) -> float:
    try:
        width = shape.line.width
        return _pt(width) if width else 1.0
    except Exception:
        return 1.0


def _draw_picture(page, shape) -> None:
    rect = pymupdf.Rect(
        _pt(shape.left), _pt(shape.top), _pt(shape.left + shape.width), _pt(shape.top + shape.height)
    )
    try:
        blob = shape.image.blob
    except Exception:
        return
    try:
        crops = (shape.crop_left, shape.crop_top, shape.crop_right, shape.crop_bottom)
        if any(crops):
            with Image.open(io.BytesIO(blob)) as image:
                width, height = image.size
                box = (
                    int(width * crops[0]),
                    int(height * crops[1]),
                    int(width * (1 - crops[2])),
                    int(height * (1 - crops[3])),
                )
                buffer = io.BytesIO()
                image.crop(box).save(buffer, format="PNG")
                blob = buffer.getvalue()
        page.insert_image(rect, stream=blob)
    except Exception:
        pass


def _draw_shape(page, shape) -> None:
    rect = pymupdf.Rect(
        _pt(shape.left), _pt(shape.top), _pt(shape.left + shape.width), _pt(shape.top + shape.height)
    )
    fill = _fill_color(shape)
    stroke = _line_color(shape)
    width = _line_width(shape)
    is_oval = False
    try:
        is_oval = shape.auto_shape_type == MSO_SHAPE.OVAL
    except Exception:
        pass
    if is_oval:
        page.draw_circle(
            ((rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2),
            max(abs(rect.x1 - rect.x0), abs(rect.y1 - rect.y0)) / 2,
            color=stroke,
            fill=fill,
            width=width,
        )
    else:
        page.draw_rect(rect, color=stroke, fill=fill, width=width)


def _draw_connector(page, shape) -> None:
    try:
        start = (shape.begin_x, shape.begin_y)
        end = (shape.end_x, shape.end_y)
    except Exception:
        return
    page.draw_line(
        pymupdf.Point(_pt(start[0]), _pt(start[1])),
        pymupdf.Point(_pt(end[0]), _pt(end[1])),
        color=_line_color(shape) or (0, 0, 0),
        width=_line_width(shape),
    )


def _draw_text(page, shape) -> None:
    """绘制文本框内容。

    PPTX 中的文本框是 ``word_wrap=False``（不换行），因此这里使用
    ``insert_text`` 单行绘制，而不是会自动折行的 ``insert_textbox``，
    否则预览图会出现文字换行、与相邻文本框重叠的假象。
    """
    text_frame = shape.text_frame
    text = text_frame.text
    if not text.strip():
        return
    rect = pymupdf.Rect(
        _pt(shape.left), _pt(shape.top), _pt(shape.left + shape.width), _pt(shape.top + shape.height)
    )
    size = 12.0
    color = (0, 0, 0)
    try:
        paragraph = text_frame.paragraphs[0]
        if paragraph.runs:
            font = paragraph.runs[0].font
            if font.size:
                size = font.size.pt
            if font.color and font.color.type is not None:
                color = _rgb(font.color.rgb)
    except Exception:
        pass

    fontname = _pick_font(text)
    spacing = _char_spacing(shape)
    font = pymupdf.Font(fontname)
    # 估算文字宽度，若超出文本框则等比缩小字号（与原页面观感保持一致）
    try:
        measured = font.text_length(text, fontsize=size)
    except Exception:
        measured = 0.0
    if spacing:
        measured += spacing * max(0, len(text) - 1)
    box_width = rect.x1 - rect.x0
    if measured > box_width > 0:
        size = max(6.0, size * box_width / measured)

    baseline = rect.y0 + size * font.ascender
    try:
        if spacing > 0.01:
            # 逐字符绘制以体现字符间距
            x = rect.x0
            for char in text:
                if char == "\n":
                    continue
                page.insert_text(
                    pymupdf.Point(x, baseline),
                    char,
                    fontsize=size,
                    fontname=fontname,
                    color=color,
                )
                x += font.text_length(char, fontsize=size) + spacing
        else:
            page.insert_text(
                pymupdf.Point(rect.x0, baseline),
                text,
                fontsize=size,
                fontname=fontname,
                color=color,
            )
    except Exception:
        pass


def _draw_table(page, shape) -> None:
    table = shape.table
    left = _pt(shape.left)
    top = _pt(shape.top)
    row_heights = [_pt(row.height) for row in table.rows]
    col_widths = [_pt(column.width) for column in table.columns]

    y = top
    for r, row in enumerate(table.rows):
        x = left
        for c in range(len(col_widths)):
            cell_rect = pymupdf.Rect(x, y, x + col_widths[c], y + row_heights[r])
            page.draw_rect(cell_rect, color=(0.75, 0.78, 0.82), width=0.6)
            cell = table.cell(r, c)
            text = cell.text
            if text.strip():
                try:
                    paragraph = cell.text_frame.paragraphs[0]
                    font_size = 12.0
                    if paragraph.runs and paragraph.runs[0].font.size:
                        font_size = paragraph.runs[0].font.size.pt
                except Exception:
                    pass
                page.insert_textbox(
                    pymupdf.Rect(cell_rect.x0 + 3, cell_rect.y0 + 2, cell_rect.x1 - 3, cell_rect.y1 - 2),
                    text,
                    fontsize=font_size,
                    fontname=DEFAULT_FONT,
                    align=1,
                )
            x += col_widths[c]
        y += row_heights[r]


def render_slide(slide, width_pt: float, height_pt: float) -> pymupdf.Document:
    """把一页幻灯片近似渲染为 PDF 页。"""
    document = pymupdf.open()
    page = document.new_page(width=width_pt, height=height_pt)
    for shape in slide.shapes:
        try:
            if shape.has_table:
                _draw_table(page, shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                _draw_picture(page, shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.LINE:
                _draw_connector(page, shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                _draw_shape(page, shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                _draw_text(page, shape)
            elif shape.has_text_frame:
                _draw_text(page, shape)
        except Exception:
            continue
    return document


def render_pptx(pptx_path: Path, output_dir: Path, zoom: float = 1.5) -> list[Path]:
    """渲染 PPTX 的每一页，输出 PNG。"""
    presentation = Presentation(str(pptx_path))
    output_dir.mkdir(parents=True, exist_ok=True)
    width_pt = _pt(presentation.slide_width)
    height_pt = _pt(presentation.slide_height)

    outputs: list[Path] = []
    for index, slide in enumerate(presentation.slides, start=1):
        document = render_slide(slide, width_pt, height_pt)
        pixmap = document[0].get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        target = output_dir / f"slide_{index:02d}.png"
        pixmap.save(str(target))
        document.close()
        outputs.append(target)
        print(f"已渲染 {target}")
    return outputs


def main() -> None:
    parser = argparse.ArgumentParser(description="把 PPTX 渲染为 PNG 预览")
    parser.add_argument("pptx", help="PPTX 文件路径")
    parser.add_argument("-o", "--output", default="preview", help="输出目录")
    parser.add_argument("-z", "--zoom", type=float, default=1.5, help="缩放倍数")
    args = parser.parse_args()

    path = Path(args.pptx)
    if not path.exists():
        print(f"文件不存在：{path}")
        raise SystemExit(1)
    render_pptx(path, Path(args.output), args.zoom)


if __name__ == "__main__":
    main()