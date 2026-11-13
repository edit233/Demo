"""PPTX 构建：把解析出的元素还原为 PowerPoint 原生对象。

还原策略
--------
* 文本 → 文本框（每行一个文本框，行内多样式使用多个 run）
* 图片 → 图片对象（独立可选中、移动、缩放）
* 矩形/直线/椭圆 → PPT 自选图形与连接符
* 表格 → PPT 原生表格
* 复杂矢量路径 → 以透明 PNG 作为独立图片对象保留
"""

from __future__ import annotations

import io
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu, Pt

from app.services.fonts import resolve_font
from app.services.pdf_parser import (
    EMU_PER_POINT,
    ImageElement,
    PageModel,
    ShapeElement,
    TableElement,
    TextLine,
)

#: PowerPoint 允许的最大页面尺寸（英寸）
MAX_SLIDE_INCHES = 56.0
#: 无样式、无网格的表格样式 GUID
PLAIN_TABLE_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"

_BORDER_TAGS = ("a:lnL", "a:lnR", "a:lnT", "a:lnB")
_FILL_TAGS = ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill")


def _emu(value: float) -> int:
    """点 → EMU。"""
    return int(round(value * EMU_PER_POINT))


def _hex(color: tuple[int, int, int]) -> str:
    return "{:02X}{:02X}{:02X}".format(*color)


def _rgb(color: tuple[int, int, int]) -> RGBColor:
    return RGBColor(*color)


def _scaled_color(color, opacity: float = 1.0) -> tuple[int, int, int]:
    """PDF 颜色（0~1）→ PPT 颜色（0~255）。"""
    return tuple(max(0, min(255, int(round(channel * 255)))) for channel in color)  # type: ignore[return-value]


def _set_ea_font(run, typeface: str) -> None:
    """设置东亚字体（``a:ea``），保证中文在 PowerPoint 中按指定字体渲染。"""
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    if latin is None:
        latin = rPr.get_or_add_latin()
    existing = rPr.find(qn("a:ea"))
    if existing is not None:
        existing.set("typeface", typeface)
        return
    element = parse_xml('<a:ea %s typeface="%s"/>' % (nsdecls("a"), typeface))
    latin.addnext(element)


def _apply_alpha(sp_pr, alpha: float) -> None:
    """给填充色 / 线条色附加透明度。"""
    if alpha >= 0.999:
        return
    value = max(0, min(100000, int(round(alpha * 100000))))
    for parent in (sp_pr, sp_pr.find(qn("a:ln"))):
        if parent is None:
            continue
        solid = parent.find(qn("a:solidFill"))
        if solid is None:
            continue
        color = solid.find(qn("a:srgbClr"))
        if color is None:
            continue
        alpha_el = color.find(qn("a:alpha"))
        if alpha_el is None:
            alpha_el = parse_xml('<a:alpha %s val="%d"/>' % (nsdecls("a"), value))
            color.append(alpha_el)
        else:
            alpha_el.set("val", str(value))


def _set_cell_border(cell, color: tuple[int, int, int], width_pt: float) -> None:
    """给表格单元格四周加边框（python-pptx 未提供该 API）。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    for tag in _BORDER_TAGS:
        existing = tc_pr.find(qn(tag))
        if existing is not None:
            tc_pr.remove(existing)
    width_emu = max(1, int(round(width_pt * EMU_PER_POINT)))
    for tag in reversed(_BORDER_TAGS):
        xml = (
            '<{tag} {ns} w="{width}" cap="flat" cmpd="sng" algn="ctr">'
            '<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            '<a:prstDash val="solid"/>'
            "</{tag}>"
        ).format(tag=tag, ns=nsdecls("a"), width=width_emu, color=_hex(color))
        tc_pr.insert(0, parse_xml(xml))


def _set_cell_fill(cell, color: tuple[int, int, int] | None) -> None:
    """设置单元格底色。"""
    tc_pr = cell._tc.get_or_add_tcPr()
    for tag in _FILL_TAGS:
        existing = tc_pr.find(qn(tag))
        if existing is not None:
            tc_pr.remove(existing)
    if color is None:
        return
    fill = parse_xml(
        '<a:solidFill %s><a:srgbClr val="%s"/></a:solidFill>' % (nsdecls("a"), _hex(color))
    )
    index = 0
    for tag in _BORDER_TAGS:
        if tc_pr.find(qn(tag)) is not None:
            index += 1
    tc_pr.insert(index, fill)


def _set_plain_table_style(table) -> None:
    """使用「无样式，无网格」表格样式，避免主题默认蓝色条纹。"""
    tbl = table._tbl
    tbl_pr = tbl.find(qn("a:tblPr"))
    if tbl_pr is None:
        tbl_pr = parse_xml("<a:tblPr %s/>" % nsdecls("a"))
        tbl.insert(0, tbl_pr)
    tbl_pr.set("firstRow", "0")
    tbl_pr.set("bandRow", "0")
    style_id = tbl_pr.find(qn("a:tableStyleId"))
    if style_id is None:
        style_id = parse_xml("<a:tableStyleId %s/>" % nsdecls("a"))
        tbl_pr.append(style_id)
    style_id.text = PLAIN_TABLE_STYLE


@dataclass
class _Transform:
    """PDF 点坐标 → PPT EMU 坐标的换算。"""

    scale: float = 1.0
    offset_x: float = 0.0
    offset_y: float = 0.0

    def x(self, value: float) -> float:
        return self.offset_x + value * self.scale

    def y(self, value: float) -> float:
        return self.offset_y + value * self.scale

    def length(self, value: float) -> float:
        return value * self.scale


class PptxBuilder:
    """把页面模型写入 PPTX。"""

    def __init__(self, keep_art_layer: bool = True) -> None:
        self.keep_art_layer = keep_art_layer

    # ------------------------------------------------------------------ 入口
    def build(
        self,
        pages: list[PageModel],
        art_images: dict[int, bytes],
        output_path: Path | str,
    ) -> None:
        """生成 PPTX 文件。"""
        if not pages:
            raise ValueError("没有可转换的页面")

        presentation = Presentation()
        presentation.slide_width = Emu(self._slide_size(pages[0].width))
        presentation.slide_height = Emu(self._slide_size(pages[0].height))
        blank_layout = presentation.slide_layouts[6]

        for model in pages:
            slide = presentation.slides.add_slide(blank_layout)
            self._fill_slide(
                slide,
                model,
                art_images.get(model.index),
                presentation.slide_width,
                presentation.slide_height,
            )

        core = presentation.core_properties
        core.title = "PDF 转可编辑 PPTX"
        core.author = "pdf2pptx-service"
        core.comments = "由 pdf2pptx-service 自动转换生成"

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        presentation.save(str(output_path))

    @staticmethod
    def _slide_size(points: float) -> int:
        """把 PDF 页面尺寸换算为 PPT 页面尺寸，并限制在 PowerPoint 允许范围内。"""
        size = _emu(points)
        limit = int(MAX_SLIDE_INCHES * 914400)
        return max(size, 1) if size <= limit else limit

    def _transform(self, model: PageModel, slide_width: int, slide_height: int) -> _Transform:
        page_width = _emu(model.width)
        page_height = _emu(model.height)
        scale = min(slide_width / page_width, slide_height / page_height)
        offset_x = (slide_width - page_width * scale) / 2.0
        offset_y = (slide_height - page_height * scale) / 2.0
        return _Transform(scale=scale, offset_x=offset_x, offset_y=offset_y)

    # ------------------------------------------------------------------ 页面
    def _fill_slide(
        self,
        slide,
        model: PageModel,
        art_png: bytes | None,
        slide_width: int,
        slide_height: int,
    ) -> None:
        transform = self._transform(model, slide_width, slide_height)

        page_area = model.width * model.height
        backgrounds = [
            shape
            for shape in model.shapes
            if shape.kind == "rect" and shape.fill and _shape_area(shape) >= page_area * 0.5
        ]
        others = [shape for shape in model.shapes if shape not in backgrounds]

        # 1) 整页底色（置于最底层）
        for shape in backgrounds:
            self._add_shape(slide, shape, transform)
        # 2) 复杂矢量层（透明 PNG）
        if self.keep_art_layer and art_png and model.art_rect:
            self._add_art_layer(slide, model.art_rect, art_png, transform)
        # 3) 其它矢量图形
        for shape in others:
            self._add_shape(slide, shape, transform)
        # 4) 图片
        for image in model.images:
            self._add_picture(slide, image, transform)
        # 5) 表格
        for table in model.tables:
            self._add_table(slide, table, transform)
        # 6) 文本
        for line in model.texts:
            self._add_text(slide, line, transform)

    # ------------------------------------------------------------------ 矢量
    def _add_shape(self, slide, shape: ShapeElement, transform: _Transform) -> None:
        if shape.kind == "line" and shape.start and shape.end:
            x1 = int(transform.x(_emu(shape.start[0])))
            y1 = int(transform.y(_emu(shape.start[1])))
            x2 = int(transform.x(_emu(shape.end[0])))
            y2 = int(transform.y(_emu(shape.end[1])))
            connector = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT, Emu(x1), Emu(y1), Emu(x2), Emu(y2)
            )
            if shape.color:
                connector.line.color.rgb = _rgb(_scaled_color(shape.color))
                connector.line.width = Pt(max(0.25, shape.width * transform.scale))
            else:
                connector.line.fill.background()
            if shape.dashed:
                try:
                    connector.line.dash_style = MSO_LINE_DASH_STYLE.DASH
                except Exception:
                    pass
            _apply_alpha(connector._element.spPr, shape.stroke_opacity)
            _disable_shadow(connector)
            return

        left = int(transform.x(_emu(shape.rect[0])))
        top = int(transform.y(_emu(shape.rect[1])))
        width = max(1, int(transform.length(_emu(shape.rect[2] - shape.rect[0]))))
        height = max(1, int(transform.length(_emu(shape.rect[3] - shape.rect[1]))))

        if shape.kind == "ellipse":
            auto_shape = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Emu(left), Emu(top), Emu(width), Emu(height)
            )
        else:
            auto_shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Emu(left), Emu(top), Emu(width), Emu(height)
            )

        if shape.fill:
            auto_shape.fill.solid()
            auto_shape.fill.fore_color.rgb = _rgb(_scaled_color(shape.fill))
        else:
            auto_shape.fill.background()

        if shape.color:
            auto_shape.line.color.rgb = _rgb(_scaled_color(shape.color))
            auto_shape.line.width = Pt(max(0.25, shape.width * transform.scale))
        else:
            auto_shape.line.fill.background()

        if shape.dashed:
            try:
                auto_shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            except Exception:
                pass

        _apply_alpha(auto_shape._element.spPr, shape.fill_opacity)
        _disable_shadow(auto_shape)

    # ------------------------------------------------------------------ 图片
    def _add_picture(self, slide, image: ImageElement, transform: _Transform) -> None:
        left = int(transform.x(_emu(image.rect[0])))
        top = int(transform.y(_emu(image.rect[1])))
        width = max(1, int(transform.length(_emu(image.rect[2] - image.rect[0]))))
        height = max(1, int(transform.length(_emu(image.rect[3] - image.rect[1]))))
        try:
            picture = slide.shapes.add_picture(
                io.BytesIO(image.data), Emu(left), Emu(top), Emu(width), Emu(height)
            )
        except Exception:
            return
        _apply_crop(picture, image.pixel_width, image.pixel_height, width, height)
        _disable_shadow(picture)

    def _add_art_layer(
        self,
        slide,
        rect: tuple[float, float, float, float],
        png: bytes,
        transform: _Transform,
    ) -> None:
        left = int(transform.x(_emu(rect[0])))
        top = int(transform.y(_emu(rect[1])))
        width = max(1, int(transform.length(_emu(rect[2] - rect[0]))))
        height = max(1, int(transform.length(_emu(rect[3] - rect[1]))))
        try:
            picture = slide.shapes.add_picture(
                io.BytesIO(png), Emu(left), Emu(top), Emu(width), Emu(height)
            )
        except Exception:
            return
        pixel_width, pixel_height = _png_size(png)
        if pixel_width and pixel_height:
            _apply_crop(picture, pixel_width, pixel_height, width, height)
        _disable_shadow(picture)

    # ------------------------------------------------------------------ 表格
    def _add_table(self, slide, table: TableElement, transform: _Transform) -> None:
        rows = len(table.data)
        cols = len(table.data[0]) if rows else 0
        if rows == 0 or cols == 0:
            return

        left = int(transform.x(_emu(table.rect[0])))
        top = int(transform.y(_emu(table.rect[1])))
        width = max(1, int(transform.length(_emu(table.rect[2] - table.rect[0]))))
        height = max(1, int(transform.length(_emu(table.rect[3] - table.rect[1]))))

        graphic_frame = slide.shapes.add_table(
            rows, cols, Emu(left), Emu(top), Emu(width), Emu(height)
        )
        pptx_table = graphic_frame.table
        _set_plain_table_style(pptx_table)

        if table.col_bounds and len(table.col_bounds) == cols:
            total = sum(b[2] - b[0] for b in table.col_bounds) or 1.0
            for index, bound in enumerate(table.col_bounds):
                pptx_table.columns[index].width = Emu(
                    max(1, int(width * (bound[2] - bound[0]) / total))
                )
        if table.row_bounds and len(table.row_bounds) == rows:
            total = sum(b[3] - b[1] for b in table.row_bounds) or 1.0
            for index, bound in enumerate(table.row_bounds):
                pptx_table.rows[index].height = Emu(
                    max(1, int(height * (bound[3] - bound[1]) / total))
                )

        font_size = max(6.0, table.font_size * transform.scale)
        for r, row in enumerate(table.data):
            for c, value in enumerate(row):
                cell = pptx_table.cell(r, c)
                cell.margin_left = Emu(45720)
                cell.margin_right = Emu(45720)
                cell.margin_top = Emu(18288)
                cell.margin_bottom = Emu(18288)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                _set_cell_border(cell, table.border_color, table.border_width)
                fill = None
                if r < len(table.cell_fills) and c < len(table.cell_fills[r]):
                    fill = table.cell_fills[r][c]
                _set_cell_fill(cell, fill)

                text_frame = cell.text_frame
                text_frame.word_wrap = True
                lines = str(value).split("\n")
                for index, line_text in enumerate(lines):
                    paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
                    paragraph.alignment = PP_ALIGN.CENTER
                    paragraph.space_before = Pt(0)
                    paragraph.space_after = Pt(0)
                    run = paragraph.add_run()
                    run.text = line_text
                    font = run.font
                    font.size = Pt(font_size)
                    font.bold = False
                    font.italic = False
                    latin, east_asian = resolve_font(table.font_name, line_text)
                    font.name = latin
                    _set_ea_font(run, east_asian)
                    font.color.rgb = _rgb(
                        (
                            (table.font_color >> 16) & 0xFF,
                            (table.font_color >> 8) & 0xFF,
                            table.font_color & 0xFF,
                        )
                    )

    # ------------------------------------------------------------------ 文本
    def _add_text(self, slide, line: TextLine, transform: _Transform) -> None:
        x0, y0, x1, y1 = line.bbox
        width_pt = x1 - x0
        height_pt = y1 - y0
        rotation = line.rotation

        if abs(rotation) > 1.0:
            # 旋转文本：先按未旋转尺寸建框，再整体旋转
            width_pt, height_pt = height_pt, width_pt

        left = int(transform.x(_emu(x0)))
        top = int(transform.y(_emu(y0)))
        width = max(1, int(transform.length(_emu(width_pt))))
        height = max(1, int(transform.length(_emu(height_pt))))

        box = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
        if abs(rotation) > 1.0:
            box.rotation = rotation

        text_frame = box.text_frame
        text_frame.word_wrap = False
        text_frame.auto_size = MSO_AUTO_SIZE.NONE
        text_frame.margin_left = 0
        text_frame.margin_right = 0
        text_frame.margin_top = 0
        text_frame.margin_bottom = 0
        text_frame.vertical_anchor = MSO_ANCHOR.TOP

        paragraph = text_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.LEFT
        paragraph.space_before = Pt(0)
        paragraph.space_after = Pt(0)

        for span in line.spans:
            run = paragraph.add_run()
            run.text = span.text
            font = run.font
            font.size = Pt(max(1.0, span.size * transform.scale))
            font.bold = span.bold
            font.italic = span.italic
            latin, east_asian = resolve_font(span.font, span.text)
            font.name = latin
            _set_ea_font(run, east_asian)
            font.color.rgb = _rgb(
                (
                    (span.color >> 16) & 0xFF,
                    (span.color >> 8) & 0xFF,
                    span.color & 0xFF,
                )
            )
            # 字符间距（``a:rPr/@spc``，单位 1/100 磅）
            if span.spacing:
                run._r.get_or_add_rPr().set(
                    "spc", str(int(round(span.spacing * transform.scale * 100)))
                )


def _shape_area(shape: ShapeElement) -> float:
    return max(0.0, shape.rect[2] - shape.rect[0]) * max(0.0, shape.rect[3] - shape.rect[1])


def _disable_shadow(shape) -> None:
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def _apply_crop(picture, pixel_width: int, pixel_height: int, width: int, height: int) -> None:
    """按目标框比例裁剪图片，避免拉伸变形。"""
    if not pixel_width or not pixel_height or width <= 0 or height <= 0:
        return
    image_ratio = pixel_width / pixel_height
    box_ratio = width / height
    try:
        if image_ratio > box_ratio:
            crop = (1 - box_ratio / image_ratio) / 2
            picture.crop_left = crop
            picture.crop_right = crop
        elif image_ratio < box_ratio:
            crop = (1 - image_ratio / box_ratio) / 2
            picture.crop_top = crop
            picture.crop_bottom = crop
    except Exception:
        pass


def _png_size(data: bytes) -> tuple[int, int]:
    """读取 PNG 像素尺寸。"""
    try:
        from PIL import Image

        with Image.open(io.BytesIO(data)) as image:
            return image.size
    except Exception:
        return 0, 0
