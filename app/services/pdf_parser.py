"""PDF 结构化解析。

把 PDF 页面拆解为「文本行 / 图片 / 矢量图形 / 表格」四类元素，并统一
使用 PyMuPDF 的归一化坐标（左上角原点、Y 轴向下、单位为点）。

坐标说明
--------
PDF 原始坐标系原点在左下角，PowerPoint 原点在左上角。PyMuPDF 已经把
页面坐标归一化到左上角原点，因此与 PPT 的换算只剩下单位换算：
``1 点 = 12700 EMU``。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf

EMU_PER_POINT = 12700

#: 矩形，格式为 (x0, y0, x1, y1)，单位：点
Rect = tuple[float, float, float, float]
#: 颜色，格式为 (r, g, b)，取值 0~1
Color = tuple[float, float, float]

_BOLD_TOKENS = ("bold", "black", "heavy", "semibold", "demibold", "ultrabold")
_ITALIC_TOKENS = ("italic", "oblique")


@dataclass
class TextSpan:
    """一段连续文本（同一行内样式一致的部分）。"""

    text: str
    font: str
    size: float
    color: int
    bold: bool
    italic: bool
    bbox: Rect
    #: 字符间距（点）；PDF 文本层一般没有该信息，OCR 还原时会估算
    spacing: float = 0.0


@dataclass
class TextLine:
    """一行文本，可包含多个 span。"""

    spans: list[TextSpan]
    bbox: Rect
    rotation: float = 0.0

    @property
    def text(self) -> str:
        return "".join(span.text for span in self.spans)

    @property
    def char_count(self) -> int:
        return sum(len(span.text) for span in self.spans)


@dataclass
class ImageElement:
    """页面中的一张位图。"""

    xref: int
    rect: Rect
    data: bytes
    ext: str
    pixel_width: int
    pixel_height: int
    alpha: bool = False


@dataclass
class ShapeElement:
    """矢量图形。

    ``kind`` 取值：

    * ``rect``：矩形，可直接还原为 PPT 矩形自选图形
    * ``line``：直线，可还原为 PPT 连接符
    * ``ellipse``：椭圆，可还原为 PPT 椭圆自选图形
    * ``complex``：复杂路径，无法结构化，保留为图片对象
    """

    kind: str
    rect: Rect
    color: Color | None = None
    fill: Color | None = None
    width: float = 1.0
    start: tuple[float, float] | None = None
    end: tuple[float, float] | None = None
    fill_opacity: float = 1.0
    stroke_opacity: float = 1.0
    dashed: bool = False


@dataclass
class TableElement:
    """表格。"""

    rect: Rect
    data: list[list[str]]
    row_bounds: list[Rect] = field(default_factory=list)
    col_bounds: list[Rect] = field(default_factory=list)
    border_color: tuple[int, int, int] = (128, 128, 128)
    border_width: float = 0.75
    font_name: str = ""
    font_size: float = 12.0
    font_color: int = 0
    cell_fills: list[list[tuple[int, int, int] | None]] = field(default_factory=list)


@dataclass
class PageModel:
    """一页 PDF 的结构化结果。"""

    index: int
    width: float
    height: float
    texts: list[TextLine] = field(default_factory=list)
    images: list[ImageElement] = field(default_factory=list)
    shapes: list[ShapeElement] = field(default_factory=list)
    tables: list[TableElement] = field(default_factory=list)
    art_rect: Rect | None = None
    #: 整页位图、无文本层，需要走 OCR 才能还原文字
    image_only: bool = False
    #: 覆盖整页的底图
    background: ImageElement | None = None
    #: 是否已经过 OCR 处理
    ocr_applied: bool = False


def _bbox(raw) -> Rect:
    return (float(raw[0]), float(raw[1]), float(raw[2]), float(raw[3]))


def _area(rect: Rect) -> float:
    return max(0.0, rect[2] - rect[0]) * max(0.0, rect[3] - rect[1])


def _center(rect: Rect) -> tuple[float, float]:
    return ((rect[0] + rect[2]) / 2.0, (rect[1] + rect[3]) / 2.0)


def _contains(outer: Rect, inner: Rect) -> bool:
    """判断 ``inner`` 是否完全落在 ``outer`` 内。"""
    return (
        inner[0] >= outer[0] - 0.5
        and inner[1] >= outer[1] - 0.5
        and inner[2] <= outer[2] + 0.5
        and inner[3] <= outer[3] + 0.5
    )


def _overlaps(a: Rect, b: Rect) -> bool:
    return not (a[2] <= b[0] or b[2] <= a[0] or a[3] <= b[1] or b[3] <= a[1])


def _union(rects: list[Rect]) -> Rect | None:
    if not rects:
        return None
    x0 = min(r[0] for r in rects)
    y0 = min(r[1] for r in rects)
    x1 = max(r[2] for r in rects)
    y1 = max(r[3] for r in rects)
    return (x0, y0, x1, y1)


def _detect_style(flags: int, font_name: str) -> tuple[bool, bool]:
    """从字体标志位与字体名推断粗体/斜体。"""
    lowered = (font_name or "").lower()
    bold = bool(flags & 16) or any(token in lowered for token in _BOLD_TOKENS)
    italic = bool(flags & 2) or any(token in lowered for token in _ITALIC_TOKENS)
    return bold, italic


def _is_axis_aligned_rect(items: list) -> bool:
    """判断若干直线段是否构成轴对齐矩形。"""
    points: list[tuple[float, float]] = []
    for item in items:
        if item[0] != "l":
            return False
        points.append((float(item[1].x), float(item[1].y)))
        points.append((float(item[2].x), float(item[2].y)))
    xs = {round(x, 2) for x, _ in points}
    ys = {round(y, 2) for _, y in points}
    if len(xs) != 2 or len(ys) != 2:
        return False
    corners = {(round(x, 2), round(y, 2)) for x, y in points}
    expected = {(x, y) for x in xs for y in ys}
    return corners == expected


class PdfParser:
    """把 PDF 解析为结构化页面模型。"""

    def __init__(self, source: bytes | str | Path, detect_tables: bool = True) -> None:
        if isinstance(source, (bytes, bytearray)):
            self.doc = pymupdf.open(stream=bytes(source), filetype="pdf")
        else:
            self.doc = pymupdf.open(str(source))
        self.detect_tables = detect_tables
        self._image_cache: dict[int, tuple[bytes, str, int, int, bool]] = {}

    # ------------------------------------------------------------------ 生命周期
    def close(self) -> None:
        try:
            self.doc.close()
        except Exception:  # pragma: no cover - 关闭失败不影响主流程
            pass

    def __enter__(self) -> "PdfParser":
        return self

    def __exit__(self, *_exc) -> None:
        self.close()

    @property
    def page_count(self) -> int:
        return self.doc.page_count

    # ------------------------------------------------------------------ 主流程
    def parse(self) -> list[PageModel]:
        """解析全部页面。"""
        return [self._parse_page(index) for index in range(self.doc.page_count)]

    def _parse_page(self, index: int) -> PageModel:
        page = self.doc[index]
        page_rect = page.rect
        model = PageModel(
            index=index,
            width=float(page_rect.width),
            height=float(page_rect.height),
        )

        shapes = self._parse_shapes(page)
        tables = self._parse_tables(page, shapes)
        table_rects = [table.rect for table in tables]

        model.shapes = self._filter_shapes_by_tables(shapes, table_rects, tables)
        model.tables = tables
        model.texts = self._parse_texts(page, table_rects)
        model.images = self._parse_images(page)

        complex_rects = [shape.rect for shape in model.shapes if shape.kind == "complex"]
        model.art_rect = _union(complex_rects)

        # 整页位图且没有文本层 → 需要 OCR 才能还原文字
        if not model.texts and model.images:
            page_rect = (0.0, 0.0, model.width, model.height)
            covering = [image for image in model.images if _contains(page_rect, image.rect)]
            if covering:
                model.image_only = True
                model.background = max(covering, key=lambda item: _area(item.rect))
        return model

    # ------------------------------------------------------------------ 文本
    def _parse_texts(self, page, excluded: list[Rect]) -> list[TextLine]:
        lines: list[TextLine] = []
        data = page.get_text("dict")
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                line_bbox = _bbox(line["bbox"])
                if excluded and any(_contains(rect, line_bbox) for rect in excluded):
                    continue
                spans: list[TextSpan] = []
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if not text.strip():
                        continue
                    bold, italic = _detect_style(int(span.get("flags", 0)), span.get("font", ""))
                    spans.append(
                        TextSpan(
                            text=text,
                            font=span.get("font", ""),
                            size=float(span.get("size", 12.0)),
                            color=int(span.get("color", 0)),
                            bold=bold,
                            italic=italic,
                            bbox=_bbox(span["bbox"]),
                        )
                    )
                if not spans:
                    continue
                rotation = 0.0
                direction = line.get("dir") or (1.0, 0.0)
                if abs(direction[0] - 1.0) > 1e-3 or abs(direction[1]) > 1e-3:
                    rotation = math.degrees(math.atan2(float(direction[1]), float(direction[0])))
                lines.append(TextLine(spans=spans, bbox=line_bbox, rotation=rotation))
        return lines

    # ------------------------------------------------------------------ 矢量图形
    def _parse_shapes(self, page) -> list[ShapeElement]:
        shapes: list[ShapeElement] = []
        for drawing in page.get_drawings():
            items = drawing.get("items", [])
            if not items:
                continue
            rect = _bbox(drawing["rect"])
            color = drawing.get("color")
            fill = drawing.get("fill")
            width = float(drawing.get("width") or 0.8)
            fill_opacity = float(drawing.get("fill_opacity", 1.0) or 1.0)
            stroke_opacity = float(drawing.get("stroke_opacity", 1.0) or 1.0)
            dashes = drawing.get("dashes") or ""
            dashed = bool(dashes) and not dashes.strip().startswith("[] 0")

            kind = self._classify_items(items)
            if kind == "line":
                start = items[0][1]
                end = items[0][2]
                shapes.append(
                    ShapeElement(
                        kind="line",
                        rect=rect,
                        color=color,
                        fill=None,
                        width=width,
                        start=(float(start.x), float(start.y)),
                        end=(float(end.x), float(end.y)),
                        stroke_opacity=stroke_opacity,
                        dashed=dashed,
                    )
                )
            elif kind == "rect":
                shapes.append(
                    ShapeElement(
                        kind="rect",
                        rect=rect,
                        color=color,
                        fill=fill,
                        width=width,
                        fill_opacity=fill_opacity,
                        stroke_opacity=stroke_opacity,
                        dashed=dashed,
                    )
                )
            elif kind == "ellipse":
                shapes.append(
                    ShapeElement(
                        kind="ellipse",
                        rect=rect,
                        color=color,
                        fill=fill,
                        width=width,
                        fill_opacity=fill_opacity,
                        stroke_opacity=stroke_opacity,
                    )
                )
            else:
                shapes.append(
                    ShapeElement(
                        kind="complex",
                        rect=rect,
                        color=color,
                        fill=fill,
                        width=width,
                        fill_opacity=fill_opacity,
                        stroke_opacity=stroke_opacity,
                    )
                )
        return shapes

    @staticmethod
    def _classify_items(items: list) -> str:
        kinds = {item[0] for item in items}
        if kinds == {"re"}:
            return "rect"
        if kinds == {"l"}:
            if len(items) == 1:
                return "line"
            return "rect" if _is_axis_aligned_rect(items) else "complex"
        if kinds == {"c"} and len(items) == 4:
            return "ellipse"
        if kinds <= {"re", "l"} and "re" in kinds and _is_axis_aligned_rect(
            [item for item in items if item[0] == "l"]
        ):
            return "rect"
        return "complex"

    def _filter_shapes_by_tables(
        self,
        shapes: list[ShapeElement],
        table_rects: list[Rect],
        tables: list[TableElement],
    ) -> list[ShapeElement]:
        """表格区域的线条由原生表格自带，这里剔除，避免重复绘制。"""
        if not table_rects:
            return shapes

        kept: list[ShapeElement] = []
        for shape in shapes:
            if any(_contains(rect, shape.rect) for rect in table_rects):
                continue
            kept.append(shape)

        # 汇总表格边框颜色
        for table in tables:
            colors: list[tuple[int, int, int]] = []
            widths: list[float] = []
            for shape in shapes:
                if not _contains(table.rect, shape.rect):
                    continue
                if shape.color:
                    colors.append(tuple(int(round(channel * 255)) for channel in shape.color))
                    widths.append(shape.width)
            if colors:
                table.border_color = max(set(colors), key=colors.count)
            if widths:
                table.border_width = max(0.5, min(widths))
        return kept

    # ------------------------------------------------------------------ 表格
    def _parse_tables(self, page, shapes: list[ShapeElement]) -> list[TableElement]:
        if not self.detect_tables:
            return []
        try:
            finder = page.find_tables()
        except Exception:
            return []

        page_area = float(page.rect.width * page.rect.height)
        tables: list[TableElement] = []
        for table in finder.tables:
            rect = _bbox(table.bbox)
            if _area(rect) <= 1:
                continue
            if page_area and _area(rect) > page_area * 0.9:
                # 几乎覆盖整页，多为误识别
                continue
            try:
                extracted = table.extract()
            except Exception:
                continue
            data = [
                ["" if cell is None else str(cell).strip() for cell in row]
                for row in extracted
            ]
            if len(data) < 2:
                continue
            col_count = max((len(row) for row in data), default=0)
            if col_count < 2:
                continue
            data = [row + [""] * (col_count - len(row)) for row in data]

            row_bounds: list[Rect] = []
            try:
                row_bounds = [_bbox(row.bbox) for row in table.rows]
            except Exception:
                row_bounds = []
            col_bounds: list[Rect] = []
            try:
                cells = table.rows[0].cells
                col_bounds = [_bbox(cell.bbox) for cell in cells]
            except Exception:
                col_bounds = []

            element = TableElement(
                rect=rect,
                data=data,
                row_bounds=row_bounds[: len(data)],
                col_bounds=col_bounds[:col_count],
            )
            element.cell_fills = self._detect_cell_fills(element, shapes)
            tables.append(element)
        return tables

    @staticmethod
    def _detect_cell_fills(
        table: TableElement, shapes: list[ShapeElement]
    ) -> list[list[tuple[int, int, int] | None]]:
        """用表格内部的填充矩形推断单元格底色。"""
        fills: list[tuple[Rect, tuple[int, int, int]]] = []
        for shape in shapes:
            if shape.kind != "rect" or not shape.fill:
                continue
            if not _contains(table.rect, shape.rect):
                continue
            color = tuple(int(round(channel * 255)) for channel in shape.fill)
            fills.append((shape.rect, color))  # type: ignore[arg-type]

        grid: list[list[tuple[int, int, int] | None]] = []
        for row in table.data:
            line: list[tuple[int, int, int] | None] = []
            for _ in row:
                line.append(None)
            grid.append(line)

        rows = len(table.data)
        cols = len(table.data[0]) if rows else 0
        row_bounds = table.row_bounds or []
        col_bounds = table.col_bounds or []
        for r in range(rows):
            for c in range(cols):
                cell_rect = PdfParser._cell_rect(row_bounds, col_bounds, r, c, table.rect)
                if cell_rect is None:
                    continue
                cx, cy = _center(cell_rect)
                for fill_rect, color in fills:
                    if (
                        fill_rect[0] <= cx <= fill_rect[2]
                        and fill_rect[1] <= cy <= fill_rect[3]
                    ):
                        grid[r][c] = color
                        break
        return grid

    @staticmethod
    def _cell_rect(
        row_bounds: list[Rect],
        col_bounds: list[Rect],
        row: int,
        col: int,
        table_rect: Rect,
    ) -> Rect | None:
        if row < len(row_bounds) and col < len(col_bounds):
            r0, r1 = row_bounds[row]
            c0, c1 = col_bounds[col]
            return (c0, r0, c1, r1)
        if len(col_bounds) == 2:
            rows = len(row_bounds)
            top = row_bounds[row][1] if row < rows else table_rect[1]
            bottom = row_bounds[row + 1][1] if row + 1 < rows else table_rect[3]
            return (col_bounds[0][0], top, col_bounds[1][2], bottom)
        return None

    # ------------------------------------------------------------------ 图片
    def _parse_images(self, page) -> list[ImageElement]:
        images: list[ImageElement] = []
        seen: set[tuple[int, int, int, int]] = set()
        for info in page.get_images(full=True):
            xref = int(info[0])
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                continue
            for rect in rects:
                bbox = _bbox(rect)
                if _area(bbox) <= 1:
                    continue
                key = (xref, int(bbox[0]), int(bbox[1]), int(bbox[2]))
                if key in seen:
                    continue
                seen.add(key)
                data, ext, width, height, alpha = self._extract_image(xref)
                if not data:
                    continue
                images.append(
                    ImageElement(
                        xref=xref,
                        rect=bbox,
                        data=data,
                        ext=ext,
                        pixel_width=width,
                        pixel_height=height,
                        alpha=alpha,
                    )
                )
        return images

    def _extract_image(self, xref: int) -> tuple[bytes, str, int, int, bool]:
        """提取图片二进制，必要时合成软掩码（alpha 通道）。"""
        if xref in self._image_cache:
            return self._image_cache[xref]
        try:
            info = self.doc.extract_image(xref)
        except Exception:
            info = None
        if not info or not info.get("image"):
            return b"", "png", 0, 0, False

        smask = int(info.get("smask") or 0)
        if smask:
            try:
                pixmap = pymupdf.Pixmap(self.doc, xref)
                mask = pymupdf.Pixmap(self.doc, smask)
                pixmap = pymupdf.Pixmap(pixmap, mask)
                data = pixmap.tobytes("png")
                result = (data, "png", pixmap.width, pixmap.height, True)
                self._image_cache[xref] = result
                return result
            except Exception:
                pass

        ext = info.get("ext", "png")
        data = info["image"]
        width = int(info.get("width", 0))
        height = int(info.get("height", 0))
        if ext in ("raw", "jpx", "jp2"):
            try:
                pixmap = pymupdf.Pixmap(self.doc, xref)
                data = pixmap.tobytes("png")
                ext = "png"
                width, height = pixmap.width, pixmap.height
            except Exception:
                pass
        result = (data, ext, width, height, False)
        self._image_cache[xref] = result
        return result
