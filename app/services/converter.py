"""转换流程编排：解析 → 渲染复杂矢量层 → 构建 PPTX → 生成校验报告。"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf

from app.core.config import settings
from app.services.ocr import OcrEngine, erase_regions, pixel_regions
from app.services.pdf_parser import ImageElement, PageModel, PdfParser, Rect, TextLine, TextSpan
from app.services.pptx_builder import PptxBuilder


class ConversionError(Exception):
    """转换失败，消息可直接展示给用户。"""


@dataclass
class ConvertOptions:
    """转换选项。"""

    keep_art_layer: bool = True
    detect_tables: bool = True
    render_scale: float = 2.0
    enable_ocr: bool = True
    ocr_min_confidence: float = 0.5
    ocr_erase_text: bool = True


@dataclass
class ConversionResult:
    """转换产物。"""

    output_path: Path
    report_path: Path
    report: dict = field(default_factory=dict)


def _to_rect(bbox: Rect, pad: float = 0.0) -> pymupdf.Rect:
    """转成 PyMuPDF 矩形，并保证宽高不为 0。"""
    x0, y0, x1, y1 = bbox
    if x1 - x0 < pad:
        x0 -= pad / 2
        x1 += pad / 2
    if y1 - y0 < pad:
        y0 -= pad / 2
        y1 += pad / 2
    return pymupdf.Rect(x0, y0, x1, y1)


class ArtRenderer:
    """把无法结构化还原的复杂矢量内容渲染为透明 PNG。

    做法：先在 PDF 副本上「涂掉」已经用原生对象还原的文本、图片与简单
    矢量图形，再只渲染剩余的复杂路径区域。这样输出的 PPTX 既保留了
    原始视觉效果，又不会把文字重复画在图片里。
    """

    def __init__(self, pdf_bytes: bytes, scale: float = 2.0) -> None:
        self.doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        self.scale = max(0.5, min(4.0, scale))

    def close(self) -> None:
        try:
            self.doc.close()
        except Exception:  # pragma: no cover
            pass

    def render(self, model: PageModel) -> bytes | None:
        """渲染该页的复杂矢量层，返回 PNG 二进制。"""
        if not model.art_rect:
            return None
        page = self.doc[model.index]
        try:
            self._redact(page, model)
            clip = _to_rect(model.art_rect, pad=0.5)
            matrix = pymupdf.Matrix(self.scale, self.scale)
            pixmap = page.get_pixmap(clip=clip, matrix=matrix, alpha=True, annots=False)
            return pixmap.tobytes("png")
        except Exception:
            return None

    @staticmethod
    def _redact(page, model: PageModel) -> None:
        # 第一轮：移除文本与图片（它们会以原生对象重新加入 PPTX）
        for line in model.texts:
            page.add_redact_annot(_to_rect(line.bbox, pad=0.2))
        for image in model.images:
            page.add_redact_annot(_to_rect(image.rect, pad=0.2))
        page.apply_redactions(
            images=pymupdf.PDF_REDACT_IMAGE_REMOVE,
            text=pymupdf.PDF_REDACT_TEXT_REMOVE,
            graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
        )

        # 第二轮：移除已还原为原生形状的矢量，保留复杂路径
        for shape in model.shapes:
            if shape.kind in {"rect", "line", "ellipse"}:
                page.add_redact_annot(_to_rect(shape.rect, pad=0.2))
        page.apply_redactions(
            images=pymupdf.PDF_REDACT_IMAGE_NONE,
            text=pymupdf.PDF_REDACT_TEXT_NONE,
            graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
        )


def _color_to_int(color: tuple[int, int, int]) -> int:
    return (color[0] << 16) | (color[1] << 8) | color[2]


def apply_ocr(model: PageModel, options: ConvertOptions) -> int:
    """对「整页位图」页面执行 OCR，把文字还原为可编辑文本框。

    返回识别出的文字块数量。
    """
    if not options.enable_ocr or not model.image_only or model.background is None:
        return 0
    engine = OcrEngine.instance()
    if not engine.available():
        return 0

    background = model.background
    texts = engine.recognize(
        background.data, background.rect, min_confidence=options.ocr_min_confidence
    )
    if not texts:
        return 0

    # 1) 把文字从底图中抹掉，避免与叠加文字重影
    if options.ocr_erase_text:
        size = (background.pixel_width, background.pixel_height)
        regions = pixel_regions(texts, background.rect, size)
        try:
            cleaned = erase_regions(background.data, regions)
            model.images = [
                ImageElement(
                    xref=background.xref,
                    rect=background.rect,
                    data=cleaned,
                    ext="jpeg",
                    pixel_width=background.pixel_width,
                    pixel_height=background.pixel_height,
                    alpha=False,
                )
            ]
            model.background = model.images[0]
        except Exception:
            pass

    # 2) 文字转为可编辑文本框
    lines: list[TextLine] = []
    for item in texts:
        box = item.layout_rect()
        lines.append(
            TextLine(
                spans=[
                    TextSpan(
                        text=item.text,
                        font="",
                        size=item.size,
                        color=_color_to_int(item.color),
                        bold=False,
                        italic=False,
                        bbox=box,
                        spacing=item.spacing,
                    )
                ],
                bbox=box,
            )
        )
    model.texts = lines
    model.ocr_applied = True
    return len(lines)


def _element_report(model: PageModel, has_art: bool) -> list[dict]:
    """列出该页所有被还原出来的对象，用于可编辑性验证。"""
    elements: list[dict] = []
    for line in model.texts:
        first = line.spans[0]
        elements.append(
            {
                "type": "text",
                "text": line.text[:120],
                "bbox": [round(v, 2) for v in line.bbox],
                "font": first.font,
                "size": round(first.size, 2),
                "bold": first.bold,
                "italic": first.italic,
                "runs": len(line.spans),
            }
        )
    for image in model.images:
        elements.append(
            {
                "type": "image",
                "bbox": [round(v, 2) for v in image.rect],
                "pixel": [image.pixel_width, image.pixel_height],
                "format": image.ext,
                "alpha": image.alpha,
            }
        )
    for shape in model.shapes:
        elements.append(
            {
                "type": "shape",
                "kind": shape.kind,
                "bbox": [round(v, 2) for v in shape.rect],
                "fill": _color_hex(shape.fill),
                "stroke": _color_hex(shape.color),
            }
        )
    for table in model.tables:
        elements.append(
            {
                "type": "table",
                "bbox": [round(v, 2) for v in table.rect],
                "rows": len(table.data),
                "cols": len(table.data[0]) if table.data else 0,
            }
        )
    if has_art and model.art_rect:
        elements.append(
            {
                "type": "image",
                "kind": "complex_vector_layer",
                "bbox": [round(v, 2) for v in model.art_rect],
                "note": "复杂矢量路径，保留为独立图片对象",
            }
        )
    return elements


def _color_hex(color) -> str | None:
    if not color:
        return None
    return "#{:02X}{:02X}{:02X}".format(
        *[max(0, min(255, int(round(channel * 255)))) for channel in color]
    )


def build_report(pages: list[PageModel], art_pages: set[int], source_name: str) -> dict:
    """生成可编辑性验证报告。"""
    page_reports = []
    totals = {"text_lines": 0, "text_chars": 0, "images": 0, "shapes": 0, "tables": 0, "art_layers": 0}

    for model in pages:
        has_art = model.index in art_pages
        stats = {
            "text_lines": len(model.texts),
            "text_chars": sum(line.char_count for line in model.texts),
            "images": len(model.images),
            "shapes": len(model.shapes),
            "tables": len(model.tables),
            "art_layers": 1 if has_art else 0,
        }
        for key, value in stats.items():
            totals[key] += value
        page_reports.append(
            {
                "index": model.index + 1,
                "width": round(model.width, 2),
                "height": round(model.height, 2),
                "ocr": model.ocr_applied,
                **stats,
                "elements": _element_report(model, has_art),
            }
        )

    editable = totals["text_lines"] + totals["images"] + totals["shapes"] + totals["tables"]
    total = editable + totals["art_layers"]
    return {
        "source_name": source_name,
        "page_count": len(pages),
        "summary": {
            **totals,
            "editable_objects": editable,
            "editable_ratio": round(editable / total, 4) if total else 0.0,
        },
        "pages": page_reports,
    }


def convert_pdf(
    pdf_path: Path,
    output_dir: Path,
    source_name: str,
    options: ConvertOptions | None = None,
) -> ConversionResult:
    """把单个 PDF 转换为 PPTX。"""
    options = options or ConvertOptions()
    pdf_path = Path(pdf_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_bytes = pdf_path.read_bytes()
    if not pdf_bytes:
        raise ConversionError("文件内容为空")

    parser = PdfParser(pdf_bytes, detect_tables=options.detect_tables)
    try:
        if parser.page_count == 0:
            raise ConversionError("PDF 中没有可转换的页面")
        if parser.page_count > settings.max_pages:
            raise ConversionError(f"页数超过上限（最多 {settings.max_pages} 页）")
        pages = parser.parse()
    except ConversionError:
        raise
    except Exception as exc:  # pragma: no cover - 损坏文件
        raise ConversionError(f"PDF 解析失败：{exc}") from exc
    finally:
        parser.close()

    art_images: dict[int, bytes] = {}
    if options.keep_art_layer:
        renderer = ArtRenderer(pdf_bytes, scale=options.render_scale)
        try:
            for model in pages:
                if not model.art_rect:
                    continue
                png = renderer.render(model)
                if png:
                    art_images[model.index] = png
        finally:
            renderer.close()

    # 图片型页面 → OCR 还原文字
    for model in pages:
        if model.image_only:
            try:
                apply_ocr(model, options)
            except Exception:
                continue

    stem = Path(source_name).stem or "converted"
    output_path = output_dir / f"{stem}.pptx"

    builder = PptxBuilder(keep_art_layer=options.keep_art_layer)
    try:
        builder.build(pages, art_images, output_path)
    except Exception as exc:
        raise ConversionError(f"PPTX 生成失败：{exc}") from exc

    report = build_report(pages, set(art_images.keys()), source_name)
    report_path = output_dir / "report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    return ConversionResult(output_path=output_path, report_path=report_path, report=report)
