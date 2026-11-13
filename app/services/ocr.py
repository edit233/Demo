"""OCR 模块：把「图片型 PDF 页面」还原为可编辑文字。

实际测试素材是整页位图形式的 PPT 页面，PDF 内没有任何文本层与矢量
对象，无法直接提取文字。本模块负责：

1. 调用 RapidOCR（PP-OCRv4，中英文）检测页面中的文字区域；
2. 以多个尺度检测并合并结果，避免漏掉字距较大或较小的文字；
3. 按行合并检测框后用「仅识别」模式重新识别，得到带空格的完整文本；
4. 依据文字在页面中的位置与颜色，还原为 PPT 文本框所需的属性；
5. 把文字区域从底图中「抹掉」，得到不含文字的背景图，使叠加后的
   可编辑文字与原页面观感一致。
"""

from __future__ import annotations

import io
import threading
from dataclasses import dataclass, field

import numpy as np
from PIL import Image, ImageFilter

from app.services.fonts import has_cjk

try:  # pragma: no cover - 依赖缺失时降级为纯图片输出
    from rapidocr_onnxruntime import RapidOCR
except Exception:  # pragma: no cover
    RapidOCR = None

#: OCR 结果在页面中的矩形（点）
OcrRect = tuple[float, float, float, float]
#: 像素坐标下的矩形
PixelRect = tuple[int, int, int, int]


@dataclass
class OcrText:
    """一段被识别出来的文字。"""

    text: str
    rect: OcrRect
    score: float
    color: tuple[int, int, int]
    size: float
    #: 字符间距（点），用于还原设计稿中的字距
    spacing: float = 0.0
    #: 文字墨迹的紧致范围（页面坐标）；为空时退化为使用 ``rect``
    ink_rect: OcrRect | None = None

    def layout_rect(self) -> OcrRect:
        """排版用的文本框位置：比文字墨迹略高，补偿字体的上伸部。"""
        box = self.ink_rect or self.rect
        offset = 0.12 * self.size if has_cjk(self.text) else 0.06 * self.size
        return (box[0], box[1] - offset, box[2], box[3])


@dataclass
class _Detection:
    """一次检测得到的文字区域（像素坐标）。"""

    rect: PixelRect
    score: float = 0.0
    text: str = ""
    source: str = ""


@dataclass
class _Line:
    """按行归并后的文字区域。"""

    rect: PixelRect
    detections: list[_Detection] = field(default_factory=list)


def _vertical_overlap(a: PixelRect, b: PixelRect) -> float:
    """两个框在垂直方向的重叠比例。"""
    top = max(a[1], b[1])
    bottom = min(a[3], b[3])
    if bottom <= top:
        return 0.0
    smaller = min(a[3] - a[1], b[3] - b[1])
    return (bottom - top) / smaller if smaller else 0.0


def _uncovered_segments(rect: PixelRect, accepted: list[PixelRect]) -> list[tuple[int, int]]:
    """求 ``rect`` 中未被 ``accepted`` 覆盖的水平片段。"""
    segments = [(rect[0], rect[2])]
    for other in accepted:
        if _vertical_overlap(rect, other) < 0.4:
            continue
        updated: list[tuple[int, int]] = []
        for x0, x1 in segments:
            if other[2] <= x0 or other[0] >= x1:
                updated.append((x0, x1))
                continue
            if other[0] > x0:
                updated.append((x0, other[0]))
            if other[2] < x1:
                updated.append((other[2], x1))
        segments = updated
    return segments


def _merge_detections(detections: list[_Detection]) -> list[_Detection]:
    """合并多尺度检测结果，保留未覆盖的新区域。"""
    accepted: list[_Detection] = []
    for detection in detections:
        if detection.rect[2] - detection.rect[0] < 4 or detection.rect[3] - detection.rect[1] < 4:
            continue
        covered = [
            item
            for item in accepted
            if _vertical_overlap(detection.rect, item.rect) >= 0.4
        ]
        if not covered:
            accepted.append(detection)
            continue
        segments = _uncovered_segments(detection.rect, [item.rect for item in covered])
        total = detection.rect[2] - detection.rect[0]
        for x0, x1 in segments:
            if (x1 - x0) < max(12, total * 0.3):
                continue
            accepted.append(
                _Detection(
                    rect=(x0, detection.rect[1], x1, detection.rect[3]),
                    score=detection.score,
                    source=detection.source,
                )
            )
    return accepted


def _group_lines(detections: list[_Detection]) -> list[_Line]:
    """把水平相邻、垂直重叠的检测框合并为一行。

    分两步进行：

    1. 依据垂直重叠把检测框聚成「行带」，同一行带内的检测框垂直方向
       互相咬合，必然属于同一行文字；
    2. 行带内再按水平间距切分，避免把左右两栏的内容并成一行。

    先聚行带再切分，可以避免「右侧片段先建行、左侧片段因水平间隙过大
    而无法归并」的问题（例如 "BUILD" 与 "A BETTER TOMORROW" 被拆成两行）。
    """
    bands: list[list[_Detection]] = []
    band_rects: list[PixelRect] = []
    for detection in sorted(detections, key=lambda item: (item.rect[1], item.rect[0])):
        placed = False
        for index, band in enumerate(bands):
            # 要求与行带内至少一个成员垂直咬合，避免行带纵向膨胀后误吸相邻行
            if not any(_vertical_overlap(detection.rect, item.rect) >= 0.45 for item in band):
                continue
            band.append(detection)
            box = band_rects[index]
            band_rects[index] = (
                min(box[0], detection.rect[0]),
                min(box[1], detection.rect[1]),
                max(box[2], detection.rect[2]),
                max(box[3], detection.rect[3]),
            )
            placed = True
            break
        if not placed:
            bands.append([detection])
            band_rects.append(detection.rect)

    lines: list[_Line] = []
    for band in bands:
        members = sorted(band, key=lambda item: item.rect[0])
        group: list[_Detection] = [members[0]]
        for current in members[1:]:
            previous = group[-1]
            height = max(1, current.rect[3] - current.rect[1])
            gap = max(previous.rect[0] - current.rect[2], current.rect[0] - previous.rect[2])
            if gap > height * 2.5:
                lines.append(_make_line(group))
                group = [current]
            else:
                group.append(current)
        lines.append(_make_line(group))
    return lines


def _make_line(members: list[_Detection]) -> _Line:
    """由若干检测框构造一行。"""
    rect = (
        min(item.rect[0] for item in members),
        min(item.rect[1] for item in members),
        max(item.rect[2] for item in members),
        max(item.rect[3] for item in members),
    )
    return _Line(rect=rect, detections=list(members))


class OcrEngine:
    """RapidOCR 封装：懒加载 + 线程安全。"""

    _instance: "OcrEngine | None" = None
    _instance_lock = threading.Lock()

    def __init__(self) -> None:
        self._engine = None
        self._engine_lock = threading.Lock()

    @classmethod
    def instance(cls) -> "OcrEngine":
        with cls._instance_lock:
            if cls._instance is None:
                cls._instance = cls()
            return cls._instance

    @staticmethod
    def available() -> bool:
        """OCR 依赖是否可用。"""
        return RapidOCR is not None

    def _ensure_engine(self):
        if self._engine is None:
            with self._engine_lock:
                if self._engine is None:
                    self._engine = RapidOCR()
        return self._engine

    # ------------------------------------------------------------------ 识别
    def recognize(
        self,
        image_bytes: bytes,
        rect: OcrRect,
        min_confidence: float = 0.5,
        scales: tuple[float, ...] = (1.0, 2.0),
    ) -> list[OcrText]:
        """识别图片中的文字，返回页面坐标（点）下的文字块。"""
        if not self.available():
            return []
        try:
            with Image.open(io.BytesIO(image_bytes)) as handle:
                image = handle.convert("RGB")
        except Exception:
            return []
        array = np.array(image)

        engine = self._ensure_engine()
        detections: list[_Detection] = []
        for scale in scales:
            detections.extend(self._detect(engine, array, scale))

        merged = _merge_detections(detections)
        lines = _group_lines(merged)

        texts: list[OcrText] = []
        span_x = (rect[2] - rect[0]) / max(array.shape[1], 1)
        span_y = (rect[3] - rect[1]) / max(array.shape[0], 1)
        for line in lines:
            recognized = self._recognize_strip(engine, array, line)
            if not recognized:
                continue
            content, score = recognized
            if score < min_confidence:
                continue
            box = line.rect
            # 检测框通常比真实墨迹大，用像素投影测量紧致的墨迹范围
            ink, segment_widths, segment_gaps = measure_ink(array, box)
            page_rect = self._to_page_rect(box, rect, array.shape[1], array.shape[0])
            ink_rect = self._to_page_rect(ink, rect, array.shape[1], array.shape[0])
            ink_height = (ink[3] - ink[1]) * span_y
            ink_width = (ink[2] - ink[0]) * span_x
            size, spacing = estimate_text_style(
                content,
                ink_height,
                ink_width,
                [width * span_x for width in segment_widths],
                [gap * span_x for gap in segment_gaps],
            )
            texts.append(
                OcrText(
                    text=content,
                    rect=page_rect,
                    ink_rect=ink_rect,
                    score=score,
                    color=estimate_text_color(array, box),
                    size=size,
                    spacing=spacing,
                )
            )
        return texts

    def _detect(self, engine, array: np.ndarray, scale: float) -> list[_Detection]:
        """在指定尺度上做文字检测。"""
        if scale == 1.0:
            working = array
        else:
            height = int(array.shape[0] * scale)
            width = int(array.shape[1] * scale)
            working = np.array(
                Image.fromarray(array).resize((width, height), Image.LANCZOS)
            )
        try:
            with self._engine_lock:
                result, _ = engine(working)
        except Exception:
            return []
        detections: list[_Detection] = []
        for item in result or []:
            try:
                box, score = item[0], float(item[2])
            except Exception:
                continue
            xs = [float(point[0]) / scale for point in box]
            ys = [float(point[1]) / scale for point in box]
            detections.append(
                _Detection(
                    rect=(
                        int(min(xs)),
                        int(min(ys)),
                        int(max(xs)),
                        int(max(ys)),
                    ),
                    score=score,
                    source=f"{scale}x",
                )
            )
        return detections

    def _recognize_strip(self, engine, array: np.ndarray, line: _Line) -> tuple[str, float]:
        """对整行做「仅识别」，得到带空格的完整文本。"""
        x0, y0, x1, y1 = line.rect
        pad = 2
        crop = array[
            max(0, y0 - pad) : min(array.shape[0], y1 + pad),
            max(0, x0 - pad) : min(array.shape[1], x1 + pad),
        ]
        if crop.size == 0:
            return "", 0.0
        try:
            with self._engine_lock:
                result, _ = engine(crop, use_det=False, use_cls=False, use_rec=True)
        except Exception:
            return "", 0.0
        if not result:
            return "", 0.0
        try:
            first = result[0]
            content = str(first[1] if len(first) > 2 else first[0]).strip()
            score = float(first[2]) if len(first) > 2 else float(first[1])
        except Exception:
            return "", 0.0
        return content, score

    @staticmethod
    def _span_y(rect: OcrRect, image_height: int) -> float:
        return (rect[3] - rect[1]) / max(image_height, 1)

    @staticmethod
    def _to_page_rect(
        box: PixelRect, rect: OcrRect, image_width: int, image_height: int
    ) -> OcrRect:
        span_x = (rect[2] - rect[0]) / max(image_width, 1)
        span_y = (rect[3] - rect[1]) / max(image_height, 1)
        return (
            rect[0] + box[0] * span_x,
            rect[1] + box[1] * span_y,
            rect[0] + box[2] * span_x,
            rect[1] + box[3] * span_y,
        )


def estimate_text_color(array: np.ndarray, pixel_rect) -> tuple[int, int, int]:
    """估算文字颜色：取与背景差异最大的那部分像素的中位色。"""
    x0, y0, x1, y1 = pixel_rect
    region = array[y0:y1, x0:x1].reshape(-1, 3).astype(np.float32)
    if region.size == 0:
        return (0, 0, 0)
    background = np.median(region, axis=0)
    distances = np.linalg.norm(region - background, axis=1)
    mask = distances > 55
    if int(mask.sum()) < 3:
        return tuple(int(value) for value in background)  # type: ignore[return-value]
    selected = region[mask]
    # 取差异最大的 60%，避免抗锯齿边缘把颜色拉浅
    keep = max(3, int(len(selected) * 0.6))
    order = np.argsort(-distances[mask])[:keep]
    core = selected[order]
    color = np.median(core, axis=0)
    return tuple(int(max(0, min(255, round(value)))) for value in color)  # type: ignore[return-value]


def _segments(indexes: np.ndarray) -> list[tuple[int, int]]:
    """把连续的索引序列切分成若干段。"""
    if indexes.size == 0:
        return []
    breaks = np.nonzero(np.diff(indexes) > 1)[0]
    starts = np.concatenate(([0], breaks + 1))
    ends = np.concatenate((breaks, [len(indexes) - 1]))
    return [(int(indexes[s]), int(indexes[e]) + 1) for s, e in zip(starts, ends)]


def measure_ink(
    array: np.ndarray, pixel_rect: PixelRect
) -> tuple[PixelRect, list[int], list[int]]:
    """测量检测框内文字墨迹的紧致范围（像素坐标）。

    OCR 检测框往往比真实文字大（含上下左右留白），直接用它推算字号会偏大。
    这里在框内做二值化，再用行投影 / 列投影求出墨迹的紧致范围。

    返回 ``(墨迹框, 各列墨迹段的宽度, 相邻段之间的空白宽度)``。对于中文，
    列段就是单个汉字的墨迹宽度，段间空白就是设计稿中的字间距。
    """
    x0, y0, x1, y1 = pixel_rect
    x0 = max(0, int(x0))
    y0 = max(0, int(y0))
    x1 = min(array.shape[1], int(x1))
    y1 = min(array.shape[0], int(y1))
    if x1 - x0 < 2 or y1 - y0 < 2:
        return (x0, y0, x1, y1), [], []

    region = array[y0:y1, x0:x1]
    flat = region.reshape(-1, 3).astype(np.float32)
    background = np.median(flat, axis=0)
    distance = np.linalg.norm(region.astype(np.float32) - background, axis=2)
    mask = distance > 45

    row_bands = _segments(np.nonzero(mask.sum(axis=1) > 0)[0])
    column_segments = _segments(np.nonzero(mask.sum(axis=0) > 0)[0])
    if not row_bands or not column_segments:
        return (x0, y0, x1, y1), [], []

    band = max(row_bands, key=lambda item: item[1] - item[0])
    widths = [segment[1] - segment[0] for segment in column_segments]
    gaps = [
        column_segments[index][0] - column_segments[index - 1][1]
        for index in range(1, len(column_segments))
    ]

    ink_x0 = x0 + column_segments[0][0]
    ink_x1 = x0 + column_segments[-1][1]

    # 墨迹只占框高的一小部分时，说明框内混入了相邻行，此时保留原高度
    if (band[1] - band[0]) < (y1 - y0) * 0.45:
        return (ink_x0, y0, ink_x1, y1), widths, gaps
    return (ink_x0, y0 + band[0], ink_x1, y0 + band[1]), widths, gaps


def char_width_factor(text: str) -> float:
    """文本的平均字符宽度（以 em 为单位）。"""
    total = 0.0
    for char in text:
        if has_cjk(char):
            total += 1.0
        elif char.isupper():
            total += 0.62
        elif char.islower():
            total += 0.52
        elif char.isdigit():
            total += 0.55
        elif char == " ":
            total += 0.28
        else:
            total += 0.38
    return total


def estimate_font_size(text: str, box_height: float) -> float:
    """根据文本墨迹高度估算字号（点）。"""
    if box_height <= 0:
        return 12.0
    if has_cjk(text):
        ratio = 0.85
    elif any(char.islower() for char in text):
        ratio = 0.70
    else:
        ratio = 0.72
    return box_height / ratio


def estimate_text_style(
    text: str,
    ink_height: float,
    ink_width: float,
    segment_widths: list[float] | None = None,
    segment_gaps: list[float] | None = None,
) -> tuple[float, float]:
    """估算字号与字符间距（点）。

    * 字号优先由文字墨迹高度推算；若文字以中文为主且列投影能把单个汉字
      分离出来，则改用「单字墨迹宽」反推（汉字字面宽 ≈ 0.95 em）。
    * 字间距优先由列投影中相邻墨迹段之间的空白直接测得（中文标题常带
      宽字距）；无法测得时按「总宽差额」分摊。
    """
    factor = char_width_factor(text)
    if factor <= 0:
        return 12.0, 0.0

    size = estimate_font_size(text, ink_height)
    spacing = 0.0

    # --- 中文：用单字墨迹宽反推字号，用列间隙测字距 ---
    if segment_widths and has_cjk(text) and len(segment_widths) >= max(2, len(text) // 2):
        widths = sorted(segment_widths)
        median_width = widths[len(widths) // 2]
        if median_width > 0:
            size = median_width / 0.95
        if segment_gaps:
            positive = sorted(gap for gap in segment_gaps if gap > 0.05)
            if positive:
                spacing = positive[len(positive) // 2]

    # --- 总宽约束：不得超出原页面上的文字宽度 ---
    if ink_width > 0:
        natural = factor * size
        total = natural + spacing * len(text)
        if total > ink_width:
            scale = ink_width / total if total > 0 else 1.0
            size *= scale
            spacing *= scale
        elif spacing <= 0:
            count = len(text)
            if count > 1:
                spacing = max(0.0, (ink_width - natural) / count)
    return size, spacing


def erase_regions(
    image_bytes: bytes,
    regions: list[PixelRect],
    quality: int = 92,
) -> bytes:
    """把指定区域从图片中抹掉，返回新的图片字节。

    * 背景平坦的区域（外圈采样颜色接近）：直接填充为采样到的背景色；
    * 背景复杂的区域：按「文字颜色」逐像素替换为背景色，消除文字笔画。
    """
    with Image.open(io.BytesIO(image_bytes)) as handle:
        image = handle.convert("RGB")
    array = np.array(image)

    for x0, y0, x1, y1 in regions:
        x0 = max(0, x0)
        y0 = max(0, y0)
        x1 = min(array.shape[1], x1)
        y1 = min(array.shape[0], y1)
        if x1 - x0 < 2 or y1 - y0 < 2:
            continue

        # 取矩形外圈作为背景采样（文字不会延伸到外圈）
        band = 8
        ring_parts = []
        for ring_y0, ring_y1, ring_x0, ring_x1 in (
            (max(0, y0 - band), y0, x0, x1),
            (y1, min(array.shape[0], y1 + band), x0, x1),
            (y0, y1, max(0, x0 - band), x0),
            (y0, y1, x1, min(array.shape[1], x1 + band)),
        ):
            if ring_y1 > ring_y0 and ring_x1 > ring_x0:
                ring_parts.append(array[ring_y0:ring_y1, ring_x0:ring_x1].reshape(-1, 3))
        if not ring_parts:
            continue

        ring = np.concatenate(ring_parts).astype(np.float32)
        background = np.median(ring, axis=0)
        # 背景平坦程度：外圈各像素与背景色的偏差
        deviation = float(np.abs(ring - background).mean())
        local_variance = float(np.var(ring, axis=0).mean())

        if deviation < 12.0 and local_variance < 260.0:
            # 背景平坦 → 直接填充背景色
            array[y0:y1, x0:x1] = background.astype(np.uint8)
            continue

        # 背景复杂 → 按文字颜色做像素级替换，保留背景纹理
        patch = array[y0:y1, x0:x1].astype(np.float32)
        text_color = np.array(estimate_text_color(array, (x0, y0, x1, y1)), dtype=np.float32)
        distance_text = np.linalg.norm(patch - text_color, axis=2)
        # 阈值取文字色与背景色距离的 75%，连同抗锯齿过渡像素一并替换
        threshold = max(
            20.0, float(np.linalg.norm(text_color - background)) * 0.75
        )
        mask = distance_text < threshold
        if mask.any():
            patch[mask] = background
        blurred = Image.fromarray(patch.astype(np.uint8)).filter(
            ImageFilter.GaussianBlur(radius=0.8)
        )
        array[y0:y1, x0:x1] = np.array(blurred)

    cleaned = Image.fromarray(array)
    buffer = io.BytesIO()
    cleaned.save(buffer, format="JPEG", quality=quality, optimize=True)
    return buffer.getvalue()


def pixel_regions(
    texts: list[OcrText],
    rect: OcrRect,
    image_size: tuple[int, int],
    padding: int = 2,
) -> list[PixelRect]:
    """把页面坐标下的文字块换算成像素坐标，供抹除使用。"""
    width, height = image_size
    span_x = (rect[2] - rect[0]) / max(width, 1)
    span_y = (rect[3] - rect[1]) / max(height, 1)
    regions: list[PixelRect] = []
    for item in texts:
        x0 = int((item.rect[0] - rect[0]) / span_x) - padding
        y0 = int((item.rect[1] - rect[1]) / span_y) - padding
        x1 = int((item.rect[2] - rect[0]) / span_x) + padding
        y1 = int((item.rect[3] - rect[1]) / span_y) + padding
        x0 = max(0, x0)
        y0 = max(0, y0)
        x1 = min(width, x1)
        y1 = min(height, y1)
        if x1 > x0 and y1 > y0:
            regions.append((x0, y0, x1, y1))
    return regions