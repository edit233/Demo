"""图层分解：把「整页位图」拆成背景 + 若干独立图片对象 + 原生形状。

为什么需要这个模块
------------------
本项目三套测试素材都是**整页位图型 PDF**：每页只有一张覆盖全页的 JPEG，
既没有文本层也没有矢量对象。照片、插画、Logo、装饰图形与背景在像素层面
被压平在同一张图里。如果只把这张整页图作为「背景」放进 PPT，照片就永远
被封在背景里，**无法独立选择、移动、缩放**。

因此要在像素层面把「图片区域」重新切出来，还原成一个个独立对象。

分割思路
--------
1. **纹理能量图**：对灰度图做局部标准差。照片 / 插画 / Logo 的像素级纹理
   能量高；纯色底、渐变底、平涂几何图形的纹理能量接近 0（只有边缘一两
   像素高）。这是区分「图片」与「设计图形」最稳的信号。
2. **文字感知**：用 OCR 文字框构造文字掩膜，并判断每段文字是压在图片上还是
   压在平底上。压在图片上的文字连同图片一起保留（导出时局部修复），
   压在平底上的文字从图片候选中剔除，避免文字变成图片对象。
3. **线条分离**：横贯页面的装饰线会把所有内容串成一个组件。先把细长条状
   的高纹理区域识别为「线条」并从图片候选中剔除，线条另行还原为 PPT
   原生连接符。
4. **连通域**：形态学闭运算后做连通域标记，得到候选图片区域；逐块做孔洞
   填充，把图片上文字造成的空洞补回。
5. **独立对象**：每个区域独立裁剪成带 alpha 通道的 PNG，作为独立图片对象
   写入 PPTX，可单独选中、移动、缩放。
6. **背景重建**：背景层 = 原图 − 图片区域。纯色背景直接用 PPT **原生矩形**
   （完全可编辑）；平涂色块、线条、斜切多边形还原为 PPT 原生形状；其余
   复杂背景用重建后的整页图。

本模块只依赖 numpy + Pillow，不引入 OpenCV / SciPy。
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field

import numpy as np
from PIL import Image, ImageFilter

from app.services.pdf_parser import Rect, ShapeElement

#: 像素矩形
PixelRect = tuple[int, int, int, int]


@dataclass
class LayeringOptions:
    """图层分解参数。"""

    #: 是否启用图层分解
    enable: bool = True
    #: 纹理能量阈值（局部标准差）
    texture_threshold: float = 9.0
    #: 局部标准差的取样半径（像素）
    texture_radius: int = 3
    #: 候选图片区域最小面积占页面比例
    min_area_ratio: float = 0.0015
    #: 候选图片区域最小包围盒面积占页面比例（细笔画 Logo 走这条判据）
    min_bbox_area_ratio: float = 0.0012
    #: 候选图片区域的墨迹像素绝对下限（过滤噪点）
    min_ink_area: int = 120
    #: 页面边缘需要排除的像素数（JPEG 接缝会在最外圈形成一圈假纹理）
    edge_trim: int = 5
    #: 形态学闭运算半径（像素）
    close_radius: int = 2
    #: alpha 羽化半径（像素）
    feather_radius: float = 1.0
    #: 单页最多拆出的图片对象数
    max_regions: int = 36
    #: 是否把背景重建为图片（关闭则背景只保留纯色矩形 / 原生形状）
    rebuild_background: bool = True
    #: 判定「纯色背景」的像素占比阈值
    flat_ratio: float = 0.985
    #: 判定「纯色背景」的颜色容差（0~255）
    flat_tolerance: float = 10.0

    # --- 线条 ---
    #: 线条最小长度（像素）
    line_min_length: int = 18
    #: 线条最大厚度（像素）
    line_max_thickness: int = 15
    #: 线条最小长宽比
    line_min_aspect: float = 5.0
    #: 线条颜色均匀性上限（照片内部的树枝/屋檐颜色杂乱，会被排除）
    line_color_std: float = 14.0
    #: 线条周边纹理密度上限（超过则判定为长在照片内部，不是设计线条）
    line_context_density: float = 0.45

    # --- 平涂色块 ---
    #: 色块最小面积占页面比例
    block_min_area_ratio: float = 0.004
    #: 色块最小实心度
    block_min_solidity: float = 0.90
    #: 参与检测的颜色层数上限
    block_max_levels: int = 20
    #: 色块边界与外部的最小对比度（过滤渐变色带）
    block_boundary_contrast: float = 16.0
    #: 色块内部颜色标准差上限
    block_color_std: float = 12.0
    #: 判定「平坦区」的局部标准差上限
    flat_threshold: float = 4.0

    # --- 自由曲线 ---
    #: 自由曲线最小面积占页面比例
    freeform_min_area_ratio: float = 0.02
    #: 自由曲线最多边数
    freeform_max_sides: int = 5
    #: 图片区域面积占页面比例低于此值时，视为「图内文字」（Logo / 配图里的字）
    artwork_area_ratio: float = 0.08
    #: 文字框覆盖包围盒的比例高于此值时，判定为「文字块」而非图片区域
    text_block_ratio: float = 0.80
    #: 文字与图片区域的重叠面积占比达到此值时，判定为「压在图上」
    text_overlap_ratio: float = 0.30


@dataclass
class ImageRegion:
    """从整页位图中切出的一块图片。"""

    #: 像素包围盒 (x0, y0, x1, y1)
    bbox_px: PixelRect
    #: 页面坐标下的包围盒（点）
    rect: Rect
    #: 区域掩膜（页面尺寸，布尔）
    mask: np.ndarray
    #: 区域类型：photo / illustration / logo / graphic
    kind: str = "graphic"

    @property
    def area(self) -> float:
        x0, y0, x1, y1 = self.bbox_px
        return float(max(0, x1 - x0) * max(0, y1 - y0))


@dataclass
class LayeringResult:
    """一页的图层分解结果。"""

    #: 切出的图片区域
    regions: list[ImageRegion] = field(default_factory=list)
    #: 还原为 PPT 原生对象的图形（矩形 / 直线 / 自由曲线）
    shapes: list[ShapeElement] = field(default_factory=list)
    #: 重建后的背景位图（PNG 字节）；为 None 表示背景是纯色或已被原生对象覆盖
    background: bytes | None = None
    #: 纯色背景的颜色；为 None 表示背景不是纯色
    flat_color: tuple[int, int, int] | None = None
    #: 每个图片区域内部包含的文字（下标, 像素矩形）
    region_text_rects: dict[int, list[tuple[int, PixelRect]]] = field(default_factory=dict)
    #: 属于「图内文字」（Logo / 配图里的字）的文字下标
    artwork_text: set[int] = field(default_factory=set)
    #: 需要从背景层抹掉的文字（像素矩形）
    background_text_rects: list[PixelRect] = field(default_factory=list)
    #: 统计信息
    stats: dict = field(default_factory=dict)


# --------------------------------------------------------------------- 算子
def _box_mean(a: np.ndarray, radius: int) -> np.ndarray:
    """积分图求盒式均值。窗口覆盖原图 [i-r, i+r]，越界按 edge 复制。"""
    h, w = a.shape
    pad = np.pad(a.astype(np.float64), radius, mode="edge")
    s2 = pad.cumsum(0).cumsum(1)
    s2 = np.pad(s2, ((1, 0), (1, 0)), mode="constant")
    k = 2 * radius + 1
    total = s2[k : k + h, k : k + w] - s2[0:h, k : k + w] - s2[k : k + h, 0:w] + s2[0:h, 0:w]
    return total / (k * k)


def _local_std(gray: np.ndarray, radius: int) -> np.ndarray:
    """局部标准差 —— 纹理能量图。"""
    mean = _box_mean(gray, radius)
    mean_sq = _box_mean(gray * gray, radius)
    return np.sqrt(np.maximum(mean_sq - mean * mean, 0.0))


def _dilate(mask: np.ndarray, radius: int) -> np.ndarray:
    """方形结构元的膨胀。"""
    if radius <= 0:
        return mask.copy()
    out = mask.copy()
    h, w = out.shape
    for _ in range(radius):
        p = np.pad(out, 1, mode="constant", constant_values=False)
        out = np.logical_or.reduce(
            [p[dy : dy + h, dx : dx + w] for dy in range(3) for dx in range(3)]
        )
    return out


def _erode(mask: np.ndarray, radius: int) -> np.ndarray:
    """方形结构元的腐蚀。"""
    if radius <= 0:
        return mask.copy()
    return ~_dilate(~mask, radius)


def _erode_1d(mask: np.ndarray, length: int, axis: int) -> np.ndarray:
    """一维（行 / 列）方向的腐蚀。"""
    if length <= 1:
        return mask.copy()
    half = length // 2
    pad = [(0, 0), (0, 0)]
    pad[axis] = (half, length - 1 - half)
    p = np.pad(mask, pad, mode="constant", constant_values=False)
    h, w = mask.shape
    if axis == 1:
        out = p[:, 0:w]
        for shift in range(1, length):
            out = np.logical_and(out, p[:, shift : shift + w])
    else:
        out = p[0:h, :]
        for shift in range(1, length):
            out = np.logical_and(out, p[shift : shift + h, :])
    return out


def _dilate_1d(mask: np.ndarray, length: int, axis: int) -> np.ndarray:
    """一维（行 / 列）方向的膨胀。"""
    if length <= 1:
        return mask.copy()
    half = length // 2
    pad = [(0, 0), (0, 0)]
    pad[axis] = (half, length - 1 - half)
    p = np.pad(mask, pad, mode="constant", constant_values=False)
    h, w = mask.shape
    if axis == 1:
        out = p[:, 0:w]
        for shift in range(1, length):
            out = np.logical_or(out, p[:, shift : shift + w])
    else:
        out = p[0:h, :]
        for shift in range(1, length):
            out = np.logical_or(out, p[shift : shift + h, :])
    return out


def _label_components(mask: np.ndarray) -> tuple[np.ndarray, int]:
    """按行程（run）做连通域标记，比逐像素洪泛快得多。

    临时标签从 1 开始，0 保留给背景，避免重映射时把背景算进组件。
    """
    h, w = mask.shape
    parent: list[int] = [0]

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[max(ra, rb)] = min(ra, rb)

    labels = np.zeros((h, w), dtype=np.int32)
    prev: list[tuple[int, int, int]] = []
    for y in range(h):
        row = mask[y].astype(np.uint8)
        idx = np.flatnonzero(np.diff(np.concatenate(([0], row, [0]))))
        cur: list[tuple[int, int, int]] = []
        for s, e in zip(idx[0::2], idx[1::2]):
            lbl = len(parent)
            parent.append(lbl)
            for ps, pe, pl in prev:
                if ps < e and s < pe:
                    union(lbl, pl)
            cur.append((int(s), int(e), lbl))
            labels[y, int(s) : int(e)] = lbl
        prev = cur

    if len(parent) == 1:
        return labels, 0

    remap = np.zeros(len(parent), dtype=np.int32)
    count = 0
    for i in range(1, len(parent)):
        root = find(i)
        if remap[root] == 0:
            count += 1
            remap[root] = count
        remap[i] = remap[root]
    return remap[labels], count


def _component_boxes(labels: np.ndarray, count: int) -> list[dict]:
    """每个连通域的面积与包围盒。"""
    stats: list[dict] = []
    for index in range(1, count + 1):
        ys, xs = np.nonzero(labels == index)
        if ys.size == 0:
            continue
        stats.append(
            {
                "id": index,
                "area": int(ys.size),
                "bbox": (int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1),
            }
        )
    return stats


def _fill_holes_per_component(mask: np.ndarray, labels: np.ndarray, count: int) -> np.ndarray:
    """逐个连通域填充内部孔洞。

    必须逐块做：整幅图一次性填充会把「被版面元素围起来的大片背景」误判成
    孔洞（画册页眉/页脚的长条会把页面切成两半）。

    同时要求组件本身足够「实心」：细长的边框、分隔线若也填洞，它的包围盒
    往往覆盖整页，会把页内所有内容误当成孔洞补上。
    """
    filled = mask.copy()
    for index in range(1, count + 1):
        ys, xs = np.nonzero(labels == index)
        if ys.size == 0:
            continue
        y0, y1 = int(ys.min()), int(ys.max()) + 1
        x0, x1 = int(xs.min()), int(xs.max()) + 1
        bbox_area = float((x1 - x0) * (y1 - y0))
        solidity = float(ys.size) / max(bbox_area, 1.0)
        if solidity < 0.12:
            continue
        sub = labels[y0:y1, x0:x1] == index
        holes = _enclosed_holes(sub)
        if holes.any():
            filled[y0:y1, x0:x1] |= holes
    return filled


def _enclosed_holes(mask: np.ndarray) -> np.ndarray:
    """返回被掩膜完全包围的孔洞。"""
    inv = ~mask
    if not inv.any():
        return np.zeros_like(mask)
    labels, count = _label_components(inv)
    if count == 0:
        return np.zeros_like(mask)
    border = set()
    border |= set(np.unique(labels[0, :]).tolist())
    border |= set(np.unique(labels[-1, :]).tolist())
    border |= set(np.unique(labels[:, 0]).tolist())
    border |= set(np.unique(labels[:, -1]).tolist())
    holes = np.zeros_like(mask)
    for index in range(1, count + 1):
        if index not in border:
            holes |= labels == index
    return holes


def _resize_bilinear(a: np.ndarray, out_h: int, out_w: int) -> np.ndarray:
    """双线性缩放（支持 2D / 3D 数组）。"""
    h, w = a.shape[:2]
    xi = np.clip(np.linspace(0.0, w - 1, out_w), 0.0, w - 1)
    x0 = np.floor(xi).astype(int)
    x1 = np.minimum(x0 + 1, w - 1)
    if a.ndim == 3:
        fx = (xi - x0)[None, :, None]
        tmp = a[:, x0, :] + (a[:, x1, :] - a[:, x0, :]) * fx
    else:
        fx = (xi - x0)[None, :]
        tmp = a[:, x0] + (a[:, x1] - a[:, x0]) * fx
    yi = np.clip(np.linspace(0.0, h - 1, out_h), 0.0, h - 1)
    y0 = np.floor(yi).astype(int)
    y1 = np.minimum(y0 + 1, h - 1)
    if a.ndim == 3:
        fy = (yi - y0)[:, None, None]
        top = tmp[y0, :, :]
        bot = tmp[y1, :, :]
    else:
        fy = (yi - y0)[:, None]
        top = tmp[y0, :]
        bot = tmp[y1, :]
    return top + (bot - top) * fy


def _downsample(a: np.ndarray) -> np.ndarray:
    """2x2 盒式降采样。"""
    h, w = a.shape[:2]
    h2, w2 = h // 2 * 2, w // 2 * 2
    if h2 < 2 or w2 < 2:
        return a
    if a.ndim == 3:
        return a[:h2, :w2].reshape(h2 // 2, 2, w2 // 2, 2, a.shape[2]).mean(axis=(1, 3))
    return a[:h2, :w2].reshape(h2 // 2, 2, w2 // 2, 2).mean(axis=(1, 3))


def _pull_push_fill(rgb: np.ndarray, hole: np.ndarray) -> np.ndarray:
    """金字塔 push-pull 填补空洞。

    自顶向下逐层「拉」回原分辨率：有效像素保留自身颜色，无效像素由上一级
    放大后填补，并按有效度做平滑过渡。对大面积缺失（整块照片被抽走）能
    重建出低频的背景渐变。
    """
    valid = ~hole
    if not valid.any():
        return rgb.astype(np.float32)

    pyr_img: list[np.ndarray] = [rgb.astype(np.float32)]
    pyr_val: list[np.ndarray] = [valid.astype(np.float32)]
    while min(pyr_img[-1].shape[0], pyr_img[-1].shape[1]) > 8:
        pyr_img.append(_downsample(pyr_img[-1]))
        pyr_val.append(_downsample(pyr_val[-1]))

    top = pyr_img[-1]
    tv = pyr_val[-1]
    weight = tv[..., None]
    total = (top * weight).reshape(-1, 3).sum(axis=0)
    denom = float(weight.sum())
    mean = total / max(denom, 1e-6)
    filled = np.where(tv[..., None] > 0.99, top, mean.reshape(1, 1, 3))

    for level in range(len(pyr_img) - 2, -1, -1):
        up = _resize_bilinear(filled, pyr_img[level].shape[0], pyr_img[level].shape[1])
        v = np.clip(pyr_val[level], 0.0, 1.0)[..., None]
        filled = pyr_img[level] * v + up * (1.0 - v)
    return filled


def _feather(mask: np.ndarray, radius: float) -> np.ndarray:
    """把布尔掩膜变成带羽化的 alpha（0~1）。"""
    alpha = mask.astype(np.float32)
    if radius <= 0:
        return alpha
    image = Image.fromarray((alpha * 255).astype(np.uint8), mode="L")
    blurred = image.filter(ImageFilter.GaussianBlur(radius=radius))
    return np.array(blurred).astype(np.float32) / 255.0


def _rect_to_pixels(rect: PixelRect, size: tuple[int, int]) -> PixelRect:
    width, height = size
    x0 = max(0, min(width, rect[0]))
    y0 = max(0, min(height, rect[1]))
    x1 = max(0, min(width, rect[2]))
    y1 = max(0, min(height, rect[3]))
    return x0, y0, x1, y1


def _mask_from_rects(rects: list[PixelRect], size: tuple[int, int]) -> np.ndarray:
    """由像素矩形列表生成掩膜；``size`` 为 (宽, 高)。"""
    width, height = size
    mask = np.zeros((height, width), dtype=bool)
    for rect in rects:
        x0, y0, x1, y1 = _rect_to_pixels(rect, (width, height))
        if x1 > x0 and y1 > y0:
            mask[y0:y1, x0:x1] = True
    return mask


def _ring_texture_ratio(
    std_map: np.ndarray, rect: PixelRect, threshold: float, band: int = 6
) -> float:
    """文字框外圈中「有纹理」的像素占比。

    占比高说明这段文字压在图片上（照片、插画），否则说明它压在平底上。
    """
    height, width = std_map.shape
    x0, y0, x1, y1 = rect
    band = max(2, band)
    ox0 = max(0, x0 - band)
    oy0 = max(0, y0 - band)
    ox1 = min(width, x1 + band)
    oy1 = min(height, y1 + band)
    ix0, iy0 = max(0, x0), max(0, y0)
    ix1, iy1 = min(width, x1), min(height, y1)
    if ox1 <= ox0 or oy1 <= oy0:
        return 0.0
    ring = std_map[oy0:oy1, ox0:ox1] > threshold
    hole = std_map[iy0:iy1, ix0:ix1] > threshold
    total = int(ring.size)
    covered = int(hole.size)
    if total <= covered:
        return 0.0
    return float((int(ring.sum()) - covered) / (total - covered))


def _classify_region(shape: tuple[int, int], solidity: float, touches_edge: bool) -> str:
    """粗略判断区域类型。"""
    height, width = shape
    area = float(height * width)
    if solidity < 0.32:
        return "logo"
    if area > 0.04 and touches_edge:
        return "photo"
    if solidity < 0.62:
        return "illustration"
    return "photo"


# ------------------------------------------------------------------- 线条
@dataclass
class _Line:
    """一条装饰线。"""

    bbox_px: PixelRect
    horizontal: bool
    color: tuple[int, int, int]


def _detect_lines(mask: np.ndarray, array: np.ndarray, options: LayeringOptions) -> list[_Line]:
    """在纹理掩膜中找出细长的装饰线。

    做法：对掩膜做水平 / 垂直方向的长核开运算，保留「细长且连续」的结构。
    开运算会把渐变淡出的长线打断成很多碎段，因此再沿线的方向做一次膨胀
    把碎段接回去，最后用紧致包围盒判定长宽比与厚度。

    关键判据是**颜色均匀**：设计稿里的装饰线、卡片描边是单一颜色，而照片
    内部的树枝、屋檐虽然也是细长结构，颜色却杂乱无章。没有这一条会把
    照片内容当成线条剔掉，反而在照片上打出孔洞。

    线条若不从图片候选里剔除，会把页面上所有内容串成一个组件。
    """
    length = max(8, options.line_min_length)
    merge = max(6, options.line_min_length // 2)
    results: list[_Line] = []
    for horizontal, axis in ((True, 1), (False, 0)):
        opened = _dilate_1d(_erode_1d(mask, length, axis), length, axis)
        merged = _dilate_1d(opened, merge, axis)
        labels, count = _label_components(merged)
        if count == 0:
            continue
        for stat in _component_boxes(labels, count):
            x0, y0, x1, y1 = stat["bbox"]
            component = labels[y0:y1, x0:x1] == stat["id"]
            ink = component & opened[y0:y1, x0:x1]
            if not ink.any():
                continue
            pixels = array[y0:y1, x0:x1][ink]
            if pixels.size == 0:
                continue
            color = np.median(pixels, axis=0)
            spread = float(np.linalg.norm(pixels.astype(np.float32) - color, axis=1).mean())
            if spread > options.line_color_std:
                # 颜色杂乱 → 照片内部结构，不是设计线条
                continue
            tight = _tight_bbox(mask, (x0, y0, x1, y1), horizontal)
            if tight is None:
                continue
            tx0, ty0, tx1, ty1 = tight
            bw, bh = tx1 - tx0, ty1 - ty0
            if horizontal:
                if bw < options.line_min_length or bh > options.line_max_thickness:
                    continue
                if bw / max(bh, 1) < options.line_min_aspect:
                    continue
            else:
                if bh < options.line_min_length or bw > options.line_max_thickness:
                    continue
                if bh / max(bw, 1) < options.line_min_aspect:
                    continue
            # 护栏：线两侧若充满纹理，说明它长在照片内部（屋檐、栏杆等），
            # 不是设计线条，不应从图片候选里剔除
            if _surrounding_density(mask, (tx0, ty0, tx1, ty1)) > options.line_context_density:
                continue
            results.append(
                _Line(
                    bbox_px=(tx0, ty0, tx1, ty1),
                    horizontal=horizontal,
                    color=tuple(int(round(v)) for v in color),  # type: ignore[arg-type]
                )
            )
    return results


def _surrounding_density(mask: np.ndarray, bbox: PixelRect, band: int = 8) -> float:
    """线条周围一圈的纹理密度。"""
    height, width = mask.shape
    x0, y0, x1, y1 = bbox
    ox0, oy0 = max(0, x0 - band), max(0, y0 - band)
    ox1, oy1 = min(width, x1 + band), min(height, y1 + band)
    if ox1 <= ox0 or oy1 <= oy0:
        return 1.0
    region = mask[oy0:oy1, ox0:ox1]
    total = float(region.size)
    inner = region[
        max(0, y0 - oy0) : max(0, y1 - oy0),
        max(0, x0 - ox0) : max(0, x1 - ox0),
    ]
    covered = float(inner.size)
    if total <= covered:
        return 1.0
    return float((float(region.sum()) - covered) / (total - covered))


def _tight_bbox(
    mask: np.ndarray, bbox: PixelRect, horizontal: bool
) -> PixelRect | None:
    """在给定范围内求真实像素的紧致包围盒。"""
    height, width = mask.shape
    x0, y0, x1, y1 = bbox
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(width, x1), min(height, y1)
    if x1 <= x0 or y1 <= y0:
        return None
    sub = mask[y0:y1, x0:x1]
    if not sub.any():
        return None
    rows = np.nonzero(sub.any(axis=1))[0]
    cols = np.nonzero(sub.any(axis=0))[0]
    if horizontal:
        return (x0 + int(cols.min()), y0 + int(rows.min()), x0 + int(cols.max()) + 1, y0 + int(rows.max()) + 1)
    return (x0 + int(cols.min()), y0 + int(rows.min()), x0 + int(cols.max()) + 1, y0 + int(rows.max()) + 1)


def _lines_to_shapes(lines: list[_Line], page_rect: Rect) -> list[ShapeElement]:
    """把装饰线转成 PPT 连接符。"""
    shapes: list[ShapeElement] = []
    for line in lines:
        x0, y0, x1, y1 = line.bbox_px
        if line.horizontal:
            start = (float(page_rect[0] + x0), float(page_rect[1] + (y0 + y1) / 2))
            end = (float(page_rect[0] + x1), float(page_rect[1] + (y0 + y1) / 2))
            rect = (float(page_rect[0] + x0), float(page_rect[1] + y0), float(page_rect[0] + x1), float(page_rect[1] + y1))
        else:
            start = (float(page_rect[0] + (x0 + x1) / 2), float(page_rect[1] + y0))
            end = (float(page_rect[0] + (x0 + x1) / 2), float(page_rect[1] + y1))
            rect = (float(page_rect[0] + x0), float(page_rect[1] + y0), float(page_rect[0] + x1), float(page_rect[1] + y1))
        color = tuple(channel / 255.0 for channel in line.color)
        shapes.append(
            ShapeElement(
                kind="line",
                rect=rect,
                color=color,  # type: ignore[arg-type]
                fill=None,
                width=1.0,
                start=start,
                end=end,
            )
        )
    return shapes


# ------------------------------------------------------------- 平涂色块
def _boundary_points(mask: np.ndarray) -> np.ndarray:
    """掩膜的边界像素坐标（x, y）。"""
    padded = np.pad(mask, 1, mode="constant", constant_values=False)
    interior = (
        padded[:-2, 1:-1]
        & padded[2:, 1:-1]
        & padded[1:-1, :-2]
        & padded[1:-1, 2:]
    )
    edge = mask & ~interior
    ys, xs = np.nonzero(edge)
    if ys.size == 0:
        return np.zeros((0, 2), dtype=np.float64)
    return np.stack([xs, ys], axis=1).astype(np.float64)


def _fit_polygon(
    mask: np.ndarray, max_sides: int = 5, tolerance: float = 2.8, iterations: int = 400
) -> list[tuple[float, float]] | None:
    """用 RANSAC 直线拟合多边形轮廓（用于斜切等不规则色块）。"""
    boundary = _boundary_points(mask)
    if boundary.shape[0] < 16:
        return None
    rng = np.random.default_rng(20240517)
    remaining = boundary
    lines: list[tuple[float, float, float]] = []
    total = boundary.shape[0]
    for _ in range(max_sides):
        n = remaining.shape[0]
        if n < max(16, int(total * 0.08)):
            break
        idx = rng.integers(0, n, size=(iterations, 2))
        p1 = remaining[idx[:, 0]]
        p2 = remaining[idx[:, 1]]
        delta = p2 - p1
        norm = np.hypot(delta[:, 0], delta[:, 1])
        keep = norm > 5.0
        p1, p2, delta, norm = p1[keep], p2[keep], delta[keep], norm[keep]
        if p1.shape[0] == 0:
            break
        a = delta[:, 1]
        b = -delta[:, 0]
        c = -(a * p1[:, 0] + b * p1[:, 1])
        length = np.hypot(a, b)
        length[length == 0] = 1.0
        a, b, c = a / length, b / length, c / length
        distances = np.abs(
            a[:, None] * remaining[None, :, 0]
            + b[:, None] * remaining[None, :, 1]
            + c[:, None]
        )
        inliers = distances < tolerance
        best = int(np.argmax(inliers.sum(axis=1)))
        count = int(inliers[best].sum())
        if count < max(16, int(total * 0.12)):
            break
        lines.append((float(a[best]), float(b[best]), float(c[best])))
        remaining = remaining[~inliers[best]]
    if len(lines) < 3:
        return None

    # 按角度排序，相邻直线求交得到顶点
    ordered = sorted(lines, key=lambda item: np.arctan2(item[1], item[0]))
    points: list[tuple[float, float]] = []
    for index, current in enumerate(ordered):
        following = ordered[(index + 1) % len(ordered)]
        det = current[0] * following[1] - current[1] * following[0]
        if abs(det) < 1e-6:
            return None
        x = (current[2] * following[1] - following[2] * current[1]) / det
        y = (current[0] * following[2] - following[0] * current[2]) / det
        points.append((float(x), float(y)))
    if len(points) < 3:
        return None
    return points


def _polygon_area(points: list[tuple[float, float]]) -> float:
    total = 0.0
    for index in range(len(points)):
        x0, y0 = points[index]
        x1, y1 = points[(index + 1) % len(points)]
        total += x0 * y1 - x1 * y0
    return abs(total) / 2.0


def _detect_flat_regions(
    array: np.ndarray,
    std_map: np.ndarray,
    exclude: np.ndarray,
    page_rect: Rect,
    options: LayeringOptions,
    skip_color: tuple[int, int, int] | None = None,
    taken: list[PixelRect] | None = None,
) -> list[ShapeElement]:
    """按「局部平坦度」分割背景，还原平涂色块。

    按颜色分桶对**渐变**面板天然脆弱：一张浅蓝渐变卡片会被量化打碎成几千个
    碎片（实测 p7 的浅蓝面板碎成 4476 块）。改用平坦度分割就稳得多：

    * 平坦区 = 局部标准差低于阈值的像素（纯色、渐变、留白都算）
    * 色块与背景之间的**边界**是硬边 → 局部标准差高 → 天然把两者分开
    * 连通域后得到一块块平坦区域，再判断颜色是否足够均匀：
      纯色 → 原生矩形；渐变 → 保留在背景图里，不做原生对象（否则会丢掉渐变）
    """
    height, width = array.shape[:2]
    page_area = float(width * height)
    min_area = max(600.0, page_area * options.block_min_area_ratio)
    max_area = page_area * 0.35

    flat = std_map < options.flat_threshold
    flat &= ~exclude
    # 去掉 JPEG 噪点造成的麻点
    flat = _dilate(_erode(flat, 1), 1)

    labels, count = _label_components(flat)
    shapes: list[ShapeElement] = []
    if taken is None:
        taken = []
    for stat in _component_boxes(labels, count):
        x0, y0, x1, y1 = stat["bbox"]
        area = stat["area"]
        if area < min_area or area > max_area:
            continue
        bw, bh = x1 - x0, y1 - y0
        if bw < 10 or bh < 10:
            continue
        component = labels[y0:y1, x0:x1] == stat["id"]
        pixels = array[y0:y1, x0:x1][component]
        if pixels.size == 0:
            continue
        color = np.median(pixels, axis=0)
        spread = float(np.linalg.norm(pixels.astype(np.float32) - color, axis=1).mean())
        if spread > options.block_color_std:
            # 渐变色块：保留在背景图里，避免用纯色原生矩形覆盖掉渐变
            continue
        if skip_color is not None:
            distance = float(
                np.linalg.norm(
                    color.astype(np.float32) - np.array(skip_color, dtype=np.float32)
                )
            )
            if distance < 18.0:
                continue
        filled = component | _enclosed_holes(component)
        solidity = float(filled.sum()) / float(bw * bh)
        if solidity < options.block_min_solidity:
            continue
        if any(_overlaps((x0, y0, x1, y1), other) for other in taken):
            continue
        taken.append((x0, y0, x1, y1))
        shapes.append(
            ShapeElement(
                kind="rect",
                rect=(
                    float(page_rect[0] + x0),
                    float(page_rect[1] + y0),
                    float(page_rect[0] + x1),
                    float(page_rect[1] + y1),
                ),
                color=None,
                fill=tuple(float(c / 255.0) for c in color),  # type: ignore[arg-type]
                width=0.0,
            )
        )
    return shapes


def _detect_flat_blocks(
    array: np.ndarray,
    exclude: np.ndarray,
    page_rect: Rect,
    options: LayeringOptions,
    skip_color: tuple[int, int, int] | None = None,
    taken: list[PixelRect] | None = None,
) -> list[ShapeElement]:
    """在背景区域中找出平涂色块，还原为 PPT 原生矩形 / 自由曲线。

    平涂色块在画册里非常常见（信息卡片、面板、色块标题）。它们是纯色的，
    因此可以完全用 PPT 原生形状还原，用户可以直接改填充色和圆角。

    色块上通常压着文字，文字会把色块「啃」出洞，导致实心度被低估。因此
    实心度要用「补洞之后」的面积计算，而颜色纯度要用「补洞之前」的像素
    计算（文字像素不属于色块）。
    """
    height, width = array.shape[:2]
    page_area = float(width * height)
    min_area = max(600.0, page_area * options.block_min_area_ratio)
    max_area = page_area * 0.35

    quantized = (array >> 4).astype(np.int32)
    keys = (quantized[:, :, 0] << 8) | (quantized[:, :, 1] << 4) | quantized[:, :, 2]
    usable = ~exclude
    keys = np.where(usable, keys, -1)
    counts = np.bincount(keys[usable].ravel(), minlength=4096)
    order = np.argsort(-counts)[: options.block_max_levels]

    shapes: list[ShapeElement] = []
    if taken is None:
        taken = []
    for key in order:
        if key < 0:
            continue
        mask = (keys == key) & usable
        total = int(mask.sum())
        if total < min_area:
            continue
        ys, xs = np.nonzero(mask)
        bx0, bx1 = int(xs.min()), int(xs.max()) + 1
        by0, by1 = int(ys.min()), int(ys.max()) + 1
        labels, count = _label_components(mask)
        for stat in _component_boxes(labels, count):
            x0, y0, x1, y1 = stat["bbox"]
            area = stat["area"]
            if area < min_area or area > max_area:
                continue
            bw, bh = x1 - x0, y1 - y0
            if bw < 10 or bh < 10:
                continue
            component = labels[y0:y1, x0:x1] == stat["id"]
            pixels = array[y0:y1, x0:x1][component]
            if pixels.size == 0:
                continue
            color = np.median(pixels, axis=0)
            spread = float(np.linalg.norm(pixels.astype(np.float32) - color, axis=1).mean())
            if spread > options.block_color_std:
                continue
            if skip_color is not None:
                distance = float(
                    np.linalg.norm(color.astype(np.float32) - np.array(skip_color, dtype=np.float32))
                )
                if distance < 18.0:
                    # 与背景同色，画出来也没有意义
                    continue
            # 补洞后再算实心度（色块上的文字把色块啃出了洞）
            filled = component | _enclosed_holes(component)
            solidity = float(filled.sum()) / float(bw * bh)
            if solidity >= options.block_min_solidity:
                if not _boundary_contrast_ok(
                    array, (x0, y0, x1, y1), color, options.block_boundary_contrast
                ):
                    continue
                if any(_overlaps((x0, y0, x1, y1), other) for other in taken):
                    continue
                taken.append((x0, y0, x1, y1))
                shapes.append(
                    ShapeElement(
                        kind="rect",
                        rect=(
                            float(page_rect[0] + x0),
                            float(page_rect[1] + y0),
                            float(page_rect[0] + x1),
                            float(page_rect[1] + y1),
                        ),
                        color=None,
                        fill=tuple(float(c / 255.0) for c in color),  # type: ignore[arg-type]
                        width=0.0,
                    )
                )
            else:
                points = _fit_polygon(filled, options.freeform_max_sides)
                if not points:
                    continue
                poly_area = _polygon_area(points)
                if poly_area < page_area * options.freeform_min_area_ratio:
                    continue
                if not (0.6 * area <= poly_area <= 1.6 * area):
                    continue
                if any(_overlaps((x0, y0, x1, y1), other) for other in taken):
                    continue
                taken.append((x0, y0, x1, y1))
                shapes.append(
                    ShapeElement(
                        kind="freeform",
                        rect=(
                            float(page_rect[0] + x0),
                            float(page_rect[1] + y0),
                            float(page_rect[0] + x1),
                            float(page_rect[1] + y1),
                        ),
                        color=None,
                        fill=tuple(float(c / 255.0) for c in color),  # type: ignore[arg-type]
                        width=0.0,
                        points=[
                            (float(page_rect[0] + px), float(page_rect[1] + py))
                            for px, py in points
                        ],
                    )
                )
    return shapes


def _overlaps(a: PixelRect, b: PixelRect) -> bool:
    """两个像素矩形是否相交。"""
    return not (a[2] <= b[0] or b[2] <= a[0] or a[3] <= b[1] or b[3] <= a[1])


def _boundary_contrast_ok(
    array: np.ndarray, bbox: PixelRect, color: np.ndarray, threshold: float
) -> bool:
    """色块边界与外部的对比度是否足够（用于排除渐变色带）。"""
    height, width = array.shape[:2]
    x0, y0, x1, y1 = bbox
    band = 3
    outer_parts = []
    ox0, oy0 = max(0, x0 - band), max(0, y0 - band)
    ox1, oy1 = min(width, x1 + band), min(height, y1 + band)
    if oy0 < y0:
        outer_parts.append(array[oy0:y0, ox0:ox1].reshape(-1, 3))
    if oy1 > y1:
        outer_parts.append(array[y1:oy1, ox0:ox1].reshape(-1, 3))
    if ox0 < x0:
        outer_parts.append(array[y0:y1, ox0:x0].reshape(-1, 3))
    if ox1 > x1:
        outer_parts.append(array[y0:y1, x1:ox1].reshape(-1, 3))
    if not outer_parts:
        return True
    outer = np.concatenate(outer_parts).astype(np.float32)
    distance = np.linalg.norm(outer - color.reshape(1, 3), axis=1)
    return float(np.median(distance)) >= threshold


# ------------------------------------------------------------------ 主流程
def split_page(
    bitmap: bytes,
    page_rect: Rect,
    text_rects: list[PixelRect],
    options: LayeringOptions | None = None,
) -> LayeringResult:
    """把一页整页位图分解成「背景 + 独立图片区域 + 原生形状」。"""
    options = options or LayeringOptions()
    result = LayeringResult()

    try:
        with Image.open(io.BytesIO(bitmap)) as handle:
            image = handle.convert("RGB")
    except Exception:
        return result

    array = np.array(image)
    height, width = array.shape[:2]
    page_area = float(width * height)

    # --- 1) 纹理能量图 -------------------------------------------------
    gray = array.astype(np.float64).mean(axis=2)
    std_map = _local_std(gray, options.texture_radius)
    texture_mask = std_map > options.texture_threshold

    # 去掉页面最外圈：整页 JPEG 在边缘会留下一圈接缝伪纹理，
    # 它会把页面四周的所有内容连成一个组件。
    trim = max(0, int(options.edge_trim))
    if trim > 0:
        texture_mask[:trim, :] = False
        texture_mask[-trim:, :] = False
        texture_mask[:, :trim] = False
        texture_mask[:, -trim:] = False

    # --- 2) 文字掩膜 ----------------------------------------------------
    # 注意：文字**保留在候选掩膜里**。若先把压在平底上的文字从掩膜中剔除，
    # Logo 里的字就会与 Logo 图形断开，Logo 区域不再包含自己的文字，
    # 后面就无法判断「这段字属于图片还是属于页面」。
    # 纯文字块改在连通域阶段用「文字占比」剔除。
    text_mask = _mask_from_rects(text_rects, (width, height))

    # 图片候选 = 全部纹理区域（含文字）
    candidate = texture_mask.copy()

    # --- 3) 线条分离 ----------------------------------------------------
    # 先把卡片描边、分隔线这类设计线条识别出来并从图片候选中剔除，
    # 否则它们会把页面上所有内容串成一个组件。
    # 这里**不做**形态学开运算：开运算会把细笔画的 Logo（如「阿里云」）
    # 整个抹掉，反而丢内容。断连接交给线条检测来做。
    lines = _detect_lines(candidate, array, options)
    line_mask = np.zeros((height, width), dtype=bool)
    for line in lines:
        x0, y0, x1, y1 = line.bbox_px
        line_mask[y0:y1, x0:x1] = True
    candidate = candidate & ~line_mask

    # --- 4) 形态学 + 连通域 --------------------------------------------
    candidate = _dilate(candidate, options.close_radius)
    candidate = _erode(candidate, options.close_radius)
    labels, count = _label_components(candidate)
    if count == 0:
        result.stats = {"regions": 0, "reason": "no_texture"}
        return result

    candidate = _fill_holes_per_component(candidate, labels, count)
    labels, count = _label_components(candidate)

    # --- 5) 筛出真正的图片区域 ------------------------------------------
    min_area = max(400.0, page_area * options.min_area_ratio)
    regions: list[ImageRegion] = []
    for stat in _component_boxes(labels, count):
        ys, xs = np.nonzero(labels == stat["id"])
        if ys.size == 0:
            continue
        if stat["area"] < min_area:
            continue
        component = labels == stat["id"]
        x0, x1 = int(xs.min()), int(xs.max()) + 1
        y0, y1 = int(ys.min()), int(ys.max()) + 1
        bbox_w, bbox_h = x1 - x0, y1 - y0
        if bbox_w < 6 or bbox_h < 6:
            continue
        # 面积判据：细笔画 Logo（如「阿里云」）墨迹很少，但包围盒够大；
        # 只按墨迹面积过滤会把这类 Logo 漏掉。因此取「墨迹面积」与
        # 「包围盒面积」两条判据的宽松者，同时用绝对下限挡住噪点。
        bbox_area = float(bbox_w * bbox_h)
        big_enough = stat["area"] >= min_area or (
            bbox_area >= page_area * options.min_bbox_area_ratio
            and stat["area"] >= options.min_ink_area
        )
        if not big_enough:
            continue
        # 文字块判定：文字框几乎盖满整个包围盒 → 这是标题/正文，不是图片。
        # 不能用「组件内文字像素占比」：Logo 是「图形 + 文字」，
        # 文字像素占比可能过半，会把真 Logo 误删。
        text_coverage = float(text_mask[y0:y1, x0:x1].sum()) / float(bbox_w * bbox_h)
        if text_coverage > options.text_block_ratio:
            continue
        solidity = stat["area"] / float(bbox_w * bbox_h)
        touches_edge = x0 == 0 or y0 == 0 or x1 == width or y1 == height
        regions.append(
            ImageRegion(
                bbox_px=(x0, y0, x1, y1),
                rect=(
                    page_rect[0] + x0,
                    page_rect[1] + y0,
                    page_rect[0] + x1,
                    page_rect[1] + y1,
                ),
                mask=component,
                kind=_classify_region((bbox_h, bbox_w), solidity, touches_edge),
            )
        )

    regions.sort(key=lambda item: -item.area)
    if len(regions) > options.max_regions:
        regions = regions[: options.max_regions]

    if not regions:
        result.stats = {"regions": 0, "reason": "all_filtered"}
        return result

    union = np.zeros((height, width), dtype=bool)
    for region in regions:
        union |= region.mask

    # --- 6) 文字归属：图内文字 vs 页面文字 -------------------------------
    # 按面积重叠判定归属：文字中心常常落在 Logo 的白色缝隙上，只看中心点会漏判。
    region_text_rects: dict[int, list[tuple[int, PixelRect]]] = {
        index: [] for index in range(len(regions))
    }
    artwork_text: set[int] = set()
    region_pixels = [float(region.mask.sum()) for region in regions]
    for index, rect in enumerate(text_rects):
        rx0, ry0, rx1, ry1 = rect
        text_area = float(max(rx1 - rx0, 1) * max(ry1 - ry0, 1))
        overlaps: list[tuple[int, float]] = []
        for region_index, region in enumerate(regions):
            x0, y0, x1, y1 = region.bbox_px
            ix0, iy0 = max(rx0, x0), max(ry0, y0)
            ix1, iy1 = min(rx1, x1), min(ry1, y1)
            if ix1 <= ix0 or iy1 <= iy0:
                continue
            # 按面积重叠判定归属：文字中心常常落在 Logo 的白色缝隙上，
            # 只看中心点会漏判
            overlap = float(region.mask[iy0:iy1, ix0:ix1].sum()) / text_area
            if overlap >= options.text_overlap_ratio:
                overlaps.append((region_index, overlap))
        if not overlaps:
            continue
        if len(overlaps) > 1:
            # 一行文字同时压到多个图片区域 → 这是 OCR 把相邻 Logo 的字
            # 误合并成了一行。这类文字属于画面内容，留在图里即可。
            artwork_text.add(index)
            continue
        region_index, _overlap = max(overlaps, key=lambda item: item[1])
        region_text_rects[region_index].append((index, rect))
        if region_pixels[region_index] <= page_area * options.artwork_area_ratio:
            # 小尺寸图片（Logo、配图）里的字属于画面内容，留在图里即可
            artwork_text.add(index)

    covered_text: set[int] = set()
    for items in region_text_rects.values():
        for index, _rect in items:
            covered_text.add(index)
    background_text = [
        rect for index, rect in enumerate(text_rects) if index not in covered_text
    ]

    # --- 7) 纯色背景判定（先做，便于跳过与背景同色的色块） -------------
    flat_color = None
    sample_mask = ~(union | text_mask)
    if sample_mask.sum() > 100:
        candidate_color = _detect_flat_background(array, sample_mask, options.flat_tolerance)
        if candidate_color is not None:
            ratio = _flat_coverage(array, sample_mask, candidate_color, options.flat_tolerance)
            if ratio >= options.flat_ratio:
                flat_color = candidate_color

    # --- 8) 背景层的原生形状 --------------------------------------------
    exclude = union | text_mask | line_mask
    taken: list[PixelRect] = []
    shapes = _detect_flat_regions(
        array, std_map, exclude, page_rect, options, flat_color, taken
    )
    shapes.extend(
        _detect_flat_blocks(array, exclude, page_rect, options, flat_color, taken)
    )
    shapes.extend(_lines_to_shapes(lines, page_rect))

    # --- 9) 背景重建 ----------------------------------------------------
    covered = float(union.sum()) / page_area
    background_array: np.ndarray | None = array.astype(np.float32)

    if covered >= 0.985:
        # 整页就是一张满版图，不再需要背景层
        background_array = None
    elif options.rebuild_background:
        background_array = _pull_push_fill(background_array, union)
    else:
        background_array[union] = _dominant_color(array, ~union)

    if background_array is not None and background_text:
        background_array = _erase_flat_text(background_array, array, background_text)

    if flat_color is not None:
        result.flat_color = flat_color
        result.background = None
    elif background_array is not None:
        result.background = _encode_png(background_array)
    else:
        result.background = None

    result.regions = regions
    result.shapes = shapes
    result.region_text_rects = region_text_rects
    result.artwork_text = artwork_text
    result.background_text_rects = background_text
    result.stats = {
        "regions": len(regions),
        "shapes": len(shapes),
        "lines": len(lines),
        "image_coverage": round(covered, 4),
        "flat_background": flat_color is not None,
        "artwork_text": len(artwork_text),
        "background_text": len(background_text),
    }
    return result


def _dominant_color(array: np.ndarray, mask: np.ndarray) -> np.ndarray:
    """掩膜区域内的主色。"""
    if not mask.any():
        return np.array([255.0, 255.0, 255.0], dtype=np.float32)
    pixels = array[mask]
    return np.median(pixels, axis=0).astype(np.float32)


def _detect_flat_background(
    array: np.ndarray, sample_mask: np.ndarray, tolerance: float
) -> tuple[int, int, int] | None:
    """判断背景是否可以用单一纯色表示。"""
    pixels = array[sample_mask]
    if pixels.size == 0:
        return None
    color = np.median(pixels, axis=0)
    distance = np.linalg.norm(pixels.astype(np.float32) - color, axis=1)
    coverage = float((distance <= tolerance).mean())
    if coverage >= 0.90:
        return tuple(int(round(value)) for value in color)  # type: ignore[return-value]
    return None


def _flat_coverage(
    array: np.ndarray,
    sample_mask: np.ndarray,
    color: tuple[int, int, int],
    tolerance: float,
) -> float:
    pixels = array[sample_mask]
    if pixels.size == 0:
        return 0.0
    reference = np.array(color, dtype=np.float32)
    distance = np.linalg.norm(pixels.astype(np.float32) - reference, axis=1)
    return float((distance <= tolerance).mean())


def _erase_flat_text(
    background: np.ndarray, original: np.ndarray, rects: list[PixelRect]
) -> np.ndarray:
    """把压在平底上的文字从背景里抹掉（用外圈采样色填充）。"""
    out = background.copy()
    height, width = out.shape[:2]
    for rect in rects:
        x0, y0, x1, y1 = _rect_to_pixels(rect, (width, height))
        if x1 <= x0 or y1 <= y0:
            continue
        band = 8
        parts = []
        if y0 - band >= 0:
            parts.append(original[max(0, y0 - band) : y0, x0:x1].reshape(-1, 3))
        if y1 + band <= height:
            parts.append(original[y1 : min(height, y1 + band), x0:x1].reshape(-1, 3))
        if x0 - band >= 0:
            parts.append(original[y0:y1, max(0, x0 - band) : x0].reshape(-1, 3))
        if x1 + band <= width:
            parts.append(original[y0:y1, x1 : min(width, x1 + band)].reshape(-1, 3))
        if not parts:
            continue
        ring = np.concatenate(parts).astype(np.float32)
        color = np.median(ring, axis=0)
        out[y0:y1, x0:x1] = color.reshape(1, 1, 3)
    return out


def _encode_png(array: np.ndarray) -> bytes:
    """把浮点数组编码为 PNG。"""
    data = np.clip(array, 0, 255).astype(np.uint8)
    buffer = io.BytesIO()
    Image.fromarray(data).save(buffer, format="PNG", optimize=True)
    return buffer.getvalue()


def export_region(
    array: np.ndarray,
    region: ImageRegion,
    text_rects: list[PixelRect],
    options: LayeringOptions | None = None,
) -> tuple[bytes, int, int]:
    """把一个区域导出成带 alpha 的 PNG。

    区域内压在图片上的文字会被局部修复（推拉填补），避免文字在 PPT 里
    出现两份。返回 ``(png 字节, 像素宽, 像素高)``。
    """
    options = options or LayeringOptions()
    x0, y0, x1, y1 = region.bbox_px
    crop = array[y0:y1, x0:x1].astype(np.float32).copy()
    height, width = crop.shape[:2]

    holes = np.zeros((height, width), dtype=bool)
    for rect in text_rects:
        rx0, ry0, rx1, ry1 = rect
        ix0, iy0 = max(0, rx0 - x0), max(0, ry0 - y0)
        ix1, iy1 = min(width, rx1 - x0), min(height, ry1 - y0)
        if ix1 > ix0 and iy1 > iy0:
            holes[iy0:iy1, ix0:ix1] = True
    if holes.any():
        crop = _pull_push_fill(crop, holes)

    alpha = _feather(region.mask[y0:y1, x0:x1], options.feather_radius)
    rgba = np.dstack(
        [
            np.clip(crop, 0, 255).astype(np.uint8),
            (np.clip(alpha, 0, 1) * 255).astype(np.uint8),
        ]
    )
    buffer = io.BytesIO()
    Image.fromarray(rgba, mode="RGBA").save(buffer, format="PNG", optimize=True)
    return buffer.getvalue(), width, height
