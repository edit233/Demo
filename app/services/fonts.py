"""PDF 字体名 → PowerPoint 字体名映射。

PDF 内嵌字体通常是 ``ABCDEF+Arial-BoldMT`` 这类子集名，需要剥离前缀后
再做归一化匹配。中文 PDF 常见的宋体/黑体/雅黑等直接映射为操作系统上
存在的标准中文字体，保证在 PowerPoint / WPS 中可正常编辑。
"""

from __future__ import annotations

import re

_SUBSET_PREFIX = re.compile(r"^[A-Za-z]{6}\+")

# 命中顺序即优先级，越靠前优先级越高
_CJK_RULES: tuple[tuple[str, str], ...] = (
    ("simhei", "黑体"),
    ("heiti", "黑体"),
    ("msyh", "微软雅黑"),
    ("yahei", "微软雅黑"),
    ("microsoftyahei", "微软雅黑"),
    ("dengxian", "等线"),
    ("fangsong", "仿宋"),
    ("kaiti", "楷体"),
    ("simsun", "宋体"),
    ("songti", "宋体"),
    ("stsong", "宋体"),
    ("sourcehansans", "微软雅黑"),
    ("notosanscjksc", "微软雅黑"),
    ("notosanssc", "微软雅黑"),
)

_LATIN_RULES: tuple[tuple[str, str], ...] = (
    ("timesnewroman", "Times New Roman"),
    ("times", "Times New Roman"),
    ("arial", "Arial"),
    ("helvetica", "Arial"),
    ("calibri", "Calibri"),
    ("cambria", "Cambria"),
    ("georgia", "Georgia"),
    ("verdana", "Verdana"),
    ("tahoma", "Tahoma"),
    ("segoeui", "Segoe UI"),
    ("couriernew", "Courier New"),
    ("courier", "Courier New"),
    ("consolas", "Consolas"),
)


def strip_subset_prefix(name: str) -> str:
    """剥离 PDF 子集前缀，如 ``ABCDEF+``。"""
    return _SUBSET_PREFIX.sub("", name or "")


def has_cjk(text: str) -> bool:
    """判断文本是否包含中日韩字符。"""
    return any(
        "\u2e80" <= ch <= "\u9fff"
        or "\uac00" <= ch <= "\ud7af"
        or "\uff00" <= ch <= "\uffef"
        for ch in text or ""
    )


def _normalize(name: str) -> str:
    text = strip_subset_prefix(name).lower()
    for ch in " -,_()":
        text = text.replace(ch, "")
    return text


def resolve_font(font_name: str, sample_text: str = "") -> tuple[str, str]:
    """把 PDF 字体名解析为 ``(西文字体, 中文字体)``。

    源字体为中文字体时，西文与中文统一使用同一字体，避免中英混排时
    字重、粗细不一致；无法识别时按文本是否含中文回退。
    """
    normalized = _normalize(font_name)

    cjk_font = ""
    for key, value in _CJK_RULES:
        if key in normalized:
            cjk_font = value
            break

    if cjk_font:
        return cjk_font, cjk_font

    latin = "Arial"
    for key, value in _LATIN_RULES:
        if key in normalized:
            latin = value
            break

    east_asian = "微软雅黑" if has_cjk(sample_text) else latin
    return latin, east_asian
