"""服务全局配置。

所有配置项均可通过环境变量覆盖，统一前缀 ``PDF2PPTX_``。
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

# app/core/config.py -> app/core -> app -> 项目根目录
BASE_DIR = Path(__file__).resolve().parents[2]


def _env_str(name: str, default: str) -> str:
    value = os.getenv(name)
    return value if value not in (None, "") else default


def _env_int(name: str, default: int) -> int:
    try:
        return int(float(_env_str(name, str(default))))
    except (TypeError, ValueError):
        return default


def _env_float(name: str, default: float) -> float:
    try:
        return float(_env_str(name, str(default)))
    except (TypeError, ValueError):
        return default


def _env_bool(name: str, default: bool) -> bool:
    return _env_str(name, "1" if default else "0").strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class Settings:
    """运行期配置。"""

    app_name: str = _env_str("PDF2PPTX_APP_NAME", "PDF 转可编辑 PPTX 服务")
    version: str = _env_str("PDF2PPTX_VERSION", "1.0.0")
    base_dir: Path = BASE_DIR
    upload_dir: Path = Path(_env_str("PDF2PPTX_UPLOAD_DIR", str(BASE_DIR / "uploads")))
    output_dir: Path = Path(_env_str("PDF2PPTX_OUTPUT_DIR", str(BASE_DIR / "outputs")))
    max_upload_mb: int = _env_int("PDF2PPTX_MAX_UPLOAD_MB", 100)
    max_pages: int = _env_int("PDF2PPTX_MAX_PAGES", 300)
    retention_hours: int = _env_int("PDF2PPTX_RETENTION_HOURS", 6)
    render_scale: float = _env_float("PDF2PPTX_RENDER_SCALE", 2.0)
    keep_art_layer: bool = _env_bool("PDF2PPTX_KEEP_ART_LAYER", True)
    detect_tables: bool = _env_bool("PDF2PPTX_DETECT_TABLES", True)
    enable_ocr: bool = _env_bool("PDF2PPTX_ENABLE_OCR", True)
    ocr_min_confidence: float = _env_float("PDF2PPTX_OCR_MIN_CONFIDENCE", 0.5)
    ocr_erase_text: bool = _env_bool("PDF2PPTX_OCR_ERASE_TEXT", True)
    allowed_origins: tuple[str, ...] = tuple(
        item.strip()
        for item in _env_str("PDF2PPTX_ALLOWED_ORIGINS", "*").split(",")
        if item.strip()
    )

    @property
    def max_upload_bytes(self) -> int:
        """单文件上传体积上限（字节）。"""
        return self.max_upload_mb * 1024 * 1024

    def ensure_dirs(self) -> None:
        """确保上传/输出目录存在。"""
        self.upload_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)


settings = Settings()
