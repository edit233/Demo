"""临时文件清理：按保留时长删除上传与输出目录。"""

from __future__ import annotations

import asyncio
import shutil
import time
from pathlib import Path

from app.core.config import settings


def cleanup_expired(now: float | None = None) -> int:
    """删除过期的任务目录，返回删除数量。"""
    current = now if now is not None else time.time()
    deadline = current - settings.retention_hours * 3600
    removed = 0
    for root in (settings.upload_dir, settings.output_dir):
        if not root.exists():
            continue
        for item in root.iterdir():
            try:
                if not item.is_dir():
                    continue
                if item.stat().st_mtime < deadline:
                    shutil.rmtree(item, ignore_errors=True)
                    removed += 1
            except OSError:
                continue
    return removed


async def cleanup_loop(interval: int = 1800) -> None:
    """后台定时清理。"""
    while True:
        await asyncio.sleep(interval)
        try:
            cleanup_expired()
        except Exception:  # pragma: no cover - 后台任务不因异常退出
            continue
