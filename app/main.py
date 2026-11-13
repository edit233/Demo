"""FastAPI 应用入口。"""

from __future__ import annotations

import asyncio
from contextlib import asynccontextmanager, suppress
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.core.config import settings
from app.routers import convert as convert_router
from app.services.cleanup import cleanup_loop

STATIC_DIR = Path(__file__).resolve().parent / "static"


@asynccontextmanager
async def lifespan(app: FastAPI):
    """启动时准备目录并开启定时清理任务。"""
    settings.ensure_dirs()
    task = asyncio.create_task(cleanup_loop())
    try:
        yield
    finally:
        task.cancel()
        with suppress(asyncio.CancelledError):
            await task


app = FastAPI(
    title=settings.app_name,
    version=settings.version,
    description=(
        "把 PDF 页面还原为可编辑的 PowerPoint 对象："
        "文字、图片、矩形/线条/椭圆、表格均以原生对象输出，"
        "复杂矢量内容保留为独立图片对象。"
    ),
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=list(settings.allowed_origins),
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(convert_router.router)
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", include_in_schema=False)
async def index():
    """前端上传页面。"""
    return FileResponse(STATIC_DIR / "index.html")
