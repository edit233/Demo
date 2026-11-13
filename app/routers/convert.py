"""HTTP 接口。"""

from __future__ import annotations

import asyncio
import json
import time
import uuid
from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from app.core.config import settings
from app.models.schemas import ConvertedFile, ConvertFileStat, ConvertResponse
from app.services.converter import ConversionError, ConvertOptions, convert_pdf
from app.services.layering import LayeringOptions

router = APIRouter(prefix="/api", tags=["转换"])

PPTX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
ALLOWED_SUFFIXES = {".pdf"}


def _safe_stem(name: str) -> str:
    stem = Path(name).stem or "document"
    return "".join(ch for ch in stem if ch not in '\\/:*?"<>|').strip() or "document"


def _find_output(file_id: str) -> Path | None:
    if not file_id.isalnum():
        return None
    directory = settings.output_dir / file_id
    if not directory.is_dir():
        return None
    for item in sorted(directory.glob("*.pptx")):
        return item
    return None


@router.get("/health", summary="健康检查")
async def health() -> dict:
    """服务健康检查。"""
    return {"status": "ok", "version": settings.version}


@router.post("/convert", response_model=ConvertResponse, summary="批量转换 PDF 为 PPTX")
async def convert(
    files: list[UploadFile] = File(..., description="待转换的 PDF 文件"),
    keep_art_layer: bool | None = Form(None, description="是否保留复杂矢量为独立图片层"),
    detect_tables: bool | None = Form(None, description="是否检测并还原表格"),
    enable_ocr: bool | None = Form(None, description="是否对图片型页面启用 OCR 文字还原"),
    enable_layering: bool | None = Form(
        None, description="是否把整页位图拆成背景 + 独立图片对象 + 原生形状"
    ),
):
    """上传一个或多个 PDF，返回转换结果与下载地址。

    可选的开关未提供时，使用服务配置中的默认值。
    """
    if not files:
        raise HTTPException(status_code=400, detail="请至少选择一个 PDF 文件")

    results: list[ConvertedFile] = []
    errors: list[str] = []
    started = time.perf_counter()

    for upload in files:
        source_name = upload.filename or "document.pdf"
        suffix = Path(source_name).suffix.lower()
        if suffix not in ALLOWED_SUFFIXES:
            errors.append(f"{source_name}：仅支持 PDF 格式")
            continue

        data = await upload.read()
        await upload.close()
        if not data:
            errors.append(f"{source_name}：文件内容为空")
            continue
        if len(data) > settings.max_upload_bytes:
            errors.append(
                f"{source_name}：文件超过 {settings.max_upload_mb} MB 上限"
            )
            continue

        file_id = uuid.uuid4().hex
        task_dir = settings.upload_dir / file_id
        task_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = task_dir / f"source{suffix}"
        pdf_path.write_bytes(data)

        output_dir = settings.output_dir / file_id
        output_dir.mkdir(parents=True, exist_ok=True)

        options = ConvertOptions(
            keep_art_layer=(
                settings.keep_art_layer if keep_art_layer is None else keep_art_layer
            ),
            detect_tables=(
                settings.detect_tables if detect_tables is None else detect_tables
            ),
            render_scale=settings.render_scale,
            enable_ocr=(settings.enable_ocr if enable_ocr is None else enable_ocr),
            ocr_min_confidence=settings.ocr_min_confidence,
            ocr_erase_text=settings.ocr_erase_text,
            enable_layering=(
                settings.enable_layering if enable_layering is None else enable_layering
            ),
            layering=LayeringOptions(
                texture_threshold=settings.layer_texture_threshold,
                min_area_ratio=settings.layer_min_area_ratio,
                max_regions=settings.layer_max_regions,
                rebuild_background=settings.layer_rebuild_background,
            ),
        )
        try:
            result = await asyncio.to_thread(
                convert_pdf, pdf_path, output_dir, source_name, options
            )
        except ConversionError as exc:
            errors.append(f"{source_name}：{exc}")
            continue
        except Exception as exc:  # pragma: no cover - 兜底
            errors.append(f"{source_name}：转换异常 {exc}")
            continue

        summary = result.report.get("summary", {})
        results.append(
            ConvertedFile(
                file_id=file_id,
                source_name=source_name,
                output_name=result.output_path.name,
                download_url=f"/api/download/{file_id}",
                inspect_url=f"/api/inspect/{file_id}",
                stats=ConvertFileStat(
                    page_count=result.report.get("page_count", 0),
                    text_lines=summary.get("text_lines", 0),
                    text_chars=summary.get("text_chars", 0),
                    images=summary.get("images", 0),
                    shapes=summary.get("shapes", 0),
                    tables=summary.get("tables", 0),
                    art_layers=summary.get("art_layers", 0),
                    editable_ratio=summary.get("editable_ratio", 0.0),
                    background_layers=summary.get("background_layers", 0),
                    decomposed_pages=summary.get("decomposed_pages", 0),
                    full_page_image_pages=summary.get("full_page_image_pages", 0),
                ),
            )
        )

    if not results:
        raise HTTPException(
            status_code=400, detail="；".join(errors) or "转换失败，请检查上传的文件"
        )

    return ConvertResponse(
        files=results,
        errors=errors,
        elapsed_seconds=round(time.perf_counter() - started, 2),
    )


@router.get("/download/{file_id}", summary="下载转换结果")
async def download(file_id: str):
    """下载生成的 PPTX 文件。"""
    path = _find_output(file_id)
    if path is None:
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    return FileResponse(
        path,
        media_type=PPTX_MEDIA_TYPE,
        filename=path.name,
        headers={"Cache-Control": "no-store"},
    )


@router.get("/inspect/{file_id}", summary="查看可编辑性验证报告")
async def inspect(file_id: str):
    """返回每个页面还原出的对象清单，用于验证可编辑性。"""
    if not file_id.isalnum():
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    report_path = settings.output_dir / file_id / "report.json"
    if not report_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    return JSONResponse(
        json.loads(report_path.read_text(encoding="utf-8")),
        headers={"Cache-Control": "no-store"},
    )
