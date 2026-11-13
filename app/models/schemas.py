"""接口数据结构。"""

from __future__ import annotations

from pydantic import BaseModel, Field


class PageStat(BaseModel):
    """单页元素统计。"""

    index: int = Field(..., description="页码，从 1 开始")
    width: float = Field(..., description="页面宽度（点）")
    height: float = Field(..., description="页面高度（点）")
    text_lines: int = 0
    text_chars: int = 0
    images: int = 0
    shapes: int = 0
    tables: int = 0
    art_layers: int = 0


class ConvertFileStat(BaseModel):
    """单个文件的转换统计。"""

    page_count: int = 0
    text_lines: int = 0
    text_chars: int = 0
    images: int = 0
    shapes: int = 0
    tables: int = 0
    art_layers: int = 0
    editable_ratio: float = Field(default=0.0, description="可编辑元素占全部元素的比例")


class ConvertedFile(BaseModel):
    """转换结果条目。"""

    file_id: str
    source_name: str
    output_name: str
    download_url: str
    inspect_url: str
    stats: ConvertFileStat


class ConvertResponse(BaseModel):
    """批量转换响应。"""

    files: list[ConvertedFile] = []
    errors: list[str] = []
    elapsed_seconds: float = 0.0
