# PDF → 可编辑 PPTX 转换服务

把「整页图片型 PDF」还原为**由独立对象组成**的 PowerPoint 演示文稿：文字是可选中、可编辑、可改字号的文本框，照片是可选中、可移动、可缩放的图片对象，矢量图形、线条、表格则还原为 PPT 原生对象。

服务基于 **FastAPI + PyMuPDF + python-pptx + RapidOCR**，支持批量上传 PDF、在线下载转换后的 PPTX，并附带逐页的可编辑性验证报告。

> 本仓库同时是**题目二「PDF 转换为可编辑 PPTX」**的完整交付物：源码 + README + 3 套素材的转换结果 + 可编辑性验证材料 + 容器化部署文件。

> **本分支为部署分支（`deploy`）**：只保留运行服务所需的最小文件集（34 个文件 / 0.4 MB），不含测试素材、转换结果与验证材料，以减小镜像体积与克隆开销。完整交付材料见 [`main` 分支](https://github.com/edit233/Demo/tree/main)。

---

## 一、效果对比

左侧为原始 PDF 页面，右侧为转换后 PPTX 的近似渲染（由 `scripts/render_preview.py` 重绘）。3 套素材的对比图与预览图属于验证材料，本部署分支未收录，见 [`main` 分支](https://github.com/edit233/Demo/tree/main) 的 `outputs/` 目录。

底图上的文字已被**抹除并替换为可编辑文本框**，因此页面上不存在「一张整页截图糊在 PPT 里」的情况。

---

## 二、为什么需要 OCR

本项目的 3 套测试素材均为**整页位图形式的 PDF**：经 PyMuPDF 解析后每页只有 1 张 JPEG（有效密度 72 DPI），**文本层与矢量对象数量均为 0**。也就是说，PDF 里根本不存在「可以直接抽取的文字」。

因此服务采用两条并行路径：

| 页面类型 | 判定依据 | 处理方式 |
| --- | --- | --- |
| **图片型页面**（本项目素材） | 无文本层，且存在覆盖整页的位图 | OCR 识别文字 → 抹除底图文字 → 生成独立文本框 |
| **结构化页面**（带文本层的 PDF） | 存在文本层 | 直接抽取文字、矢量、图片、表格并还原为原生对象 |

图片型页面除了 OCR 文字之外，页面中其余的图形与照片仍以**独立图片对象**保留（可单独选中、移动、缩放），文字则逐行生成文本框。

---

## 三、元素还原情况说明

| PDF 中的元素 | 还原方式 | 是否可编辑 | 说明 |
| --- | --- | --- | --- |
| 文字（图片型页面） | OCR 识别 → PPT 文本框 | ✅ 可双击编辑 | 字号、颜色由像素测量估算，字符间距由列投影测得 |
| 文字（带文本层的 PDF） | 抽取文本层 → PPT 文本框 | ✅ 可双击编辑 | 字号、字重、斜体、颜色、字体族来自 PDF 原生属性 |
| 照片 / 位图 | `add_picture` 独立图片对象 | ✅ 可移动、缩放 | 按目标框比例自动裁剪，避免拉伸变形 |
| 矩形 | PPT 矩形自选图形 | ✅ 可改填充/描边 | 填充色、透明度、虚线样式均还原 |
| 椭圆 / 圆 | PPT 椭圆自选图形 | ✅ 可改填充/描边 | 同上 |
| 直线 | PPT 连接符（`add_connector`） | ✅ 可拖动端点 | 线宽、颜色、虚线还原 |
| 表格 | PPT 原生表格（`add_table`） | ✅ 可改单元格 | 列宽、行高、边框、单元格底色还原，并套用「无样式无网格」表格样式 |
| 复杂矢量路径（渐变、蒙版、任意多边形等） | 以透明 PNG 作为**独立图片对象**保留 | ✅ 可移动、缩放 | 无法结构化的部分按题目要求保留为图片对象 |
| 文本层中的行内多样式 | 同一文本框内多个 run | ✅ 可局部编辑 | 每个 run 保留各自样式 |

### 转换流程

```
PDF ──► PyMuPDF 解析 ──┬─► 文本层 ──────────────► 文本框 / 表格 / 形状 / 连接符 / 图片
                       │
                       └─► 整页位图 ──► RapidOCR ─┬─► 文字 ─► 文本框（+ 抹除底图文字）
                                                   └─► 其余图形 ─► 独立图片对象
                                                              │
                                                              ▼
                                                   python-pptx 原生对象写入 PPTX
```

### 坐标换算

PyMuPDF 返回的页面坐标已经归一化为**左上角原点、Y 轴向下**，与 PowerPoint 一致，因此换算只剩单位换算：

\[
\text{EMU} = \text{点} \times 12700
\]

幻灯片尺寸按首页 PDF 页面尺寸换算（上限 56 英寸，符合 PowerPoint 限制）。

---

## 四、无法还原的内容

以下内容受限于输入格式本身，属于**客观上无法还原**的部分，已在 README 中明确列出：

1. **图片型 PDF 中的原始文字属性**——素材被栅格化，字号、字体、字重、颜色都只能由像素反推，属于估算值；识别结果可能与原设计稿存在细微差异（尤其是艺术字、生僻字、极小字号）。
2. **图片型 PDF 中的表格**——表格在像素层面已经不存在边框与单元格结构，只能识别为若干行文字，**无法还原为 PPT 原生表格**。带文本层的 PDF 才能还原原生表格。
3. **字间距**——PDF 的位图不携带字符间距信息；本项目通过像素**列投影**直接测量相邻字形之间的空白来还原字距（宽字距的中文标题可较好还原），但字形本身的宽度取决于原设计所用字体，与最终渲染字体存在差异。
4. **文字中的空格**——OCR 对字间空格不敏感（如 `AND SHARES A BETTER LIFE` 可能识别为 `AND SHARESA BETTER LIFE`）。
5. **复杂矢量内容**——渐变、阴影、蒙版、任意路径、SmartArt、图表对象等统一以透明 PNG 图片对象保留（可独立选中移动，但内部不可编辑）。这符合题目「无法结构化的复杂视觉内容可用独立图片对象保留」的要求。
6. **字体**——只做字体族名称映射（如 `SimSun` → 宋体、`MicrosoftYaHei` → 微软雅黑），**不内嵌字体文件**，最终显示效果取决于打开 PPT 的机器上是否安装对应字体。
7. **母版、主题、动画、演讲者备注、页码域、公式对象（OMML）**——不转换。
8. **页面尺寸不一致的 PDF**——以首页尺寸作为整份文档的幻灯片尺寸。

---

## 五、快速开始

### 5.1 本地运行

环境要求：**Python ≥ 3.12**（推荐 3.12 / 3.13）。

```bash
# 1. 安装依赖（推荐使用 uv）
uv sync
# 或者：pip install -r requirements.txt

# 2. 启动服务
uv run uvicorn app.main:app --host 0.0.0.0 --port 8000
# 或者：python -m uvicorn app.main:app --reload
```

启动后访问：

| 地址 | 说明 |
| --- | --- |
| <http://localhost:8000/> | 前端上传页面（拖拽上传 → 转换 → 下载） |
| <http://localhost:8000/docs> | Swagger UI 接口文档 |
| <http://localhost:8000/api/health> | 健康检查 |

> OCR 模型（PP-OCRv4）随 `rapidocr-onnxruntime` 包分发，**完全离线运行**，不需要下载模型、不调用任何外部 API。

### 5.2 Docker 部署

```bash
# 方式一：docker compose
docker compose up -d --build
docker compose logs -f

# 方式二：docker run
docker build -t pdf2pptx-service:1.0.0 .
docker run -d -p 8000:8000 -v pdf2pptx-data:/data pdf2pptx-service:1.0.0
```

容器内监听 `0.0.0.0:8000`，上传文件写入 `/data/uploads`，转换结果写入 `/data/outputs`（均可用卷挂载到宿主机）。

### 5.3 批量转换本地素材

```bash
python scripts/convert_materials.py                # 转换 tests/materials 下全部素材
python scripts/convert_materials.py 测试1.pdf       # 只转换指定文件
python scripts/convert_materials.py --no-ocr       # 关闭 OCR（仅走结构化路径）
```

> 脚本固定读取仓库根目录下的 `tests/materials/`。本部署分支不包含该目录（测试素材见 [`main` 分支](https://github.com/edit233/Demo/tree/main)），使用前请自行把待转换的 PDF 放入 `tests/materials/`。

---

## 六、配置项

所有配置项均通过环境变量覆盖，统一前缀 `PDF2PPTX_`（见 `app/core/config.py`）：

| 环境变量 | 默认值 | 说明 |
| --- | --- | --- |
| `PDF2PPTX_APP_NAME` | `PDF 转可编辑 PPTX 服务` | 服务标题 |
| `PDF2PPTX_VERSION` | `1.0.0` | 版本号 |
| `PDF2PPTX_UPLOAD_DIR` | `./uploads` | 上传目录 |
| `PDF2PPTX_OUTPUT_DIR` | `./outputs` | 转换结果目录 |
| `PDF2PPTX_MAX_UPLOAD_MB` | `100` | 单文件体积上限（MB） |
| `PDF2PPTX_MAX_PAGES` | `300` | 单份 PDF 页数上限 |
| `PDF2PPTX_RETENTION_HOURS` | `6` | 上传文件与结果的保留时长（小时） |
| `PDF2PPTX_RENDER_SCALE` | `2.0` | 复杂矢量层渲染倍率 |
| `PDF2PPTX_KEEP_ART_LAYER` | `1` | 是否保留复杂矢量为透明图片层 |
| `PDF2PPTX_DETECT_TABLES` | `1` | 是否检测表格 |
| `PDF2PPTX_ENABLE_OCR` | `1` | 是否对图片型页面启用 OCR |
| `PDF2PPTX_OCR_MIN_CONFIDENCE` | `0.5` | OCR 置信度阈值 |
| `PDF2PPTX_OCR_ERASE_TEXT` | `1` | 是否从底图中抹除已识别的文字 |
| `PDF2PPTX_ALLOWED_ORIGINS` | `*` | CORS 允许的来源，逗号分隔 |

> 本仓库**不包含任何密钥**。服务不调用任何第三方付费 API，OCR 与 PDF 处理全部在本地完成。

---

## 七、API

### `POST /api/convert`

批量上传 PDF 并转换。`multipart/form-data`，字段名 `files`，可一次传多个文件。

```bash
curl -X POST http://localhost:8000/api/convert \
  -F "files=@/path/to/测试1.pdf" \
  -F "files=@/path/to/测试2.pdf"
```

响应：

```json
{
  "files": [
    {
      "file_id": "3f1c…",
      "source_name": "测试1.pdf",
      "output_name": "测试1.pptx",
      "download_url": "/api/download/3f1c…",
      "inspect_url": "/api/inspect/3f1c…",
      "stats": {
        "page_count": 10,
        "text_lines": 299,
        "text_chars": 2931,
        "images": 10,
        "shapes": 0,
        "tables": 0,
        "art_layers": 0,
        "editable_ratio": 1.0
      }
    }
  ],
  "errors": [],
  "elapsed_seconds": 53.1
}
```

### `GET /api/download/{file_id}`

下载转换后的 PPTX（`application/vnd.openxmlformats-officedocument.presentationml.presentation`）。

### `GET /api/inspect/{file_id}`

返回逐页元素明细的可编辑性验证报告（JSON）。该报告在转换时实时生成并写入输出目录，其结构与验证材料中的 `report.json` 一致。

### `GET /api/health`

健康检查。

---

## 八、目录结构

```
.
├── app/
│   ├── main.py                  # FastAPI 入口、静态资源挂载、定时清理
│   ├── core/config.py           # 配置（环境变量 PDF2PPTX_*）
│   ├── models/schemas.py        # 请求/响应模型
│   ├── routers/convert.py       # HTTP 接口
│   ├── services/
│   │   ├── pdf_parser.py        # PDF 解析：文本 / 图片 / 矢量 / 表格
│   │   ├── ocr.py               # OCR：检测、行合并、墨迹测量、底图抹字
│   │   ├── fonts.py             # PDF 字体名 → PPT 字体名映射
│   │   ├── pptx_builder.py      # 原生对象写入（文本框/图片/形状/连接符/表格）
│   │   ├── converter.py         # 转换编排与报告生成
│   │   └── cleanup.py           # 过期文件清理
│   └── static/                  # 前端页面（原生 HTML/CSS/JS）
├── scripts/                     # 批量转换与验证工具（需自备素材，见 5.3）
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
└── pyproject.toml
```

---

## 九、可编辑性验证材料

本部署分支不收录验证材料（`verification/` 与 `outputs/`），相关材料保存在 [`main` 分支](https://github.com/edit233/Demo/tree/main)：

| 材料 | 位置（`main` 分支） | 内容 |
| --- | --- | --- |
| 验证报告 | `verification/可编辑性验证报告.md` | 总体指标、OOXML 原生元素统计、逐页对象数、文本抽样 |
| 汇总数据 | `verification/summary.json` | 机器可读的统计结果 |
| 逐页对象清单 | `outputs/<素材>/objects.txt` | 每页的对象类型、位置、文本内容 |
| 元素明细报告 | `outputs/<素材>/report.json` | 每页每个元素的类型、包围盒、字号、颜色 |
| 页面预览图 | `outputs/<素材>/preview/slide_*.png` | 转换后每一页的渲染图 |
| 原图对比图 | `outputs/<素材>/compare/compare_*.png` | 原始 PDF 页面与转换结果并排对比 |
| 原始页面图 | `outputs/<素材>/compare/source/source_*.png` | 原始 PDF 每一页的渲染图，作为对比图的上半部分 |

自行复验（需自行准备素材后执行）：

```bash
# 查看某个 PPTX 里到底有哪些对象
python scripts/verify_pptx.py outputs/测试1/测试1.pptx

# 重新生成全部验证材料（对象清单、预览图、对比图、报告）
python scripts/build_verification.py

# 压缩验证材料图片（量化调色板 + 无损压缩，仓库体积可减 60%）
python scripts/optimize_images.py

# HTTP 层端到端验证（自动拉起服务，走一遍上传→转换→下载→报告）
# 依赖 tests/materials/ 下的素材，需自行放入
python scripts/e2e_http_check.py
```

**判定标准**：如果某页只由一张覆盖整页的图片构成，那才是「把 PNG 当背景」的不合格结果。验证报告中的「退化整页图」一列统计的就是这种情况，本项目输出为 **0 页**。

---

## 十、3 套素材的转换结果

3 套测试素材的转换结果（PPTX、预览图、对比图）保存在 [`main` 分支](https://github.com/edit233/Demo/tree/main)；本部署分支仅保留运行服务所需的最小文件集。转换结果的逐页统计数据如下：

| 素材 | 页数 | 文本框 | 图片对象 | 原生可编辑对象（非图片） | 对象总数 | 可编辑占比 | 耗时 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 测试1.pdf | 10 | 299 | 10 | 299 | 309 | 100% | 53.5s |
| 测试2.pdf | 10 | 195 | 10 | 195 | 205 | 100% | 46.2s |
| 测试3.pdf | 10 | 151 | 10 | 151 | 161 | 100% | 31.0s |

> **统计口径**：「原生可编辑对象」指非图片的原生对象（文本框 / 形状 / 表格 / 连接符）数量，可在 PowerPoint 中双击直接编辑；「对象总数」再加上图片对象——图片同样可独立选中、移动、缩放，因此也计入可编辑对象。「可编辑占比」= 可独立编辑的对象数 ÷ 页面对象总数，100% 意味着**没有任何一页退化成「整页背景图」**（若整页只放一张截图，占比会很低）。

---

## 十一、使用的开源项目与工具

| 项目 | 用途 | 许可证 |
| --- | --- | --- |
| [FastAPI](https://github.com/tiangolo/fastapi) | Web 框架、接口与 OpenAPI 文档 | MIT |
| [uvicorn](https://github.com/encode/uvicorn) | ASGI 服务器 | BSD-3 |
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | PDF 解析、页面渲染、图形结构提取 | AGPL-3.0 |
| [python-pptx](https://github.com/scanny/python-pptx) | 生成 PPTX 原生对象 | MIT |
| [RapidOCR](https://github.com/RapidAI/RapidOCR) / `rapidocr-onnxruntime` | 图片型 PDF 的文字识别（PP-OCRv4，中英文） | Apache-2.0 |
| [ONNX Runtime](https://github.com/microsoft/onnxruntime) | OCR 模型推理 | MIT |
| [Pillow](https://github.com/python-pillow/Pillow) | 图像处理 | MIT-CMU |
| [NumPy](https://github.com/numpy/numpy) | 数值计算 | BSD-3 |
| [uv](https://github.com/astral-sh/uv) | 依赖管理与虚拟环境 | Apache-2.0 / MIT |

开发过程中使用了 AI 编程工具辅助代码编写；核心的解析、坐标换算、OOXML 写入、OCR 行合并与墨迹测量逻辑均由本仓库代码实现。

---

## 十二、性能说明

- 10 页图片型 PDF 的转换耗时约 30~55 秒（主要开销在 OCR 的检测与识别）。
- OCR 引擎为**进程内单例**并通过线程锁串行访问，避免多请求同时初始化模型。
- 转换在 `asyncio.to_thread` 中执行，不阻塞事件循环；上传文件与结果会在 `PDF2PPTX_RETENTION_HOURS`（默认 6 小时）后自动清理。

## 十三、许可

本仓库代码采用 MIT 许可。依赖的 PyMuPDF 采用 AGPL-3.0 许可，若用于闭源商业产品请自行评估合规性（可替换为 `pypdfium2` 等宽松许可的渲染后端）。
