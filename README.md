# PDF → 可编辑 PPTX 转换服务

把「整页图片型 PDF」还原为**由独立对象组成**的 PowerPoint 演示文稿：文字是可选中、可编辑、可改字号的文本框，照片是可选中、可移动、可缩放的**独立图片对象**，矢量图形、线条、表格则还原为 PPT 原生对象。

服务基于 **FastAPI + PyMuPDF + python-pptx + RapidOCR**，支持批量上传 PDF、在线下载转换后的 PPTX，并附带逐页的可编辑性验证报告。

> 本仓库同时是**题目二「PDF 转换为可编辑 PPTX」**的完整交付物：源码 + README + 3 套素材的转换结果 + 可编辑性验证材料 + 容器化部署文件。

> **核心能力：图层分解**。图片型 PDF 的每一页只有一张整页位图，普通做法是把整页图当背景贴进去（照片被封死）。本项目实现了**像素级图层分解**：把整页位图拆成「重建背景层 + 多个独立图片对象 + 原生矢量形状 + 文本框」，照片、Logo、图标、插画都能被单独选中、移动、缩放。算法原理与完整踩坑记录见 [`docs/图层分解方案.md`](docs/图层分解方案.md)。

---

## 一、效果对比

左侧为原始 PDF 页面，右侧为转换后 PPTX 的近似渲染（由 `scripts/render_preview.py` 重绘）：

| 素材 | 第 1 页对比 |
| --- | --- |
| 测试1 | [对比图](outputs/测试1/compare/compare_01.png) |
| 测试2 | [对比图](outputs/测试2/compare/compare_01.png) |
| 测试3 | [对比图](outputs/测试3/compare/compare_01.png) |

底图上的文字已被**抹除并替换为可编辑文本框**，照片、Logo、图标被**抠出为独立图片对象**，因此页面上不存在「一张整页截图糊在 PPT 里」的情况。

---

## 二、为什么需要 OCR 与图层分解

本项目的 3 套测试素材均为**整页位图形式的 PDF**：经 PyMuPDF 解析后每页只有 1 张 JPEG（有效密度 72 DPI），**文本层与矢量对象数量均为 0**。也就是说，PDF 里根本不存在「可以直接抽取的文字」，也不存在「可以直接抽取的内嵌图片」。

这带来两个必须解决的问题：

| 问题 | 结论 | 本项目的做法 |
| --- | --- | --- |
| 文字怎么变可编辑 | 只能靠 OCR | RapidOCR 识别 → 抹除底图文字 → 逐行生成文本框 |
| 图片怎么变独立对象 | 只能靠像素级分解 | 把整页位图拆成「背景 + 独立图片对象 + 原生形状」 |

因此服务采用两条并行路径：

| 页面类型 | 判定依据 | 处理方式 |
| --- | --- | --- |
| **图片型页面**（本项目素材） | 无文本层，且存在覆盖整页的位图 | OCR 识别文字 → 抹除底图文字 → 生成独立文本框 → 图层分解出独立图片与原生形状 |
| **结构化页面**（带文本层的 PDF） | 存在文本层 | 直接抽取文字、矢量、图片、表格并还原为原生对象 |

### 图层分解做了什么

把一页整页位图当作一张「平面拼贴」来理解，图层分解要回答的问题是：**画面里哪些像素属于同一块独立素材？**

答案由四个像素信号共同给出：

| 信号 | 含义 | 用途 |
| --- | --- | --- |
| **纹理能量**（局部标准差） | 局部区域的色彩变化剧烈程度 | 照片、插画、Logo 内部纹理丰富；纯色背景与平涂色块近乎平坦 |
| **文字感知** | OCR 文本框的空间分布 | 文字区域不参与图片区域检测，避免把标题切进照片里 |
| **线条响应** | 细长、色泽均匀、周边平坦的带状结构 | 装饰线分离为 PPT 连接符，不留在图片里 |
| **平坦度** | 局部方差足够小且颜色高度一致 | 平涂色块还原为 PPT 原生矩形 |

处理流程：

```
整页位图
   │
   ├─► 纹理能量图 ─► 线条分离 ─► 闭运算 / 连通域标记 ─► 逐组件填洞
   │                                                    │
   │                                    ┌───────────────┴───────────────┐
   │                                    ▼                               ▼
   │                         图片区域（照片/插画/Logo/图标）      平坦区（平涂色块）
   │                                    │                               │
   │                                    ▼                               ▼
   │                          裁剪 + Alpha 羽化 + 边缘修复         原生矩形 / 自由曲线
   │                                    │                               │
   └────────────────────────► 背景重建（金字塔 push-pull 填补） ◄─────┘
                                        │
                                        ▼
                    底色矩形 → 背景位图 → 矢量层 → 其它形状 → 图片对象 → 表格 → 文本
```

### 图层叠放顺序（z-order）

| 层 | 内容 | 生成方式 |
| --- | --- | --- |
| 1 | 纯色底 | PPT 原生矩形 |
| 2 | 重建背景位图 | 金字塔 push-pull 填补（照片抠走后露出的底） |
| 3 | 背景层的矢量形状 | 原生矩形 / 自由曲线（斜切块） |
| 4 | 其它矢量形状 | 原生矩形 / 椭圆 / 连接符 |
| 5 | 图片对象 | 独立 PNG（可单独选中、移动、缩放） |
| 6 | 表格 | PPT 原生表格 |
| 7 | 文本 | 文本框 |

### 三个关键设计决策

1. **背景层「剔除并重建」**：照片被抠走后露出的是原始底图中的背景，本项目用金字塔 push-pull 插值把空洞补齐，而不是简单填白，因此背景的渐变与色调得以保留。
2. **矢量还原范围**：平涂色块 → PPT 原生矩形；装饰线 → PPT 连接符；斜切多边形 → PPT **自由曲线**（`build_freeform`）；渐变色块保守地留在背景位图里，不用纯色矩形去覆盖，以免丢掉渐变。
3. **图内文字保留在图里**：当一张独立图片的面积占页面比例小于 8% 时，其中的文字被判定为「图内文字」（例如 Logo 里的中英文），只保留在图片像素中，不再额外生成重复的文本框。

### 稳健性设计

图层分解是启发式算法，**任何一页分解失败都不会让整份文档转换失败**：该页自动降级为「整页擦字单图」，并在日志中输出告警。报告会通过 `decomposed_pages` 与 `full_page_image_pages` 两个指标暴露降级情况。

---

## 三、元素还原情况说明

| PDF 中的元素 | 还原方式 | 是否可编辑 | 说明 |
| --- | --- | --- | --- |
| 文字（图片型页面） | OCR 识别 → PPT 文本框 | ✅ 可双击编辑 | 字号、颜色由像素测量估算，字符间距由列投影测得 |
| 文字（带文本层的 PDF） | 抽取文本层 → PPT 文本框 | ✅ 可双击编辑 | 字号、字重、斜体、颜色、字体族来自 PDF 原生属性 |
| **照片（图片型页面）** | **图层分解 → `add_picture` 独立图片对象** | ✅ 可移动、缩放 | 照片从整页图中抠出，背景被重建，不再被封死在底图里 |
| **Logo / 图标 / 插画（图片型页面）** | **图层分解 → `add_picture` 独立图片对象** | ✅ 可移动、缩放 | 带 Alpha 通道，边缘做了羽化；图内文字保留在图中不重复生成文本框 |
| **平涂色块（图片型页面）** | **图层分解 → PPT 矩形自选图形** | ✅ 可改填充/描边 | 纯色区域用原生形状承载，可在 PPT 里直接改色 |
| **斜切块（图片型页面）** | **图层分解 → PPT 自由曲线** | ✅ 可改填充/描边 | RANSAC 直线拟合出多边形顶点，用 `build_freeform` 还原 |
| **装饰线（图片型页面）** | **图层分解 → PPT 连接符** | ✅ 可拖动端点 | 细长、色泽均匀且周边平坦的带状结构 |
| 照片 / 位图（带文本层的 PDF） | `add_picture` 独立图片对象 | ✅ 可移动、缩放 | 按目标框比例自动裁剪，避免拉伸变形 |
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
                                                   │
                                                   └─► 图层分解 ─┬─► 背景重建（背景位图 / 纯色矩形）
                                                                ├─► 图片对象（独立 PNG）
                                                                ├─► 原生矩形 / 自由曲线 / 连接符
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
4. **文字中的空格**——OCR 对字间空格不敏感（如 `AND SHARES A BETTER LIFE` 可能识别为 `AND SHARESA BETTER LIFE`）。相邻 Logo 的文字也可能被 OCR 误合并为一行（如「招商银行平安银行」），此类问题属于 OCR 引擎固有误差。
5. **复杂矢量内容**——渐变、阴影、蒙版、任意路径、SmartArt、图表对象等统一以透明 PNG 图片对象保留（可独立选中移动，但内部不可编辑）。这符合题目「无法结构化的复杂视觉内容可用独立图片对象保留」的要求。
6. **字体**——只做字体族名称映射（如 `SimSun` → 宋体、`MicrosoftYaHei` → 微软雅黑），**不内嵌字体文件**，最终显示效果取决于打开 PPT 的机器上是否安装对应字体。
7. **母版、主题、动画、演讲者备注、页码域、公式对象（OMML）**——不转换。
8. **页面尺寸不一致的 PDF**——以首页尺寸作为整份文档的幻灯片尺寸。

### 图层分解本身的已知限制

以下是**图层分解算法**（像素级启发式）的客观边界，不是实现缺陷：

1. **满版插画天然不可再分**——例如测试3 第 1 页的熊猫城市夜景是一张铺满全页的插画，画面内部没有可分辨的素材边界，只能作为一张整页图片对象保留。这是内容本身决定的，任何像素算法都无法凭空切分一张整体插画。
2. **背景重建是近似值**——照片抠走后由金字塔 push-pull 插值填补，斜切硬边会被柔和化，局部可能残留轻微的重建痕迹。
3. **源图分辨率上限**——素材 PDF 的内嵌位图只有 72 DPI（852×1212 px 对应 852×1212 pt，1 px = 1 pt）。把独立图片对象放大超过原始尺寸时会发糊，这是输入素材的分辨率限制，不是导出策略问题。
4. **渐变与纹理混合的色块保守处理**——既不是纯色、又有明显纹理的大面积区域，为避免破坏观感，会被保守地留在背景位图中，而不是强行转成纯色矩形。
5. **细窄的装饰元素可能被并入相邻区域**——宽度低于像素阈值的发丝线、极细描边，无法保证被单独分离为连接符。
6. **分解失败时自动降级**——某一页分解异常时，该页退化为「整页擦字单图」而不是整体转换失败；`full_page_image_pages` 指标即为这种情况的计数。

完整的算法原理、参数说明与 10 条实测踩坑记录见 [`docs/图层分解方案.md`](docs/图层分解方案.md)。

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
python scripts/convert_materials.py --no-layer     # 关闭图层分解（整页位图不拆分）
```

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
| `PDF2PPTX_ENABLE_LAYERING` | `1` | 是否把整页位图拆分为「背景 + 独立图片对象 + 原生形状」 |
| `PDF2PPTX_LAYER_TEXTURE_THRESHOLD` | `9.0` | 图层分解：纹理能量阈值（局部标准差），越大越保守 |
| `PDF2PPTX_LAYER_MIN_AREA_RATIO` | `0.0015` | 图层分解：图片区域最小面积占页面比例 |
| `PDF2PPTX_LAYER_MAX_REGIONS` | `36` | 图层分解：单页最多拆出的图片对象数 |
| `PDF2PPTX_LAYER_REBUILD_BACKGROUND` | `1` | 图层分解：是否把背景重建为整页位图（关闭则背景只保留纯色矩形与原生形状） |
| `PDF2PPTX_ALLOWED_ORIGINS` | `*` | CORS 允许的来源，逗号分隔 |

> 本仓库**不包含任何密钥**。服务不调用任何第三方付费 API，OCR 与 PDF 处理全部在本地完成。

---

## 七、API

### `POST /api/convert`

批量上传 PDF 并转换。`multipart/form-data`，字段名 `files`，可一次传多个文件。

| 表单字段 | 类型 | 默认 | 说明 |
| --- | --- | --- | --- |
| `files` | 文件 | 必填 | 待转换的 PDF，可多个 |
| `enable_ocr` | 布尔 | 取服务端配置 | 图片型页面的 OCR 文字还原 |
| `keep_art_layer` | 布尔 | 取服务端配置 | 复杂矢量是否保留为独立图片层 |
| `detect_tables` | 布尔 | 取服务端配置 | 是否检测并还原表格 |
| `enable_layering` | 布尔 | 取服务端配置 | 是否把整页位图拆成背景 + 独立图片对象 + 原生形状 |

```bash
curl -X POST http://localhost:8000/api/convert \
  -F "files=@tests/materials/测试1.pdf" \
  -F "files=@tests/materials/测试2.pdf" \
  -F "enable_layering=1"
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
        "text_lines": 257,
        "text_chars": 2458,
        "images": 108,
        "shapes": 28,
        "tables": 0,
        "art_layers": 0,
        "background_layers": 9,
        "decomposed_pages": 10,
        "full_page_image_pages": 0,
        "editable_ratio": 1.0
      }
    }
  ],
  "errors": [],
  "elapsed_seconds": 104.9
}
```

### 新增统计指标

| 指标 | 含义 | 本项目实测 |
| --- | --- | --- |
| `background_layers` | 重建出的背景位图层数 | 9 / 10 / 8 |
| `decomposed_pages` | 成功完成图层分解的页数 | 10 / 10 / 9 |
| `full_page_image_pages` | 退化为整页单图的页数（**越低越好**） | 0 / 0 / 0 |

### `GET /api/download/{file_id}`

下载转换后的 PPTX（`application/vnd.openxmlformats-officedocument.presentationml.presentation`）。

### `GET /api/inspect/{file_id}`

返回逐页元素明细的可编辑性验证报告（JSON），结构与 `outputs/<素材名>/report.json` 一致。

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
│   │   ├── ocr.py               # OCR：检测、行合并、栏位切分、墨迹测量、底图抹字
│   │   ├── layering.py          # 图层分解：整页位图 → 背景 + 独立图片 + 原生形状
│   │   ├── fonts.py             # PDF 字体名 → PPT 字体名映射
│   │   ├── pptx_builder.py      # 原生对象写入（文本框/图片/形状/连接符/表格/自由曲线）
│   │   ├── converter.py         # 转换编排与报告生成
│   │   └── cleanup.py           # 过期文件清理
│   └── static/                  # 前端页面（原生 HTML/CSS/JS）
├── scripts/
│   ├── convert_materials.py     # 批量转换测试素材
│   ├── verify_pptx.py           # 逐页对象清单
│   ├── render_preview.py        # PPTX → PNG 近似渲染
│   ├── compare_source.py        # 原图 vs 转换结果并排对比图
│   ├── build_verification.py    # 生成可编辑性验证材料
│   ├── optimize_images.py       # 压缩验证材料图片，减小仓库体积
│   └── e2e_http_check.py        # HTTP 层端到端验证（上传→转换→下载→报告）
├── docs/
│   └── 图层分解方案.md          # 方案对比、像素信号、完整流程、踩坑记录、已知限制
├── tests/materials/             # 3 套测试素材 PDF
├── outputs/                     # 3 套素材的转换结果 + 预览图 + 对比图
├── verification/                # 可编辑性验证报告
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
└── pyproject.toml
```

---

## 九、可编辑性验证材料

| 材料 | 位置 | 内容 |
| --- | --- | --- |
| 验证报告 | [`verification/可编辑性验证报告.md`](verification/可编辑性验证报告.md) | 总体指标、OOXML 原生元素统计、逐页对象数、文本抽样 |
| 汇总数据 | `verification/summary.json` | 机器可读的统计结果 |
| 逐页对象清单 | `outputs/<素材>/objects.txt` | 每页的对象类型、位置、文本内容 |
| 元素明细报告 | `outputs/<素材>/report.json` | 每页每个元素的类型、包围盒、字号、颜色 |
| 页面预览图 | `outputs/<素材>/preview/slide_*.png` | 转换后每一页的渲染图 |
| 原图对比图 | `outputs/<素材>/compare/compare_*.png` | 原始 PDF 页面与转换结果并排对比 |
| 原始页面图 | `outputs/<素材>/compare/source/source_*.png` | 原始 PDF 每一页的渲染图，作为对比图的上半部分 |

自行复验：

```bash
# 查看某个 PPTX 里到底有哪些对象
python scripts/verify_pptx.py outputs/测试1/测试1.pptx

# 重新生成全部验证材料（对象清单、预览图、对比图、报告）
python scripts/build_verification.py

# 压缩验证材料图片（量化调色板 + 无损压缩，仓库体积可减 60%）
python scripts/optimize_images.py

# HTTP 层端到端验证（自动拉起服务，走一遍上传→转换→下载→报告）
python scripts/e2e_http_check.py
```

**判定标准**：如果某页只由一张覆盖整页的图片构成，那才是「把 PNG 当背景」的不合格结果。验证报告中的「退化整页图」一列统计的就是这种情况，本项目输出为 **0 页**。

### 怎么验证「图片是独立对象」

1. **用 PowerPoint / WPS 打开 PPTX**，点击页面上的任意一张照片——它会被单独选中，可以直接拖动、缩放、删除，而不会带着背景一起动。
2. **看 OOXML**：`unzip -p outputs/测试1/测试1.pptx ppt/slides/slide7.xml | grep -c '<p:pic>'` —— `<p:pic>` 的数量就是该页独立图片对象的个数（测试1 共 117 个 `<p:pic>`）。
3. **看对象清单**：`outputs/测试1/objects.txt` 里逐个对象列出了类型、包围盒与文本内容，可以确认每张照片都有自己独立的坐标范围。
4. **看统计**：`verification/可编辑性验证报告.md` 的「每页对象数量分布」显示单页最少 6 个、最多 75 个对象；若仍是一张整页图，该页的对象数会是 1。

```bash
# 统计三套素材 PPTX 中的原生元素数量
python scripts/verify_pptx.py outputs/测试1/测试1.pptx
```

---

## 十、3 套素材的转换结果

素材位于 `tests/materials/`，每份 10 页。

| 素材 | 页数 | 文本框 | 图片对象 | 形状 | 自由曲线 | 连接符 | 可编辑对象（非图片） | 可编辑占比 | 耗时 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 测试1.pdf | 10 | 257 | 117 | 17 | 2 | 9 | 285 | 100% | 104.9s |
| 测试2.pdf | 10 | 161 | 91 | 12 | 1 | 19 | 193 | 100% | 126.2s |
| 测试3.pdf | 10 | 134 | 51 | 6 | 0 | 5 | 145 | 100% | 78.8s |

| 素材 | 已分解页数 | 背景层 | 退化整页图 | 每页对象数分布 |
| --- | --- | --- | --- | --- |
| 测试1 | 10 / 10 | 9 | **0** | 15 ~ 75 |
| 测试2 | 10 / 10 | 10 | **0** | 6 ~ 47 |
| 测试3 | 9 / 10 | 8 | **0** | 8 ~ 32 |

> **统计口径**：「可编辑对象（非图片）」指非图片的原生对象（文本框 / 形状 / 自由曲线 / 表格 / 连接符）数量，可在 PowerPoint 中双击直接编辑；**图片对象同样是独立对象**——每张照片、Logo、图标都是单独的 `<p:pic>`，可以独立选中、移动、缩放。「可编辑占比」= 可独立编辑的对象数 ÷ 页面对象总数，100% 意味着**没有任何一页退化成「整页背景图」**（若整页只放一张截图，占比会很低）。
>
> 「退化整页图」= 某一页只由一张覆盖 95% 以上页面面积的图片构成。三套素材均为 **0 页**。

转换结果：

| 素材 | PPTX | 预览 | 对比 |
| --- | --- | --- | --- |
| 测试1 | [下载](outputs/测试1/测试1.pptx) | [预览图](outputs/测试1/preview/slide_01.png) | [对比图](outputs/测试1/compare/compare_01.png) |
| 测试2 | [下载](outputs/测试2/测试2.pptx) | [预览图](outputs/测试2/preview/slide_01.png) | [对比图](outputs/测试2/compare/compare_01.png) |
| 测试3 | [下载](outputs/测试3/测试3.pptx) | [预览图](outputs/测试3/preview/slide_01.png) | [对比图](outputs/测试3/compare/compare_01.png) |

> 完整目录见 [`outputs/`](outputs/) 与 [`verification/`](verification/)。

### 效果抽查

| 素材 | 页 | 分解效果 |
| --- | --- | --- |
| 测试1 | 第 1 页 | 英雄区照片从页面中独立抠出，斜切背景块还原为**自由曲线**，标题文字独立成文本框 |
| 测试1 | 第 7 页 | 「核心团队」4 张人像各自独立成图片对象，图标独立，蓝色卡片保留在背景层，4 列项目符号正确分列 |
| 测试1 | 第 9 页 | Logo 墙的 19 个 Logo 全部拆成独立图片对象，Logo 内的中英文保留在图里、不重复生成文本框 |
| 测试3 | 第 1 页 | 满版熊猫城市夜景插画，天然不可再分，保留为一张整页图片对象 |

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

- 10 页图片型 PDF 的转换耗时约 **80~130 秒**（主要开销在 OCR 的检测识别与图层分解的像素运算，两者都是纯 CPU 的 NumPy 计算）。
- 图层分解本身不引入任何模型推理，全部是积分图、形态学、连通域标记一类的 NumPy 运算；每页约 2~4 秒。
- OCR 引擎为**进程内单例**并通过线程锁串行访问，避免多请求同时初始化模型。
- 转换在 `asyncio.to_thread` 中执行，不阻塞事件循环；上传文件与结果会在 `PDF2PPTX_RETENTION_HOURS`（默认 6 小时）后自动清理。
- 若不需要独立图片对象，可关闭 `PDF2PPTX_ENABLE_LAYERING`，耗时回落至 30~55 秒。

## 十三、许可

本仓库代码采用 MIT 许可。依赖的 PyMuPDF 采用 AGPL-3.0 许可，若用于闭源商业产品请自行评估合规性（可替换为 `pypdfium2` 等宽松许可的渲染后端）。
