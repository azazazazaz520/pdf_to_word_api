# pdf转word服务

一个基于 Python、FastAPI、PaddleOCR 和 `python-docx` 的 PDF 转 Word 服务。服务接收 PDF 文件，创建异步转换任务，按页面特征选择文本解析或 OCR 路线，最终生成可下载的 `.docx` 文件。

## 功能

- 提供健康检查、任务创建、任务查询、结果下载和任务取消接口。
- 校验 PDF 文件头、文件大小、页数、解析状态和加密状态。
- 使用独立 worker 进程执行转换，API 进程负责认证、上传、任务调度和结果查询。
- 使用本地 SQLite 保存任务状态，任务文件按 TTL 清理；服务重启后可以恢复仍存在输入文件的未完成任务。
- 支持 Bearer Token、`X-API-Key`，以及配置 Supabase 后使用 Supabase 用户 JWT 调用。
- 自动分析文本层，并按页面选择文本解析或 OCR；也可以通过任务参数强制指定路线。
- 支持标题、段落、项目符号、有序列表、常见公式、原生 Word 表格、嵌入图片和 PDF 书签。
- 提供两种 DOCX 导出模式：默认的结构化导出，以及按源坐标保留版式的高保真导出。
- 记录转换阶段、OCR 置信度、页面复核提示、渲染校验和质量门禁结果。

## 转换路线

服务采用“先分析，再转换”的处理方式：

```text
上传 PDF
   │
   ├─ 文件和 PDF 校验
   ├─ 创建任务并写入 SQLite
   └─ worker 处理
        ├─ auto：逐页分析文本层，有文本页面走 text，无文本页面走 ocr
        ├─ text：满足文本层阈值后，全部页面走文本布局解析
        └─ ocr：全部页面渲染后交给 PaddleOCR
              │
              ├─ structured：段落、标题、列表、表格和图片的可编辑导出
              └─ fidelity：按源坐标定位文本、表格和图片，尽量保留页面版式
```

### 路由参数

| 参数 | 可选值 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `route_mode` | `auto`、`text`、`ocr` | `auto` | 控制文本层分析和 OCR 使用方式 |
| `export_mode` | `structured`、`flow`、`fidelity`、`fidelity_hybrid` | `structured` | `flow` 是 `structured` 的兼容别名 |
| `PDF_SERVICE_ENGINE` | `structure-lite`、`structure-table-lite`、`structure`、`vl` | `structure-lite` | OCR/结构识别引擎；仅在需要 OCR 的页面加载模型 |

选择建议：

- 普通文字型 PDF：使用 `route_mode=auto` 和 `export_mode=structured`。
- 需要识别扫描文字：使用 `auto` 或 `ocr`；OCR 结果需要根据 `needs_review_pages` 和质量报告复核。
- 需要原生 Word 表格：选择支持表格识别的 `structure-table-lite`，或使用文本层布局路线。
- 更重视源文件坐标和视觉版式：选择 `fidelity` 或 `fidelity_hybrid`。高保真导出仍会受到 PDF 字体、复杂对象和 Word 排版规则的影响。
- `vl` 和完整 `structure` 引擎在当前 CPU 环境下耗时和内存开销较高，适合作为按需增强路线。

## 环境要求

- Windows 10/11，或 Linux x86_64
- Python 3.11
- CPU 版 PaddlePaddle 3.3.1
- PaddleOCR 3.7.0
- PaddleX 3.7.2，并安装 `ocr` 额外依赖
- `python-docx` 1.2.0
- `pypdfium2` 5.13.0
- FastAPI 0.116.1 和 Uvicorn 0.35.0

依赖版本以 [requirements.txt](requirements.txt) 为准。核心 PDF 解析、OCR、DOCX 导出和 HTTP 服务代码使用跨平台 Python 库，Windows 和 Linux 均可运行。当前项目主要在 Windows CPU 环境完成验证，Linux 需要按照本机发行版和 PaddlePaddle 提供的 wheel 进行安装后再做完整验收。

首次运行 OCR 路线时会下载模型，模型缓存默认放在 `model_cache/`；服务任务和 SQLite 数据默认放在 `service_data/`。这两个目录以及 `artifacts/` 都属于本地运行目录，不应提交到版本库。

Linux 运行时有两个已知差异：

- `PDF_SERVICE_RENDER_VALIDATION` 默认值为 `true`，其 DOCX 渲染步骤依赖 Microsoft Word COM，只能在安装了 Microsoft Word 的 Windows 环境执行。Linux 启动服务时建议设置为 `false`，或自行接入 LibreOffice 等兼容的 DOCX 渲染器。
- 当前字体扫描逻辑优先读取 Windows 注册表和 `C:\Windows\Fonts`。Linux 仍可完成转换，但高保真导出中的系统字体匹配可能退回字体替代；需要高保真验收时，应准备与 PDF 相同或等价的字体，并单独检查渲染结果。

## 安装

在项目根目录的 PowerShell 中执行：

```powershell
uv venv --python 3.11 .venv
& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install --index-url https://www.paddlepaddle.org.cn/packages/stable/cpu/ paddlepaddle==3.3.1
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Linux 下使用 Bash 执行等价安装：

```bash
python3.11 -m venv .venv
./.venv/bin/python -m pip install --upgrade pip
./.venv/bin/python -m pip install paddlepaddle==3.3.1
./.venv/bin/python -m pip install -r requirements.txt
```

如果当前 Linux 平台没有可用的 `paddlepaddle==3.3.1` wheel，应按照 PaddlePaddle 官方对应版本和平台说明选择安装源，再执行其余依赖安装。

可以显式指定模型缓存目录：

```powershell
$env:PADDLE_PDX_CACHE_HOME = (Join-Path (Get-Location) "model_cache")
```

模型完成缓存后，如需离线复测，可以在当前 PowerShell 会话中跳过模型源检查：

```powershell
$env:PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK = "True"
```

## 启动服务

服务默认监听 `127.0.0.1:8765`，并要求配置 `PDF_SERVICE_TOKEN`。

Windows PowerShell：

```powershell
$env:PDF_SERVICE_TOKEN = "请替换为随机生成的长令牌"
$env:PDF_SERVICE_HOST = "127.0.0.1"
$env:PDF_SERVICE_PORT = "8765"

& .\.venv\Scripts\python.exe -m src.service.app
```

Linux Bash：

```bash
export PDF_SERVICE_TOKEN="请替换为随机生成的长令牌"
export PDF_SERVICE_HOST="127.0.0.1"
export PDF_SERVICE_PORT="8765"
export PDF_SERVICE_RENDER_VALIDATION="false"

./.venv/bin/python -m src.service.app
```

启动后可以访问 FastAPI 文档页：

```text
http://127.0.0.1:8765/docs
```

服务令牌只应保存在服务端环境变量或密钥管理系统中。部署到服务器时，应在反向代理层配置 HTTPS、访问控制、请求体限制和限流；确需局域网访问时，再将 `PDF_SERVICE_HOST` 设置为 `0.0.0.0`。

## API 使用

所有业务 API 接口都需要通过以下任一种方式提供凭证：

```http
Authorization: Bearer <token>
```

或：

```http
X-API-Key: <token>
```

当 `SUPABASE_URL` 和 `SUPABASE_ANON_KEY` 都已配置时，服务令牌之外还接受 Supabase 匿名会话产生的短期 JWT，并按用户限制任务查询范围。`PDF_SERVICE_TOKEN` 仍然必须配置。

### 接口列表

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `GET` | `/health` | 返回服务、模型、队列、路由和质量配置；需要认证 |
| `POST` | `/api/pdf-to-word/jobs` | 上传 PDF 并创建转换任务 |
| `GET` | `/api/pdf-to-word/jobs/{job_id}` | 查询任务状态和质量摘要 |
| `GET` | `/api/pdf-to-word/jobs/{job_id}/result` | 下载生成的 DOCX |
| `DELETE` | `/api/pdf-to-word/jobs/{job_id}` | 取消排队或处理中的任务 |

任务状态包括：`queued`、`processing`、`succeeded`、`failed`、`cancelled` 和 `timed_out`。

### 创建任务

上传字段：

- `file`：必填，PDF 文件。
- `route_mode`：可选，`auto`、`text` 或 `ocr`。
- `export_mode`：可选，`structured`、`flow`、`fidelity` 或 `fidelity_hybrid`。

PowerShell 示例：

```powershell
$token = $env:PDF_SERVICE_TOKEN
$pdfPath = (Resolve-Path .\fixtures\synthetic_text_table.pdf).Path

$job = curl.exe -sS `
  -X POST "http://127.0.0.1:8765/api/pdf-to-word/jobs" `
  -H "Authorization: Bearer $token" `
  -F "file=@$pdfPath" `
  -F "route_mode=auto" `
  -F "export_mode=structured" | ConvertFrom-Json

$job | ConvertTo-Json -Depth 8
$jobId = $job.job_id
```

创建成功后，响应中会返回 `job_id`、`status`、`progress`、`route`、`export_mode`、`page_count`、`download_url` 和质量摘要字段。

### 查询任务并下载结果

```powershell
$status = curl.exe -sS `
  -H "Authorization: Bearer $token" `
  "http://127.0.0.1:8765/api/pdf-to-word/jobs/$jobId" | ConvertFrom-Json

$status | Select-Object job_id, status, progress, route, route_reason, table_count, needs_review_pages

curl.exe -sS -L `
  -H "Authorization: Bearer $token" `
  "http://127.0.0.1:8765/api/pdf-to-word/jobs/$jobId/result" `
  -o .\output\converted.docx
```

只有任务状态为 `succeeded` 时才能下载结果。客户端应在 `queued` 或 `processing` 状态下按间隔轮询，收到 `failed`、`cancelled` 或 `timed_out` 后读取 `error` 和 `warnings` 字段。

### 常见响应状态码

| 状态码 | 场景 |
| --- | --- |
| `400` | PDF 解析失败、加密 PDF 或路由/导出参数无效 |
| `401` | 缺少或无法验证访问凭证 |
| `409` | 任务尚未成功，暂时没有可下载结果 |
| `413` | 上传文件过大、页数超过限制 |
| `415` | 文件头不是有效 PDF |
| `429` | 任务队列已满 |
| `503` | 服务未配置令牌，或 Supabase 身份服务暂时不可用 |

## 配置

### 服务与任务

| 环境变量 | 默认值 | 说明 |
| --- | --- | --- |
| `PDF_SERVICE_HOST` | `127.0.0.1` | 监听地址 |
| `PDF_SERVICE_PORT` | `8765` | 监听端口 |
| `PDF_SERVICE_TOKEN` | 无 | 必填，服务端长期令牌 |
| `PDF_SERVICE_ENGINE` | `structure-lite` | OCR/结构识别引擎 |
| `PDF_SERVICE_WORKER_PROCESSES` | `1` | worker 数量，范围为 1–4；每个 worker 可能加载一份模型 |
| `PDF_SERVICE_MAX_PENDING_JOBS` | `4` | 允许排队或处理中的最大任务数，不能小于 worker 数量 |
| `PDF_SERVICE_DATA_ROOT` | `service_data` | 任务文件、SQLite 和阶段日志根目录 |
| `PDF_SERVICE_JOB_TTL_SECONDS` | `3600` | 任务及产物保留时间 |
| `PDF_SERVICE_CLEANUP_INTERVAL_SECONDS` | `60` | 清理周期 |
| `PDF_SERVICE_MAX_RETRIES` | `1` | worker 失败或超时后的重试次数，范围为 0–3 |
| `PDF_SERVICE_TASK_TIMEOUT_SECONDS` | `300` | 单任务软超时预算 |
| `PDF_SERVICE_OCR_TIME_BUDGET_SECONDS` | `60` | 含 OCR 页面任务的 OCR 时间预算 |
| `PDF_SERVICE_MAX_UPLOAD_BYTES` | `52428800` | 最大上传大小，默认 50 MiB |
| `PDF_SERVICE_MAX_PAGES` | `100` | 单个 PDF 的最大页数 |
| `PDF_SERVICE_ALLOWED_ORIGINS` | 空 | 逗号分隔的 CORS 来源列表 |

### 路由与图像

| 环境变量 | 默认值 | 说明 |
| --- | --- | --- |
| `PDF_SERVICE_ROUTE_MODE` | `auto` | `auto`、`text` 或 `ocr` |
| `PDF_SERVICE_TEXT_MIN_PAGE_CHARS` | `20` | 页面被视为可用文本页的最小字符数 |
| `PDF_SERVICE_TEXT_MIN_PAGE_RATIO` | `0.6` | `text` 强制路线要求满足文本阈值的页面比例 |
| `PDF_SERVICE_TEXT_HIGH_QUALITY_RATIO` | `0.8` | 文本层高质量页面比例阈值，随健康信息返回 |
| `PDF_SERVICE_TEXT_FULL_PAGE_IMAGE_MIN_PIXELS` | `300000` | 判断全页图像信号的最小像素数 |
| `PDF_SERVICE_TEXT_GARBLED_CHAR_RATIO` | `0.05` | 异常字形比例阈值 |
| `PDF_SERVICE_PAGE_IMAGE_MAX_PIXELS` | `4194304` | OCR 页面渲染像素上限，默认 4 MP |
| `PDF_SERVICE_PAGE_IMAGE_JPEG_QUALITY` | `88` | OCR 页面 JPEG 质量 |
| `PDF_SERVICE_EMBEDDED_IMAGE_MAX_PIXELS` | `6000000` | DOCX 内嵌图片像素上限 |
| `PDF_SERVICE_EMBEDDED_IMAGE_JPEG_QUALITY` | `85` | DOCX 内嵌 JPEG 质量 |
| `PDF_SERVICE_EMBEDDED_IMAGE_PNG_OPTIMIZE` | `false` | 是否优化内嵌 PNG |

### 导出、字体与质量校验

| 环境变量 | 默认值 | 说明 |
| --- | --- | --- |
| `PDF_SERVICE_EXPORT_MODE` | `structured` | 默认 `structured`；支持 `flow`、`fidelity`、`fidelity_hybrid` |
| `PDF_SERVICE_EMBED_PDF_FONTS` | `true` | 尝试嵌入 PDF 中使用的字体 |
| `PDF_SERVICE_CALIBRATE_FONT_METRICS` | `true` | 为高保真文本定位标定字体指标 |
| `PDF_SERVICE_INCLUDE_TOC` | `false` | 是否生成目录 |
| `PDF_SERVICE_INCLUDE_BOOKMARKS` | `true` | 是否读取 PDF 书签并应用到标题 |
| `PDF_SERVICE_BLOCK_READING_ORDER_ENABLED` | `true` | 是否启用内容块阅读顺序分析 |
| `PDF_SERVICE_BLOCK_READING_ORDER_FALLBACK_ENABLED` | `true` | 阅读顺序无法完整建立时是否使用坐标顺序 |
| `PDF_SERVICE_RENDER_VALIDATION` | `true` | 是否渲染 DOCX 并进行输出检查 |
| `PDF_SERVICE_RENDER_SSIM` | `true` | 是否计算页面结构相似度 |
| `PDF_SERVICE_RENDER_SSIM_DPI` | `110` | 相似度检查的渲染 DPI |
| `PDF_SERVICE_RENDER_TEXT_COMPARE` | `true` | 是否进行文本比较 |
| `PDF_SERVICE_SSIM_THRESHOLD` | `0.98` | 相似度质量阈值 |
| `PDF_SERVICE_RENDER_TIMEOUT_SECONDS` | `180` | 渲染校验超时时间 |
| `PDF_SERVICE_QUALITY_GATE_ENABLED` | `true` | 是否启用质量门禁 |
| `PDF_SERVICE_PAGE_DELTA_WARN_RATIO` | `0.05` | 页数变化比例告警阈值 |
| `PDF_SERVICE_PAGE_DELTA_WARN_ABSOLUTE` | `3` | 页数变化绝对值告警阈值 |

### 高保真兼容路径

| 环境变量 | 默认值 | 说明 |
| --- | --- | --- |
| `PDF_SERVICE_FIDELITY_AUTO_FALLBACK` | `false` | 是否允许高保真区域自动降级 |
| `PDF_SERVICE_FIDELITY_FALLBACK_DPI` | `200` | 降级页面渲染 DPI |
| `PDF_SERVICE_FIDELITY_FALLBACK_MAX_PIXELS` | `2000000` | 降级图像最大像素数 |
| `PDF_SERVICE_FIDELITY_PAGE_IMAGE_MAX_PIXELS` | `8000000` | 高保真页面图像最大像素数 |
| `PDF_SERVICE_FIDELITY_REVALIDATE_AFTER_FALLBACK` | `true` | 降级后是否重新进行质量校验 |

### Supabase 身份校验

| 环境变量 | 默认值 | 说明 |
| --- | --- | --- |
| `SUPABASE_URL` | 空 | Supabase 项目地址 |
| `SUPABASE_ANON_KEY` | 空 | Supabase 匿名公钥 |
| `PDF_SERVICE_SUPABASE_AUTH_TIMEOUT_SECONDS` | `5` | 身份校验请求超时时间 |
| `PDF_SERVICE_SUPABASE_AUTH_CACHE_SECONDS` | `60` | JWT 身份缓存时间；设置为 `0` 可关闭缓存 |

## 本地验证

生成合成 PDF 样本：

```powershell
& .\.venv\Scripts\python.exe scripts\generate_fixture.py
& .\.venv\Scripts\python.exe scripts\generate_table_fixture.py
```

运行独立转换验证：

```powershell
# 轻量文本和版面解析
& .\.venv\Scripts\python.exe -m src.run_validation `
  --engine structure-lite

# 启用表格识别
& .\.venv\Scripts\python.exe -m src.run_validation `
  --engine structure-table-lite `
  --input .\fixtures\synthetic_table_one_page.pdf

# PaddleOCR-VL 多页重组；CPU 环境耗时较长
& .\.venv\Scripts\python.exe -m src.run_validation `
  --engine vl `
  --input .\fixtures\synthetic_text_table.pdf
```

指定输入文件和输出目录：

```powershell
& .\.venv\Scripts\python.exe -m src.run_validation `
  --engine structure-lite `
  --input .\fixtures\synthetic_text_table.pdf `
  --output .\artifacts\outputs
```

每次验证会在 `artifacts\outputs\<engine>_<timestamp>\` 下生成：

- `json\`：识别结果 JSON；
- `markdown\`：PaddleOCR Markdown 结果；
- `word\`：DOCX 结果；
- `report.json`：运行耗时、输入信息、结果摘要和 DOCX 完整性检查。

运行完整自动化测试：

```powershell
& .\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

测试覆盖文本层路由、版面和阅读顺序、表格、公式、字体、结构化导出、高保真导出、混合页面、质量校验、worker 调度、认证和服务接口。

## 任务运行数据

默认每个任务目录位于 `service_data\jobs\<job_id>\`，包含：

- 输入 PDF；
- 输出 DOCX；
- `progress.json`：最新进度和状态；
- `stages.jsonl`：按时间追加的阶段事件；
- 取消标记和其他任务元数据。

任务状态同时保存在 `service_data\jobs.sqlite3`。阶段日志可用于定位文件校验、文本分析、模型加载、逐页 OCR、布局解析、DOCX 导出、渲染校验和质量门禁中的具体问题。服务默认只保留 TTL 内的任务产物，生产环境应根据数据合规要求调整存储位置和清理策略。

## 项目结构

```text
pdf_to_word_api/
├─ src/
│  ├─ service/       FastAPI 服务、任务管理和 SQLite 状态存储
│  ├─ worker/        worker 进度与质量处理
│  ├─ layout/        PDF 文本、图片、矢量和表格版面解析
│  ├─ ir/            文档中间表示和书签处理
│  ├─ export/        structured / fidelity DOCX 导出
│  ├─ fonts/         字体解析、匹配、嵌入和指标标定
│  ├─ validate/      DOCX 渲染、文本和相似度校验
│  ├─ pdf_worker.py  单任务转换入口
│  ├─ pdf_routing.py 文本层和页面特征分析
│  └─ run_validation.py 独立验证入口
├─ tests/            unittest 自动化测试
├─ scripts/          合成 PDF 样本生成脚本
├─ fixtures/         可复现的 PDF 输入样本
├─ docs/             评估报告、实施记录和验证结论
├─ artifacts/        本地验证产物
├─ model_cache/      PaddleOCR 模型缓存
├─ service_data/     服务任务和状态数据
├─ requirements.txt  Python 依赖
└─ README.md         使用说明
```

## 当前边界

- PDF 转 Word 的结果受原始 PDF 文本层、扫描质量、字体、阅读顺序和版面复杂度影响，不能保证所有文件达到逐像素一致。
- `structured` 模式优先保留可编辑结构，正文可能按照 Word 页面宽度重新排版。
- OCR 路线会产生模型加载和逐页推理开销；当前 CPU 实测时间随模型、页数和文件内容变化，不能据少量样本承诺固定耗时。
- `fidelity` 路线适合高保真兼容场景，复杂字体、透明对象、无边框表格和特殊 PDF 绘图指令仍需单独验收。
- 当前任务存储为单机本地 SQLite 和文件目录，适合独立服务和联调。多实例部署前还需要共享队列、跨主机 worker 租约、进程树回收、资源隔离和反向代理安全配置。
- 真实业务样本不应直接放入版本库；建议在受控目录中进行批量质量、耗时、页数、DOCX 体积和失败率评估。

## 相关文档

- [项目目录说明](docs/project-structure.md)
- [验证摘要](docs/validation-summary.md)
- [当前讨论结论](docs/current-discussion-results.md)
- [WeKnora 转换评估](docs/weknora-conversion-assessment.md)
