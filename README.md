# 扫描版 PDF 翻译工具

基于 PaddleOCR 版面分析 + 大语言模型翻译，将扫描版 PDF 文件翻译为中文（或其他语言），同时尽量还原原文的双栏、表格、图片等排版布局。

## 功能特点

- **版面智能分析**：使用 PaddleOCR PPStructure 自动识别标题、段落、表格、图片等区域
- **双栏布局保留**：自动检测双栏排版，按阅读顺序（左栏→右栏）重组内容
- **图片原样保留**：图片区域直接复制，不进行 OCR 或翻译
- **表格结构翻译**：识别表格结构并翻译单元格内容，保留行列格式
- **LLM 批量翻译**：支持 Anthropic Claude 和 OpenAI 兼容接口，批量翻译减少 API 调用
- **双格式输出**：支持输出 DOCX（Word）或 PDF

## 环境要求

- Python 3.8+
- Windows / Linux / macOS

## 安装依赖

```bash
# 安装 PaddlePaddle（CPU 版，固定版本以避免 Windows 兼容问题）
pip install paddlepaddle==2.6.2

# 安装其余依赖
pip install paddleocr==2.8.1
pip install PyMuPDF anthropic openai reportlab python-docx beautifulsoup4
```

> **注意**：paddlepaddle 3.x 在 Windows 上存在 PIR 后端兼容问题，请务必使用 2.6.2 版本。

## 快速开始

### 1. 配置 API Key

在项目根目录创建 `.env` 文件（复制 `.env.example` 后填写）：

```bash
cp .env.example .env
```

编辑 `.env`：

```ini
# 使用 Anthropic Claude（推荐）
ANTHROPIC_API_KEY=sk-ant-...

# 或使用 OpenAI / DeepSeek 等兼容接口
# OPENAI_API_KEY=sk-...
# OPENAI_API_URL=https://api.deepseek.com/v1
# LLM_MODEL=deepseek-chat
```

> `.env` 已加入 `.gitignore`，不会被提交到 git，API Key 不会泄露。

### 2. 执行翻译

```bash
# 最简用法：翻译整个 PDF，输出 DOCX
python main.py 论文.pdf

# 指定输出路径和格式
python main.py 论文.pdf -o 论文_中文版.docx -f docx

# 输出为 PDF
python main.py 论文.pdf -f pdf -o 论文_中文版.pdf

# 只翻译第 1-3 页（用于测试）
python main.py 论文.pdf --pages 1-3

# 翻译特定页码（如第 1、3、5 页）
python main.py 论文.pdf --pages 1,3,5
```

## 完整参数说明

```
python main.py <输入PDF> [选项]

位置参数:
  input                   输入 PDF 文件路径

输出选项:
  -o, --output            输出文件路径（默认与输入同名，后缀改为 .translated.docx）
  -f, --format            输出格式：docx（默认）或 pdf

语言选项:
  -l, --target-lang       目标语言代码（默认: zh）
  --target-lang-name      目标语言名称，用于提示词（默认: 简体中文）
  --source-lang           源语言代码，用于 OCR 识别（默认: en）

API 选项:
  --api-key               LLM API Key（也可通过环境变量设置）
  --api-url               OpenAI 兼容接口 URL
  --provider              API 提供商：anthropic（默认）或 openai
  --model                 模型名称（默认: claude-sonnet-4-6）

其他选项:
  --dpi                   PDF 转图像分辨率（默认: 200，越高越清晰但越慢）
  --no-gpu                禁用 GPU（默认使用 CPU）
  --pages                 只处理指定页面，如 1-3 或 1,3,5
```

## 使用示例

### 翻译英文学术论文

```bash
python main.py paper.pdf \
  --pages 1-5 \
  --format docx \
  -o paper_zh.docx
```

### 使用 OpenAI 兼容接口（如 DeepSeek、Qwen 等）

```bash
python main.py paper.pdf \
  --provider openai \
  --api-key "your-api-key" \
  --api-url "https://api.deepseek.com/v1" \
  --model "deepseek-chat"
```

### 翻译为其他语言

```bash
# 翻译为日文
python main.py paper.pdf \
  --target-lang ja \
  --target-lang-name "日本語"

# 翻译为法文
python main.py paper.pdf \
  --target-lang fr \
  --target-lang-name "Français"
```

### 高质量扫描件（提高 DPI）

```bash
python main.py paper.pdf --dpi 300
```

## 处理流程

```
输入 PDF
  │
  ▼ Step 1: PDF → 图像（PyMuPDF）
  │  每页渲染为高分辨率图像
  │
  ▼ Step 2: 版面分析（PaddleOCR PPStructure）
  │  识别标题、段落、表格、图片区域及 bbox 坐标
  │  自动检测单栏 / 双栏布局
  │
  ▼ Step 3: 内容翻译（LLM API）
  │  文本/标题/图注：批量翻译（减少 API 调用）
  │  表格：解析 HTML 结构，逐单元格翻译后重建
  │  图片：直接跳过，原样保留
  │
  ▼ Step 4: 输出构建
     DOCX：无边框双列 Table 作布局容器，还原双栏效果
     PDF ：reportlab 按原始坐标缩放定位文本块
```

## 支持的版面区域类型

| 类型 | 处理方式 |
|------|---------|
| 标题（title） | 翻译，居中加粗显示 |
| 正文段落（text） | 翻译，两端对齐 |
| 表格（table） | 翻译单元格内容，保留表格结构 |
| 图片（figure） | 直接复制原图，不做任何处理 |
| 图注/表注（caption） | 翻译 |

## 常见问题

**Q: 首次运行很慢？**
A: PaddleOCR 首次运行会自动下载版面分析和 OCR 模型（约 200-500MB），下载完成后后续运行正常。

**Q: 翻译结果乱码或字体缺失（PDF 输出）？**
A: PDF 输出依赖系统中文字体。程序默认使用 `C:\Windows\Fonts\simhei.ttf`（黑体）。如果找不到，可在 `config.py` 中修改 `cn_font_path` 指定其他字体路径。

**Q: Windows 上 PaddlePaddle 报错 `ConvertPirAttribute2RuntimeAttribute`？**
A: 这是 paddlepaddle 3.x 在 Windows 上的已知兼容问题。请确认使用 `paddlepaddle==2.6.2`：
```bash
pip install paddlepaddle==2.6.2
```

**Q: 如何配置 API Key？**
A: 在项目根目录创建 `.env` 文件（参考 `.env.example`），填入 API Key 即可。程序启动时自动加载，无需每次手动设置。`.env` 已被 `.gitignore` 排除，不会泄露。

## 项目结构

```
doc-translate/
├── main.py              # CLI 入口
├── config.py            # 全局配置（API、模型、字体等）
├── pdf_processor.py     # PDF → 图像（PyMuPDF）
├── layout_analyzer.py   # 版面分析（PaddleOCR PPStructure）
├── translator.py        # LLM 翻译（Anthropic / OpenAI 兼容）
├── docx_builder.py      # DOCX 输出构建
├── pdf_builder.py       # PDF 输出构建（reportlab）
└── requirements.txt     # 依赖列表
```

## 许可证

MIT License
