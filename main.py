#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scanned-pdf-translate: 扫描版 PDF 翻译工具
用法示例：
    python main.py input.pdf -o output.docx -f docx -l zh
    python main.py input.pdf -f pdf --model claude-sonnet-4-6

API Key 通过项目根目录的 .env 文件配置（参见 .env.example），
也可通过环境变量 ANTHROPIC_API_KEY / OPENAI_API_KEY 设置。
"""

import argparse
import sys
import os
from pathlib import Path


def parse_args():
    parser = argparse.ArgumentParser(
        description="扫描版 PDF 翻译工具（基于 PaddleOCR + LLM）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
API Key 配置（优先级从高到低）:
  1. 项目根目录 .env 文件（推荐，参见 .env.example）
  2. 环境变量 ANTHROPIC_API_KEY / OPENAI_API_KEY
        """,
    )
    parser.add_argument("input", help="输入 PDF 文件路径")
    parser.add_argument("-o", "--output", help="输出文件路径（默认：与输入同名）")
    parser.add_argument(
        "-f", "--format", choices=["docx", "pdf"], default="docx",
        help="输出格式 (默认: docx)"
    )
    parser.add_argument(
        "-l", "--target-lang", default="zh",
        help="目标语言代码 (默认: zh 简体中文)"
    )
    parser.add_argument(
        "--target-lang-name", default="简体中文",
        help="目标语言名称，用于 Prompt (默认: 简体中文)"
    )
    parser.add_argument(
        "--source-lang", default="en",
        help="源语言代码，用于 OCR (默认: en)"
    )
    parser.add_argument("--api-url", help="LLM API URL（OpenAI 兼容接口，覆盖 .env 中的值）")
    parser.add_argument(
        "--provider", choices=["anthropic", "openai"], default=None,
        help="API 提供商 (默认: anthropic，覆盖 .env 中的值)"
    )
    parser.add_argument("--model", help="模型名称")
    parser.add_argument(
        "--dpi", type=int, default=200,
        help="PDF 转图像分辨率 (默认: 200)"
    )
    parser.add_argument(
        "--no-gpu", action="store_true",
        help="禁用 GPU（PaddleOCR）"
    )
    parser.add_argument(
        "--pages", type=str, default=None,
        help="只处理指定页面，如 '1-3,5'（默认：全部）"
    )
    return parser.parse_args()


def parse_page_range(page_str: str, total: int):
    """解析页面范围字符串，如 '1-3,5' → {1,2,3,5}"""
    pages = set()
    for part in page_str.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            pages.update(range(int(start), int(end) + 1))
        elif part.isdigit():
            pages.add(int(part))
    return pages


def main():
    args = parse_args()

    # ── 配置 ─────────────────────────────────────────────────────────────────
    from config import Config
    config = Config.from_env()

    if args.api_url:
        config.api_url = args.api_url
    if args.provider:
        config.api_provider = args.provider
    if args.model:
        config.model = args.model

    config.source_lang = args.source_lang
    config.target_lang = args.target_lang
    config.target_lang_name = args.target_lang_name
    config.dpi = args.dpi
    config.output_format = args.format

    if not config.api_key:
        print("错误：未找到 API Key。")
        print("请在项目根目录创建 .env 文件（参考 .env.example），或设置环境变量。")
        sys.exit(1)

    # ── 路径处理 ──────────────────────────────────────────────────────────────
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误：文件不存在: {args.input}")
        sys.exit(1)

    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix(f".translated.{args.format}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"  输入:  {input_path}")
    print(f"  输出:  {output_path}")
    print(f"  格式:  {args.format.upper()}")
    print(f"  翻译:  {args.source_lang} → {args.target_lang_name}")
    print(f"  模型:  {config.model} ({config.api_provider})")
    print(f"{'='*60}\n")

    # ── Step 1: PDF → 图像 ────────────────────────────────────────────────────
    print("【Step 1】PDF 转图像...")
    from pdf_processor import PDFProcessor
    processor = PDFProcessor(dpi=config.dpi)
    all_pages = processor.convert(str(input_path))

    # 过滤页面
    if args.pages:
        selected = parse_page_range(args.pages, len(all_pages))
        pages = [p for p in all_pages if p['page_num'] in selected]
        print(f"  仅处理页面: {sorted(selected)}")
    else:
        pages = all_pages

    if not pages:
        print("错误：没有可处理的页面")
        sys.exit(1)

    # ── Step 2: 版面分析 ──────────────────────────────────────────────────────
    print("\n【Step 2】版面分析（PaddleOCR PPStructure）...")
    from layout_analyzer import LayoutAnalyzer
    analyzer = LayoutAnalyzer(lang=config.source_lang,
                               use_gpu=not args.no_gpu)
    layouts = analyzer.analyze(pages)

    total_regions = sum(len(l.regions) for l in layouts)
    print(f"  共检测到 {total_regions} 个版面区域")

    # ── Step 3: 翻译 ──────────────────────────────────────────────────────────
    print("\n【Step 3】翻译内容...")
    from translator import Translator
    translator = Translator(config)
    translator.translate_layout(layouts)

    # ── Step 4: 生成输出 ──────────────────────────────────────────────────────
    print(f"\n【Step 4】生成 {args.format.upper()} 输出...")
    if args.format == "docx":
        from docx_builder import DocxBuilder
        builder = DocxBuilder()
        builder.build(layouts, str(output_path))
    else:
        from pdf_builder import PDFBuilder
        builder = PDFBuilder(cn_font_path=config.cn_font_path,
                             cn_font_name=config.cn_font_name)
        builder.build(layouts, str(output_path))

    print(f"\n完成！输出文件: {output_path}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
