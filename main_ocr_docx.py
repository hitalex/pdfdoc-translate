#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ocr-docx 两步翻译流程：
  步骤 1  PDF → 图像 → PaddleOCR 版面分析 → 重建 DOCX（保留原文）
  步骤 2  读取原文 DOCX → LLM 批量翻译段落 → 输出翻译版 DOCX

用法：
  # 完整两步（默认）
  python main_ocr_docx.py input.pdf

  # 只做步骤 1（生成原文 docx，不翻译，便于检查 OCR 质量）
  python main_ocr_docx.py input.pdf --step 1

  # 只做步骤 2（已有 _ocr.docx，直接翻译）
  python main_ocr_docx.py input.pdf --step 2

  # 自定义路径
  python main_ocr_docx.py input.pdf -o output_translated.docx --ocr-docx intermediate.docx
"""

import argparse
import sys
from pathlib import Path


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="两步式 OCR → DOCX → 翻译 流程",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input", help="输入 PDF 文件路径")
    parser.add_argument(
        "-o", "--output",
        help="翻译后 DOCX 输出路径（默认：<input>_translated.docx）",
    )
    parser.add_argument(
        "--ocr-docx",
        help="OCR 重建 DOCX 中间文件路径（默认：<input>_ocr.docx）",
    )
    parser.add_argument(
        "--step", choices=["1", "2", "all"], default="all",
        help="执行步骤: 1=仅 OCR→DOCX, 2=仅翻译 DOCX, all=两步都做（默认: all）",
    )
    parser.add_argument(
        "-l", "--target-lang", default="zh",
        help="目标语言代码（默认: zh）",
    )
    parser.add_argument(
        "--target-lang-name", default="简体中文",
        help="目标语言名称，用于 Prompt（默认: 简体中文）",
    )
    parser.add_argument(
        "--source-lang", default="en",
        help="源语言代码，用于 OCR（默认: en）",
    )
    parser.add_argument("--api-url", help="LLM API URL（OpenAI 兼容接口）")
    parser.add_argument(
        "--provider", choices=["anthropic", "openai"], default=None,
        help="API 提供商（默认: anthropic）",
    )
    parser.add_argument("--model", help="模型名称")
    parser.add_argument(
        "--dpi", type=int, default=200,
        help="PDF 转图像分辨率（默认: 200）",
    )
    parser.add_argument(
        "--no-gpu", action="store_true",
        help="禁用 GPU（PaddleOCR）",
    )
    parser.add_argument(
        "--pages", type=str, default=None,
        help="只处理指定页面，如 '1-3,5'（默认：全部）",
    )
    parser.add_argument(
        "--batch-size", type=int, default=20,
        help="翻译时每批段落数（默认: 20）",
    )
    return parser.parse_args()


def parse_page_range(page_str: str, total: int):
    pages = set()
    for part in page_str.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            pages.update(range(int(start), int(end) + 1))
        elif part.isdigit():
            pages.add(int(part))
    return pages


# ── 步骤 1：PDF → OCR → 原文 DOCX ─────────────────────────────────────────────

def step1_ocr_to_docx(args, config, ocr_docx_path: Path):
    """运行 OCR 版面分析，用原始文本（不翻译）重建 DOCX。"""
    print("\n【Step 1-A】PDF 转图像...")
    from pdf_processor import PDFProcessor
    processor = PDFProcessor(dpi=config.dpi)
    all_pages = processor.convert(str(args.input))

    if args.pages:
        selected = parse_page_range(args.pages, len(all_pages))
        pages = [p for p in all_pages if p["page_num"] in selected]
        print(f"  仅处理页面: {sorted(selected)}")
    else:
        pages = all_pages

    if not pages:
        print("错误：没有可处理的页面")
        sys.exit(1)

    print(f"\n【Step 1-B】版面分析（PaddleOCR PPStructure）...")
    from layout_analyzer import LayoutAnalyzer
    analyzer = LayoutAnalyzer(lang=config.source_lang, use_gpu=not args.no_gpu)
    layouts = analyzer.analyze(pages)

    total_regions = sum(len(l.regions) for l in layouts)
    print(f"  共检测到 {total_regions} 个版面区域")

    print(f"\n【Step 1-C】重建原文 DOCX（跳过翻译）...")
    from docx_builder import DocxBuilder
    builder = DocxBuilder()
    builder.build(layouts, str(ocr_docx_path))
    print(f"  原文 DOCX 已保存: {ocr_docx_path}")

    return layouts  # 供调试或 step 2 直接复用（可选）


# ── 步骤 2：翻译 DOCX ──────────────────────────────────────────────────────────

def step2_translate_docx(args, config, ocr_docx_path: Path, translated_path: Path):
    """读取 OCR 重建 DOCX，批量翻译所有段落，保存翻译版 DOCX。"""
    if not ocr_docx_path.exists():
        print(f"错误：找不到原文 DOCX：{ocr_docx_path}")
        print("请先运行步骤 1，或使用 --ocr-docx 指定正确路径。")
        sys.exit(1)

    if not config.api_key:
        print("错误：未找到 API Key。")
        print("请在项目根目录创建 .env 文件（参考 .env.example），或设置环境变量。")
        sys.exit(1)

    print(f"\n【Step 2】翻译 DOCX（{ocr_docx_path.name} → {translated_path.name}）...")
    from translator import Translator
    from docx_paragraph_translator import translate_docx_file

    translator = Translator(config)
    translate_docx_file(
        input_path=str(ocr_docx_path),
        output_path=str(translated_path),
        translator=translator,
        batch_size=args.batch_size,
    )


# ── 主流程 ─────────────────────────────────────────────────────────────────────

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

    # ── 路径 ─────────────────────────────────────────────────────────────────
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误：文件不存在: {args.input}")
        sys.exit(1)

    ocr_docx_path = (
        Path(args.ocr_docx) if args.ocr_docx
        else input_path.with_suffix("").with_name(input_path.stem + "_ocr.docx")
    )
    translated_path = (
        Path(args.output) if args.output
        else input_path.with_suffix("").with_name(input_path.stem + "_translated.docx")
    )

    for p in (ocr_docx_path, translated_path):
        p.parent.mkdir(parents=True, exist_ok=True)

    # ── 打印摘要 ──────────────────────────────────────────────────────────────
    step_label = {"1": "仅步骤1 (OCR→DOCX)", "2": "仅步骤2 (翻译DOCX)", "all": "全流程"}[args.step]
    print(f"\n{'='*60}")
    print(f"  输入 PDF:    {input_path}")
    print(f"  OCR DOCX:   {ocr_docx_path}")
    print(f"  翻译 DOCX:  {translated_path}")
    print(f"  执行步骤:   {step_label}")
    print(f"  翻译方向:   {args.source_lang} → {args.target_lang_name}")
    if args.step != "1":
        print(f"  模型:       {config.model} ({config.api_provider})")
    print(f"{'='*60}\n")

    # ── 执行 ─────────────────────────────────────────────────────────────────
    if args.step in ("1", "all"):
        step1_ocr_to_docx(args, config, ocr_docx_path)

    if args.step in ("2", "all"):
        step2_translate_docx(args, config, ocr_docx_path, translated_path)

    print(f"\n{'='*60}")
    if args.step in ("1", "all"):
        print(f"  原文 DOCX:  {ocr_docx_path}")
    if args.step in ("2", "all"):
        print(f"  翻译 DOCX:  {translated_path}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
