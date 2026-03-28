# -*- coding: utf-8 -*-
"""
全局配置：API、模型、语言、字体、输出选项
"""

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

# 自动加载项目根目录的 .env 文件（不覆盖已有的环境变量）
load_dotenv(Path(__file__).parent / ".env", override=False)


@dataclass
class Config:
    # ── 翻译 API ──────────────────────────────────────────────────────────────
    # 支持两种模式：
    #   "anthropic" : 使用 Anthropic 原生 SDK (api_url 可不填)
    #   "openai"    : 使用 OpenAI 兼容接口 (需填 api_url)
    api_provider: str = "anthropic"
    api_key: str = field(default_factory=lambda: os.environ.get("ANTHROPIC_API_KEY", ""))
    api_url: Optional[str] = None          # OpenAI兼容时填: e.g. "https://api.openai.com/v1"
    model: str = "claude-sonnet-4-6"       # 或 "gpt-4o" 等

    # ── 语言 ──────────────────────────────────────────────────────────────────
    source_lang: str = "en"                # 源语言 (用于 OCR 引擎选择)
    target_lang: str = "zh"               # 目标语言
    target_lang_name: str = "简体中文"    # 用于 prompt

    # ── OCR & 版面分析 ────────────────────────────────────────────────────────
    ocr_lang: str = "en"                   # PaddleOCR 语言: en / ch / ...
    dpi: int = 200                         # PDF → 图像分辨率

    # ── 输出 ──────────────────────────────────────────────────────────────────
    output_format: str = "docx"            # "docx" | "pdf"
    page_width_cm: float = 21.0           # A4 宽
    page_height_cm: float = 29.7          # A4 高
    margin_cm: float = 2.0               # 页边距

    # ── 字体 (PDF 输出) ───────────────────────────────────────────────────────
    # 支持中文的 TrueType 字体路径（Windows 系统字体）
    cn_font_path: str = r"C:\Windows\Fonts\simhei.ttf"   # 黑体
    cn_font_name: str = "SimHei"
    en_font_path: str = r"C:\Windows\Fonts\times.ttf"    # Times New Roman
    en_font_name: str = "TimesNewRoman"

    # ── 翻译批量 ──────────────────────────────────────────────────────────────
    translation_batch_size: int = 10      # 每批翻译的文本块数量
    max_retries: int = 3                   # 翻译失败重试次数

    @classmethod
    def from_env(cls) -> "Config":
        """从环境变量读取配置"""
        cfg = cls()
        if os.environ.get("OPENAI_API_KEY") and not os.environ.get("ANTHROPIC_API_KEY"):
            cfg.api_provider = "openai"
            cfg.api_key = os.environ["OPENAI_API_KEY"]
            cfg.api_url = os.environ.get("OPENAI_API_URL", "https://api.openai.com/v1")
            cfg.model = os.environ.get("LLM_MODEL", "gpt-4o")
        else:
            cfg.api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        return cfg
