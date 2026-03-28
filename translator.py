# -*- coding: utf-8 -*-
"""
LLM 翻译模块：支持 Anthropic Claude 和 OpenAI 兼容接口
"""

import re
import time
from typing import List, Optional
from bs4 import BeautifulSoup

from config import Config
from layout_analyzer import PageLayout, Region, TYPE_TABLE, TRANSLATABLE_TYPES


class Translator:
    """调用 LLM API 进行翻译"""

    def __init__(self, config: Config):
        self.config = config
        self._client = self._init_client()

    def _init_client(self):
        if self.config.api_provider == "anthropic":
            try:
                import anthropic
                kwargs = {}
                if self.config.api_url:
                    kwargs["base_url"] = self.config.api_url
                return anthropic.Anthropic(api_key=self.config.api_key, **kwargs)
            except ImportError:
                raise ImportError("请安装: pip install anthropic")
        else:
            try:
                from openai import OpenAI
                return OpenAI(
                    api_key=self.config.api_key,
                    base_url=self.config.api_url or "https://api.openai.com/v1",
                )
            except ImportError:
                raise ImportError("请安装: pip install openai")

    # ── 核心翻译方法 ──────────────────────────────────────────────────────────

    def translate_text(self, text: str) -> str:
        """翻译纯文本段落"""
        if not text.strip():
            return text
        prompt = self._build_text_prompt(text)
        return self._call_llm(prompt)

    def translate_texts_batch(self, texts: List[str]) -> List[str]:
        """批量翻译文本列表（减少 API 调用次数）"""
        if not texts:
            return []

        # 过滤空文本
        non_empty = [(i, t) for i, t in enumerate(texts) if t.strip()]
        if not non_empty:
            return texts

        indices, chunks = zip(*non_empty)

        # 用分隔符合并为单次请求
        separator = "\n\n<<<SEGMENT_BREAK>>>\n\n"
        combined = separator.join(chunks)
        prompt = self._build_batch_prompt(combined, len(chunks))

        raw = self._call_llm(prompt)

        # 按分隔符拆分结果
        parts = re.split(r'<<<SEGMENT_BREAK>>>', raw)
        parts = [p.strip() for p in parts]

        # 若拆分数量不对，退回逐条翻译
        if len(parts) != len(chunks):
            print(f"    批量翻译分割异常({len(parts)}!={len(chunks)})，退回逐条翻译")
            parts = [self.translate_text(t) for t in chunks]

        # 还原到原始索引
        result = list(texts)
        for idx, translated in zip(indices, parts):
            result[idx] = translated
        return result

    def translate_table_html(self, html: str) -> str:
        """翻译表格（保留 HTML 结构，只翻译单元格文字）"""
        if not html.strip():
            return html

        soup = BeautifulSoup(html, "html.parser")
        cells = soup.find_all(["td", "th"])

        if not cells:
            return html

        # 提取所有单元格文本并批量翻译
        cell_texts = [c.get_text(strip=True) for c in cells]
        translated_texts = self.translate_texts_batch(cell_texts)

        # 写回单元格
        for cell, new_text in zip(cells, translated_texts):
            cell.clear()
            cell.string = new_text

        return str(soup)

    # ── 翻译版面 ─────────────────────────────────────────────────────────────

    def translate_layout(self, layouts: List[PageLayout]) -> List[PageLayout]:
        """
        翻译所有页面版面中的可翻译区域。
        就地修改 Region.translated_text / translated_html，返回相同列表。
        """
        total_regions = sum(
            1 for layout in layouts
            for r in layout.regions
            if r.needs_translation
        )
        print(f"  共 {total_regions} 个可翻译区域")

        done = 0
        for layout in layouts:
            # 收集该页所有文本区域（非表格）批量翻译
            text_regions = [
                r for r in layout.regions
                if r.needs_translation and r.type != TYPE_TABLE
            ]
            table_regions = [
                r for r in layout.regions
                if r.needs_translation and r.type == TYPE_TABLE
            ]

            # 批量翻译文本
            if text_regions:
                texts = [r.text for r in text_regions]
                translated = self.translate_texts_batch(texts)
                for r, t in zip(text_regions, translated):
                    r.translated_text = t
                done += len(text_regions)
                print(f"    第 {layout.page_num} 页文本：{len(text_regions)} 块已翻译")

            # 逐表格翻译
            for r in table_regions:
                r.translated_html = self.translate_table_html(r.table_html)
                done += 1
                print(f"    第 {layout.page_num} 页表格：已翻译")

        return layouts

    # ── Prompt 构建 ───────────────────────────────────────────────────────────

    def _build_text_prompt(self, text: str) -> str:
        src = self.config.source_lang.upper()
        tgt = self.config.target_lang_name
        return (
            f"将以下{src}学术文本翻译为{tgt}。"
            f"要求：保持学术风格，保留专业术语，不要添加任何解释或额外内容，"
            f"只输出翻译结果。\n\n{text}"
        )

    def _build_batch_prompt(self, combined: str, count: int) -> str:
        src = self.config.source_lang.upper()
        tgt = self.config.target_lang_name
        return (
            f"将以下{count}段{src}学术文本分别翻译为{tgt}。"
            f"各段之间用 <<<SEGMENT_BREAK>>> 分隔，翻译结果也必须用 <<<SEGMENT_BREAK>>> "
            f"分隔对应段落。不要添加任何解释，只输出翻译结果。\n\n{combined}"
        )

    # ── LLM 调用 ─────────────────────────────────────────────────────────────

    def _call_llm(self, prompt: str) -> str:
        for attempt in range(self.config.max_retries):
            try:
                if self.config.api_provider == "anthropic":
                    return self._call_anthropic(prompt)
                else:
                    return self._call_openai(prompt)
            except Exception as e:
                if attempt < self.config.max_retries - 1:
                    wait = 2 ** attempt
                    print(f"    翻译失败（{e}），{wait}s 后重试...")
                    time.sleep(wait)
                else:
                    print(f"    翻译失败，返回原文: {e}")
                    return prompt  # 失败时返回原文

    def _call_anthropic(self, prompt: str) -> str:
        msg = self._client.messages.create(
            model=self.config.model,
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
        )
        return msg.content[0].text.strip()

    def _call_openai(self, prompt: str) -> str:
        resp = self._client.chat.completions.create(
            model=self.config.model,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=4096,
        )
        return resp.choices[0].message.content.strip()
