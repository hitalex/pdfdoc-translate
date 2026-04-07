# -*- coding: utf-8 -*-
"""
版面分析：使用 PaddleOCR PPStructureV3 检测文字/标题/表格/图片区域
并推断双栏布局的阅读顺序

兼容 paddleocr 3.4.x (PPStructureV3)
"""

import os
from typing import List, Optional
import numpy as np
from PIL import Image

try:
    from paddleocr import PPStructureV3
    from paddlex.inference.models.common.static_infer import PaddleInfer as _PaddleInfer

    # Windows 上 PaddlePaddle 3.x 的 PIR 新执行器（new IR）与 oneDNN 存在兼容问题：
    # paddlex 默认 enable_new_ir=True，PIR executor 内部仍调用 oneDNN 指令，
    # 导致 ConvertPirAttribute2RuntimeAttribute not support DoubleAttribute 错误。
    # Patch Python 层 PaddleInfer._create，在创建 Config 前强制关闭 PIR。
    _orig_create = _PaddleInfer._create

    def _patched_create(self):
        self._option.enable_new_ir = False
        return _orig_create(self)

    _PaddleInfer._create = _patched_create

    PADDLE_AVAILABLE = True
except ImportError:
    PADDLE_AVAILABLE = False


# ── 区域类型常量 ──────────────────────────────────────────────────────────────
TYPE_TEXT    = "text"
TYPE_TITLE   = "title"
TYPE_TABLE   = "table"
TYPE_FIGURE  = "figure"
TYPE_FIG_CAP = "figure_caption"
TYPE_TBL_CAP = "table_caption"
TYPE_UNKNOWN = "unknown"

TRANSLATABLE_TYPES = {TYPE_TEXT, TYPE_TITLE, TYPE_TABLE, TYPE_FIG_CAP, TYPE_TBL_CAP}

# PPStructureV3 block_label → 内部类型 映射
_TYPE_MAP = {
    "text":            TYPE_TEXT,
    "title":           TYPE_TITLE,
    "table":           TYPE_TABLE,
    "figure":          TYPE_FIGURE,
    "image":           TYPE_FIGURE,
    "figure_caption":  TYPE_FIG_CAP,
    "table_caption":   TYPE_TBL_CAP,
    "abstract":        TYPE_TEXT,
    "references":      TYPE_TEXT,
    "footnote":        TYPE_TEXT,
    "formula":         TYPE_TEXT,
    "seal":            TYPE_TEXT,
    "footer":          TYPE_TEXT,
    "header":          TYPE_TEXT,
    "chart":           TYPE_FIGURE,
    "page_number":     TYPE_UNKNOWN,
}


class Region:
    """一个版面区域"""
    def __init__(self, region_type: str, bbox: List[int], text: str = "",
                 table_html: str = "", image: Optional[Image.Image] = None,
                 column: str = "full", heading_level: int = 0,
                 block_order: int = 0):
        self.type = region_type
        self.bbox = bbox                    # [x1, y1, x2, y2]
        self.text = text
        self.table_html = table_html
        self.image = image
        self.column = column                # "left" | "right" | "full"
        self.heading_level = heading_level  # 0=非标题, 1/2/3=H1/H2/H3
        self.block_order = block_order      # PPStructureV3 给出的阅读顺序
        self.translated_text = ""
        self.translated_html = ""

    @property
    def needs_translation(self) -> bool:
        return self.type in TRANSLATABLE_TYPES

    @property
    def center_x(self) -> float:
        return (self.bbox[0] + self.bbox[2]) / 2

    @property
    def center_y(self) -> float:
        return (self.bbox[1] + self.bbox[3]) / 2

    @property
    def width(self) -> int:
        return self.bbox[2] - self.bbox[0]

    @property
    def height(self) -> int:
        return self.bbox[3] - self.bbox[1]

    def __repr__(self):
        return (f"Region(type={self.type}, bbox={self.bbox}, "
                f"column={self.column}, heading={self.heading_level}, "
                f"text={self.text[:30]!r})")


class PageLayout:
    """一页的版面分析结果"""
    def __init__(self, page_num: int, img_width: int, img_height: int):
        self.page_num = page_num
        self.img_width = img_width
        self.img_height = img_height
        self.regions: List[Region] = []
        self.is_two_column = False

    def get_sorted_regions(self) -> List[Region]:
        """按阅读顺序返回区域（优先使用 PPStructureV3 提供的 block_order）"""
        return sorted(self.regions, key=lambda r: (r.block_order, r.bbox[1]))


class LayoutAnalyzer:
    """使用 PaddleOCR PPStructureV3 分析 PDF 页面图像的版面（paddleocr 3.4.x）"""

    def __init__(self, lang: str = "en", use_gpu: bool = False):
        if not PADDLE_AVAILABLE:
            raise ImportError(
                "请安装 PaddleOCR:\n"
                "  pip install paddlepaddle==3.3.1 paddleocr==3.4.0"
            )
        self.lang = lang
        device = "gpu:0" if use_gpu else "cpu"
        print("  初始化 PaddleOCR PPStructureV3...")
        self._engine = PPStructureV3(
            use_table_recognition=True,
            device=device,
        )

    def analyze(self, pages: List[dict]) -> List[PageLayout]:
        layouts = []
        for page in pages:
            print(f"  分析第 {page['page_num']} 页...")
            layout = self._analyze_page(page)
            layouts.append(layout)
        return layouts

    def _analyze_page(self, page: dict) -> PageLayout:
        img = page['image']
        img_np = np.array(img)
        w, h = page['width'], page['height']

        layout = PageLayout(page['page_num'], w, h)

        # PPStructureV3.predict() 返回 list of dict，每个 dict 对应一张图
        results = list(self._engine.predict(img_np))
        if not results:
            return layout

        res = results[0]
        for block in res.get('parsing_res_list', []):
            region = self._parse_block(block, img, h)
            if region:
                layout.regions.append(region)

        self._assign_columns(layout)
        return layout

    def _parse_block(self, block: dict, page_img: Image.Image,
                     page_height: int) -> Optional[Region]:
        """
        将 PPStructureV3 parsing_res_list 中的单个块转换为 Region。

        PPStructureV3 3.4.x 块格式:
        {
            'block_label':   str,       # 区域类型
            'block_bbox':    [x1,y1,x2,y2],
            'block_content': str|dict,  # 文字内容
            'block_html':    str,       # 仅表格：HTML
            'block_order':   int,       # 阅读顺序
        }
        """
        raw_label = str(block.get("block_label", "unknown")).lower()
        bbox_raw  = block.get("block_bbox", None)

        if bbox_raw is None:
            return None

        try:
            x1, y1, x2, y2 = (int(bbox_raw[0]), int(bbox_raw[1]),
                               int(bbox_raw[2]), int(bbox_raw[3]))
        except (TypeError, IndexError, ValueError):
            return None

        try:
            region_img = page_img.crop((x1, y1, x2, y2))
        except Exception:
            region_img = None

        region_type = _TYPE_MAP.get(raw_label, TYPE_TEXT)
        block_order = int(block.get("block_order", 0))
        content     = block.get("block_content", "")

        if region_type == TYPE_TABLE:
            html = block.get("block_html", "")
            if not html and isinstance(content, str):
                html = content
            return Region(region_type, [x1, y1, x2, y2],
                          table_html=html, image=region_img,
                          block_order=block_order)

        elif region_type == TYPE_FIGURE:
            return Region(region_type, [x1, y1, x2, y2],
                          image=region_img, block_order=block_order)

        else:
            text = self._extract_text(content)
            heading_level = 0
            if region_type == TYPE_TITLE:
                heading_level = self._estimate_heading_level(
                    y2 - y1, text, page_height
                )
            return Region(region_type, [x1, y1, x2, y2],
                          text=text, image=region_img,
                          heading_level=heading_level,
                          block_order=block_order)

    @staticmethod
    def _extract_text(content) -> str:
        """从 block_content 提取纯文本"""
        if not content:
            return ""
        if isinstance(content, str):
            return content.strip()
        if isinstance(content, dict):
            return content.get("text", "").strip()
        if isinstance(content, list):
            parts = []
            for item in content:
                if isinstance(item, str):
                    parts.append(item)
                elif isinstance(item, dict):
                    parts.append(item.get("text", ""))
            return " ".join(p for p in parts if p).strip()
        return str(content).strip()

    @staticmethod
    def _estimate_heading_level(bbox_height: int, text: str,
                                page_height: int) -> int:
        """
        根据 bbox 行高与页面高度的比值推断标题层级（1/2/3）。
        使用相对比值而非绝对像素，对不同 DPI 具有鲁棒性。

        参考（A4 @ 200dpi，页面高 ≈ 2338px）：
          H1: 行高 > 3.0%  → 约 20pt+
          H2: 行高 > 1.8%  → 约 14pt+
          H3: 其余 title 区域
        """
        line_count  = max(1, text.count('\n') + 1)
        line_height = bbox_height / line_count
        ratio       = line_height / page_height if page_height > 0 else 0
        if ratio > 0.030:
            return 1
        elif ratio > 0.018:
            return 2
        else:
            return 3

    @staticmethod
    def _assign_columns(layout: PageLayout):
        """判断是否双栏，为每个区域分配 column 属性"""
        if len(layout.regions) < 2:
            for r in layout.regions:
                r.column = "full"
            return

        page_mid = layout.img_width / 2
        centers  = [r.center_x for r in layout.regions]
        has_left  = any(c < page_mid * 0.75 for c in centers)
        has_right = any(c > page_mid * 1.25 for c in centers)

        if not (has_left and has_right):
            layout.is_two_column = False
            for r in layout.regions:
                r.column = "full"
            return

        layout.is_two_column = True
        for r in layout.regions:
            if r.width > layout.img_width * 0.6:
                r.column = "full"
            elif r.center_x < page_mid:
                r.column = "left"
            else:
                r.column = "right"
