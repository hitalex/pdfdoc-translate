# -*- coding: utf-8 -*-
"""
版面分析：使用 PaddleOCR PPStructure 检测文字/标题/表格/图片区域
并推断双栏布局的阅读顺序

兼容 paddleocr 2.10.x (PPStructure)
"""

from typing import List, Optional
import numpy as np
from PIL import Image

try:
    from paddleocr import PPStructure
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

# PPStructure type 字段 → 内部类型 映射
_TYPE_MAP = {
    "text":            TYPE_TEXT,
    "title":           TYPE_TITLE,
    "table":           TYPE_TABLE,
    "figure":          TYPE_FIGURE,
    "figure_caption":  TYPE_FIG_CAP,
    "table_caption":   TYPE_TBL_CAP,
}


class Region:
    """一个版面区域"""
    def __init__(self, region_type: str, bbox: List[int], text: str = "",
                 table_html: str = "", image: Optional[Image.Image] = None,
                 column: str = "full"):
        self.type = region_type
        self.bbox = bbox                   # [x1, y1, x2, y2]
        self.text = text
        self.table_html = table_html
        self.image = image
        self.column = column               # "left" | "right" | "full"
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
                f"column={self.column}, text={self.text[:30]!r})")

    def summary(self) -> str:
        """单行摘要，用于调试报告"""
        if self.type == TYPE_TABLE:
            from bs4 import BeautifulSoup
            rows = len(BeautifulSoup(self.table_html, "html.parser").find_all("tr")) if self.table_html else 0
            detail = f"(表格 {rows} 行)"
        elif self.type == TYPE_FIGURE:
            sz = f"{self.image.size[0]}x{self.image.size[1]}" if self.image else "无图"
            detail = f"(图片 {sz})"
        else:
            preview = self.text.replace('\n', ' ')[:80]
            detail = f'"{preview}"' if preview else "(空文本)"
        return f"[{self.type:<14} {self.column:<5} bbox={self.bbox}]  {detail}"


class PageLayout:
    """一页的版面分析结果"""
    def __init__(self, page_num: int, img_width: int, img_height: int):
        self.page_num = page_num
        self.img_width = img_width
        self.img_height = img_height
        self.regions: List[Region] = []
        self.is_two_column = False

    def get_sorted_regions(self) -> List[Region]:
        """按阅读顺序返回区域（左栏→右栏，栏内从上到下）"""
        if not self.is_two_column:
            return sorted(self.regions, key=lambda r: r.bbox[1])

        full  = [r for r in self.regions if r.column == "full"]
        left  = sorted([r for r in self.regions if r.column == "left"],
                       key=lambda r: r.bbox[1])
        right = sorted([r for r in self.regions if r.column == "right"],
                       key=lambda r: r.bbox[1])

        return _interleave_columns(full, left, right)


def _interleave_columns(full: List[Region], left: List[Region],
                        right: List[Region]) -> List[Region]:
    """将全宽区域与双栏区域按垂直位置交织排列"""
    all_items = (
        [(r, "full")  for r in full]  +
        [(r, "left")  for r in left]  +
        [(r, "right") for r in right]
    )
    all_items.sort(key=lambda x: x[0].bbox[1])

    result = []
    i = 0
    while i < len(all_items):
        r, col = all_items[i]
        if col == "full":
            result.append(r)
            i += 1
        else:
            band_bottom = r.bbox[3]
            band_left  = [r] if col == "left"  else []
            band_right = [r] if col == "right" else []
            j = i + 1
            while j < len(all_items):
                nr, ncol = all_items[j]
                if ncol == "full":
                    break
                if nr.bbox[1] < band_bottom + 50:
                    band_bottom = max(band_bottom, nr.bbox[3])
                    (band_left if ncol == "left" else band_right).append(nr)
                    j += 1
                else:
                    break
            result.extend(sorted(band_left,  key=lambda r: r.bbox[1]))
            result.extend(sorted(band_right, key=lambda r: r.bbox[1]))
            i = j
    return result


class LayoutAnalyzer:
    """使用 PaddleOCR PPStructure 分析 PDF 页面图像的版面（paddleocr 2.10.x）"""

    def __init__(self, lang: str = "en", use_gpu: bool = False):
        if not PADDLE_AVAILABLE:
            raise ImportError(
                "请安装 PaddleOCR:\n"
                "  pip install paddlepaddle==3.0.0 paddleocr==2.10.0"
            )
        self.lang = lang
        print("  初始化 PaddleOCR PPStructure...")
        self._engine = PPStructure(
            table=True,
            ocr=True,
            lang=lang,
            use_gpu=use_gpu,
            show_log=False,
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

        # PPStructure.__call__(img) 返回 list of dict
        raw_results = self._engine(img_np)

        for item in raw_results:
            region = self._parse_item(item, img)
            if region:
                layout.regions.append(region)

        self._assign_columns(layout)
        return layout

    def _parse_item(self, item: dict, page_img: Image.Image) -> Optional[Region]:
        """
        将 PPStructure 的单个结果转换为 Region。

        PPStructure 2.10.x 返回格式:
        {
            'type': 'text' | 'title' | 'table' | 'figure' | ...,
            'bbox': [x1, y1, x2, y2],
            'img':  numpy array,
            'res':  list of ocr dicts (text/title) 或 HTML str (table)
        }
        """
        raw_type = str(item.get("type", "unknown")).lower()
        bbox_raw = item.get("bbox", None)

        if bbox_raw is None:
            return None

        try:
            x1, y1, x2, y2 = int(bbox_raw[0]), int(bbox_raw[1]), \
                              int(bbox_raw[2]), int(bbox_raw[3])
        except (TypeError, IndexError, ValueError):
            return None

        # 裁切区域图像
        try:
            region_img = page_img.crop((x1, y1, x2, y2))
        except Exception:
            region_img = None

        region_type = _TYPE_MAP.get(raw_type, TYPE_TEXT)
        res = item.get("res", None)

        if region_type == TYPE_TABLE:
            html = self._extract_table_html(res)
            return Region(region_type, [x1, y1, x2, y2],
                          table_html=html, image=region_img)

        elif region_type == TYPE_FIGURE:
            return Region(region_type, [x1, y1, x2, y2], image=region_img)

        else:
            text = self._extract_text(res)
            return Region(region_type, [x1, y1, x2, y2],
                          text=text, image=region_img)

    @staticmethod
    def _extract_text(res) -> str:
        """从 PPStructure OCR 结果列表中提取纯文本"""
        if not res:
            return ""
        if isinstance(res, str):
            return res
        lines = []
        for item in res:
            if isinstance(item, dict):
                t = item.get("text", "")
                if t:
                    lines.append(t)
            elif isinstance(item, (list, tuple)):
                # 格式: [[bbox_pts], [text, confidence]]
                try:
                    inner = item[1]
                    if isinstance(inner, (list, tuple)):
                        lines.append(str(inner[0]))
                    else:
                        lines.append(str(inner))
                except (IndexError, TypeError):
                    pass
        return " ".join(lines)

    @staticmethod
    def _extract_table_html(res) -> str:
        """从 PPStructure 表格结果中提取 HTML"""
        if not res:
            return ""
        if isinstance(res, str):
            return res
        if isinstance(res, dict):
            return res.get("html", "") or res.get("pred_html", "")
        # 某些版本返回包含 html 字段的 dict
        if isinstance(res, list) and res:
            first = res[0]
            if isinstance(first, dict):
                return first.get("html", "")
        return ""

    @staticmethod
    def _assign_columns(layout: PageLayout):
        """判断是否双栏，分配 column 属性"""
        if len(layout.regions) < 2:
            for r in layout.regions:
                r.column = "full"
            return

        page_mid = layout.img_width / 2
        centers = [r.center_x for r in layout.regions]
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


def dump_layout_report(layouts: List[PageLayout], txt_path: str):
    """
    将 PPStructure 解析结果输出为可读报告文件。
    每行对应一个区域，按阅读顺序排列，便于与 DOCX 内容对照核查。
    """
    lines = []
    total_regions = sum(len(l.regions) for l in layouts)
    lines.append(f"PPStructure 解析报告  共 {len(layouts)} 页，{total_regions} 个区域")
    lines.append("=" * 80)

    for layout in layouts:
        col_info = "双栏" if layout.is_two_column else "单栏"
        lines.append(f"\n=== 第 {layout.page_num} 页  ({layout.img_width}x{layout.img_height}px, {col_info}) ===")
        sorted_regions = layout.get_sorted_regions()
        if not sorted_regions:
            lines.append("  (无识别区域)")
            continue
        for i, r in enumerate(sorted_regions, 1):
            lines.append(f"  {i:02d}. {r.summary()}")

    lines.append("\n" + "=" * 80)
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"  解析报告已保存: {txt_path}")
