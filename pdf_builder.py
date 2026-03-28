# -*- coding: utf-8 -*-
"""
PDF 输出构建器：使用 reportlab，按 bbox 坐标精确还原版面
"""

import io
import os
from typing import List, Tuple

from PIL import Image as PILImage

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import inch, cm, mm
    from reportlab.lib import colors
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import Paragraph, Frame, Table, TableStyle, Image
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

from layout_analyzer import PageLayout, Region, TYPE_TITLE, TYPE_TEXT, \
    TYPE_TABLE, TYPE_FIGURE, TYPE_FIG_CAP, TYPE_TBL_CAP
from bs4 import BeautifulSoup


class PDFBuilder:
    """将翻译后版面数据用 reportlab 输出为 PDF"""

    def __init__(self, cn_font_path: str = r"C:\Windows\Fonts\simhei.ttf",
                 cn_font_name: str = "SimHei"):
        if not REPORTLAB_AVAILABLE:
            raise ImportError("请安装 reportlab: pip install reportlab")
        self.cn_font_name = cn_font_name
        self._register_fonts(cn_font_path, cn_font_name)
        self._styles = self._build_styles()

    def _register_fonts(self, cn_font_path: str, cn_font_name: str):
        """注册中文字体"""
        if cn_font_name not in pdfmetrics.getRegisteredFontNames():
            if os.path.exists(cn_font_path):
                pdfmetrics.registerFont(TTFont(cn_font_name, cn_font_path))
            else:
                # 尝试其他常见路径
                fallbacks = [
                    r"C:\Windows\Fonts\msyh.ttc",
                    r"C:\Windows\Fonts\simsun.ttc",
                    "/usr/share/fonts/truetype/arphic/uming.ttc",
                ]
                for fb in fallbacks:
                    if os.path.exists(fb):
                        pdfmetrics.registerFont(TTFont(cn_font_name, fb))
                        print(f"  使用后备中文字体: {fb}")
                        break
                else:
                    print("  警告：未找到中文字体，中文可能无法正常显示")
                    self.cn_font_name = "Helvetica"

    def _build_styles(self) -> dict:
        fn = self.cn_font_name
        return {
            "title": ParagraphStyle(
                "Title", fontName=fn, fontSize=13, leading=16,
                alignment=TA_CENTER, spaceAfter=8, spaceBefore=6,
            ),
            "body": ParagraphStyle(
                "Body", fontName=fn, fontSize=10, leading=14,
                alignment=TA_JUSTIFY, spaceAfter=4,
                firstLineIndent=20,
            ),
            "caption": ParagraphStyle(
                "Caption", fontName=fn, fontSize=9, leading=12,
                alignment=TA_CENTER, spaceAfter=4,
            ),
            "table_cell": ParagraphStyle(
                "TableCell", fontName=fn, fontSize=8, leading=11,
            ),
        }

    def build(self, layouts: List[PageLayout], output_path: str):
        page_w, page_h = A4  # 595 x 842 pt
        margin = 56          # ~2cm in points

        c = canvas.Canvas(output_path, pagesize=A4)

        for layout in layouts:
            self._draw_page(c, layout, page_w, page_h, margin)
            c.showPage()

        c.save()
        print(f"  PDF 已保存: {output_path}")

    def _draw_page(self, c: "canvas.Canvas", layout: PageLayout,
                   page_w: float, page_h: float, margin: float):
        """绘制单页内容"""
        # 缩放比：将图像坐标映射到 PDF 坐标
        scale_x = (page_w - 2 * margin) / layout.img_width
        scale_y = (page_h - 2 * margin) / layout.img_height

        regions = layout.get_sorted_regions()

        for region in regions:
            x1, y1, x2, y2 = region.bbox
            # PDF 坐标系：原点在左下角，y 轴向上
            pdf_x1 = margin + x1 * scale_x
            pdf_y2 = page_h - margin - y1 * scale_y   # 顶边（PDF 坐标）
            pdf_x2 = margin + x2 * scale_x
            pdf_y1 = page_h - margin - y2 * scale_y   # 底边（PDF 坐标）
            w = pdf_x2 - pdf_x1
            h = pdf_y2 - pdf_y1

            if h <= 2 or w <= 2:
                continue

            self._draw_region(c, region, pdf_x1, pdf_y1, w, h)

    def _draw_region(self, c, region: Region,
                     x: float, y: float, w: float, h: float):
        """在指定位置绘制一个区域"""
        if region.type == TYPE_FIGURE:
            self._draw_image(c, region, x, y, w, h)

        elif region.type == TYPE_TABLE:
            self._draw_table(c, region, x, y, w, h)

        elif region.type == TYPE_TITLE:
            text = region.translated_text or region.text
            self._draw_text_frame(c, text, x, y, w, h, style="title")

        else:
            text = region.translated_text or region.text
            if text:
                self._draw_text_frame(c, text, x, y, w, h, style="body")

    def _draw_text_frame(self, c, text: str, x: float, y: float,
                         w: float, h: float, style: str = "body"):
        """用 reportlab Frame + Paragraph 绘制文本"""
        para_style = self._styles.get(style, self._styles["body"])
        # 转义 XML 特殊字符
        safe_text = (text.replace("&", "&amp;").replace("<", "&lt;")
                         .replace(">", "&gt;"))
        para = Paragraph(safe_text, para_style)

        frame = Frame(x, y, w, h, leftPadding=0, rightPadding=0,
                      topPadding=0, bottomPadding=0, showBoundary=0)
        try:
            frame.addFromList([para], c)
        except Exception as e:
            # 退回简单文字绘制
            c.setFont(self.cn_font_name, 9)
            c.drawString(x + 2, y + h - 12, text[:80])

    def _draw_image(self, c, region: Region,
                    x: float, y: float, w: float, h: float):
        """绘制图片"""
        if region.image is None:
            return
        try:
            buf = io.BytesIO()
            region.image.save(buf, format="PNG")
            buf.seek(0)
            c.drawImage(
                c._imageCache.get(buf) if hasattr(c, "_imageCache") else buf,
                x, y, w, h, preserveAspectRatio=True, mask="auto"
            )
        except Exception:
            try:
                buf.seek(0)
                from reportlab.lib.utils import ImageReader
                img_reader = ImageReader(buf)
                c.drawImage(img_reader, x, y, w, h,
                            preserveAspectRatio=True, mask="auto")
            except Exception as e:
                print(f"    图片绘制失败: {e}")

    def _draw_table(self, c, region: Region,
                    x: float, y: float, w: float, h: float):
        """绘制翻译后的表格"""
        html = region.translated_html or region.table_html
        if not html:
            return

        soup = BeautifulSoup(html, "html.parser")
        rows_html = soup.find_all("tr")
        if not rows_html:
            return

        fn = self.cn_font_name
        cell_style = self._styles["table_cell"]

        data = []
        for tr in rows_html:
            row_data = []
            for td in tr.find_all(["td", "th"]):
                text = td.get_text(strip=True)
                safe = (text.replace("&", "&amp;").replace("<", "&lt;")
                             .replace(">", "&gt;"))
                p = Paragraph(safe, cell_style)
                row_data.append(p)
            if row_data:
                data.append(row_data)

        if not data:
            return

        # 均分列宽
        num_cols = max(len(row) for row in data)
        col_widths = [w / num_cols] * num_cols

        tbl = Table(data, colWidths=col_widths)
        tbl.setStyle(TableStyle([
            ("FONTNAME",  (0, 0), (-1, -1), fn),
            ("FONTSIZE",  (0, 0), (-1, -1), 8),
            ("GRID",      (0, 0), (-1, -1), 0.5, colors.grey),
            ("BACKGROUND",(0, 0), (-1, 0),  colors.lightgrey),
            ("VALIGN",    (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",(0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))

        try:
            tbl_w, tbl_h = tbl.wrapOn(c, w, h)
            tbl.drawOn(c, x, y + h - tbl_h)
        except Exception as e:
            print(f"    表格绘制失败: {e}")
