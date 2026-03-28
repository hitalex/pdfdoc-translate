# -*- coding: utf-8 -*-
"""
PDF → 逐页图像
使用 PyMuPDF (fitz)，无需安装 poppler
"""

from pathlib import Path
from typing import List, Tuple
import numpy as np

try:
    import fitz  # PyMuPDF
except ImportError:
    raise ImportError("请安装 PyMuPDF: pip install PyMuPDF")

try:
    from PIL import Image
except ImportError:
    raise ImportError("请安装 Pillow: pip install Pillow")


class PDFProcessor:
    """将 PDF 文件转换为逐页 PIL 图像及元数据"""

    def __init__(self, dpi: int = 200):
        self.dpi = dpi
        self._scale = dpi / 72.0  # PDF 原生单位为 72 DPI

    def convert(self, pdf_path: str) -> List[dict]:
        """
        转换 PDF 为图像列表。

        Returns:
            [
                {
                    'page_num': 1,
                    'image': PIL.Image,
                    'width': int,   # 图像宽度（像素）
                    'height': int,  # 图像高度（像素）
                    'pdf_width': float,   # PDF 页面原始宽度（点）
                    'pdf_height': float,
                },
                ...
            ]
        """
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF 文件不存在: {pdf_path}")

        doc = fitz.open(str(pdf_path))
        pages = []

        print(f"  共 {len(doc)} 页，DPI={self.dpi}")
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            mat = fitz.Matrix(self._scale, self._scale)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            # 转换为 PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            pages.append({
                'page_num': page_idx + 1,
                'image': img,
                'width': pix.width,
                'height': pix.height,
                'pdf_width': page.rect.width,
                'pdf_height': page.rect.height,
            })
            print(f"  第 {page_idx + 1} 页: {pix.width}x{pix.height} px")

        doc.close()
        return pages

    @staticmethod
    def image_to_numpy(img: Image.Image) -> np.ndarray:
        """PIL Image → numpy array (RGB)"""
        return np.array(img)
