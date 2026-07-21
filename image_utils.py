#!/usr/bin/env python3
"""
通用圖像工具模組
- HEIC 註冊支援（iPhone 預設格式）
- 統一 JPEG 轉換（JPG/JPEG 跳過）
- 從 nanonets_core.py 抽出並擴充
"""

import os
from pathlib import Path
from PIL import Image

# ── 註冊 HEIC 支援 ──
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    _HEIF_OK = True
except ImportError:
    _HEIF_OK = False


def convert_to_jpeg(image_path: str) -> str:
    """將任何非 JPEG 圖片轉為 JPEG。JPG/JPEG 直接回傳原路徑。

    支援格式：HEIC, PNG, BMP, TIFF, WebP, GIF, 及其他 PIL 支援格式。
    轉換時自動處理透明通道（RGBA/LA/P → 白底 RGB）。

    Args:
        image_path: 輸入圖片路徑。

    Returns:
        JPEG 路徑（已是 JPEG 則回傳原路徑，否則回傳新 .jpeg 路徑）。
    """
    path = Path(image_path)
    suffix = path.suffix.lower()

    # JPG/JPEG 直接跳過
    if suffix in ('.jpg', '.jpeg'):
        return image_path

    jpeg_path = str(path.with_suffix('.jpeg'))

    with Image.open(image_path) as img:
        # 處理透明通道
        if img.mode in ('RGBA', 'LA', 'P'):
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')

        img.save(jpeg_path, 'JPEG', quality=95)

    return jpeg_path