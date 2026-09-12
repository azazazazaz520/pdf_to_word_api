"""页面内嵌图片的提取与编码。

按原始格式直通、按像素上限缩放，并在照片类图片上转为 JPEG；截图、
透明图与图形保留 PNG。
"""

from __future__ import annotations
from io import BytesIO
from typing import Any
from PIL import Image
from .models import PdfImageBlock
import math
import pypdfium2 as pdfium

def _is_large_photo_like(image: Image.Image) -> bool:
    """用缩略图颜色数量粗略判断图片是否更适合 JPEG。"""
    if image.mode not in {"RGB", "L"}:
        return False
    if image.width * image.height < _JPEG_PHOTO_MIN_PIXELS:
        return False
    thumbnail = image.copy()
    try:
        thumbnail.thumbnail((256, 256))
        colors = thumbnail.getcolors(maxcolors=256 * 256)
    finally:
        thumbnail.close()
    if colors is None:
        return True
    return len(colors) > _JPEG_PHOTO_MIN_COLORS


def _encode_image_bytes(
    image: Image.Image,
    *,
    max_pixels: int,
    png_optimize: bool,
    jpeg_quality: int,
) -> tuple[bytes, str]:
    """按像素上限和图片特征选择 PNG 或 JPEG，并返回编码结果。"""
    if image.mode not in {"RGB", "RGBA", "L"}:
        image = image.convert("RGB")
    pixel_count = image.width * image.height
    if pixel_count > max_pixels:
        scale = math.sqrt(max_pixels / max(pixel_count, 1))
        target_width = max(1, int(image.width * scale))
        target_height = max(1, int(image.height * scale))
        while target_width * target_height > max_pixels:
            if target_width >= target_height:
                target_width -= 1
            else:
                target_height -= 1
        image = image.resize(
            (max(1, target_width), max(1, target_height)),
            Image.Resampling.LANCZOS,
        )
    output = BytesIO()
    if _is_large_photo_like(image):
        image.save(
            output,
            format="JPEG",
            quality=max(1, min(100, jpeg_quality)),
            optimize=False,
        )
        return output.getvalue(), "image/jpeg"
    image.save(output, format="PNG", optimize=bool(png_optimize))
    return output.getvalue(), "image/png"


def _image_bytes_from_object(
    image_object: Any,
    *,
    max_pixels: int,
    png_optimize: bool,
    jpeg_quality: int,
) -> tuple[bytes, str] | None:
    """提取图片对象；原始 JPEG 尽量直通，其他图片按特征编码。"""
    try:
        filters = image_object.get_filters(skip_simple=True)
    except Exception:
        filters = []
    if "DCTDecode" in filters:
        try:
            jpeg_data = bytes(image_object.get_data(decode_simple=True))
        except Exception:
            jpeg_data = b""
        if jpeg_data.startswith(b"\xff\xd8") and jpeg_data.endswith(
            b"\xff\xd9"
        ):
            try:
                px_size = image_object.get_px_size()
                pixel_count = int(px_size[0]) * int(px_size[1])
            except Exception:
                pixel_count = 0
            if 0 < pixel_count <= max_pixels:
                return jpeg_data, "image/jpeg"
    try:
        bitmap = image_object.get_bitmap()
    except Exception:
        return None
    try:
        image = bitmap.to_pil()
        return _encode_image_bytes(
            image,
            max_pixels=max_pixels,
            png_optimize=png_optimize,
            jpeg_quality=jpeg_quality,
        )
    except Exception:
        return None
    finally:
        close_bitmap = getattr(bitmap, "close", None)
        if close_bitmap is not None:
            close_bitmap()


def _extract_page_images(
    page: Any,
    page_height: float,
    page_number: int,
    *,
    max_pixels: int,
    png_optimize: bool,
    jpeg_quality: int,
) -> tuple[PdfImageBlock, ...]:
    """从 pdfium 页面对象中提取内嵌图片和页面坐标。"""
    images: list[PdfImageBlock] = []
    try:
        objects = list(page.get_objects())
    except Exception:
        return ()
    for object_index, image_object in enumerate(objects):
        try:
            if not isinstance(image_object, pdfium.PdfImage):
                continue
            image_index = object_index + 1
            x0, y0, x1, y1 = (
                float(value) for value in image_object.get_bounds()
            )
            width = max(x1 - x0, 0.0)
            height = max(y1 - y0, 0.0)
            if width < 6.0 or height < 6.0:
                continue
            encoded = _image_bytes_from_object(
                image_object,
                max_pixels=max_pixels,
                png_optimize=png_optimize,
                jpeg_quality=jpeg_quality,
            )
            if encoded is None:
                continue
            data, mime_type = encoded
            images.append(
                PdfImageBlock(
                    name=f"page{page_number}_image{image_index}",
                    bbox=(x0, page_height - y1, x1, page_height - y0),
                    width=width,
                    height=height,
                    data=data,
                    mime_type=mime_type,
                    z_order=object_index,
                )
            )
        finally:
            close_object = getattr(image_object, "close", None)
            if close_object is not None:
                close_object()
    return tuple(images)


DEFAULT_EMBEDDED_IMAGE_MAX_PIXELS = 6_000_000


DEFAULT_EMBEDDED_IMAGE_JPEG_QUALITY = 85


DEFAULT_EMBEDDED_IMAGE_PNG_OPTIMIZE = False


_JPEG_PHOTO_MIN_PIXELS = 1_000_000


_JPEG_PHOTO_MIN_COLORS = 4096
