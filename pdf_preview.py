"""Рендеринг сторінок PDF у `PIL.Image` (без Tk-залежностей)."""

from __future__ import annotations

from pathlib import Path
from typing import List, Tuple

try:
    import fitz  # type: ignore  # pymupdf
    from PIL import Image  # type: ignore
except ImportError:  # pragma: no cover - залежить від оточення
    fitz = None  # type: ignore
    Image = None  # type: ignore

MAX_PREVIEW_PAGES = 10
PREVIEW_FALLBACK_WIDTH = 650
MAX_RENDER_ZOOM = 2.0


class PreviewUnavailable(RuntimeError):
    """Передає ситуацію відсутності pymupdf/Pillow у середовищі."""


def render_pdf_pages(
    pdf_path: Path,
    target_width: int,
    max_pages: int = MAX_PREVIEW_PAGES,
) -> Tuple[List["Image.Image"], int]:
    """Повертає `(images, total_pages)` для відображення.

    Зображення вже відмасштабовані до `target_width`; кількість обмежена `max_pages`.
    Може кинути `PreviewUnavailable`, якщо немає pymupdf/Pillow, або виняток I/O від PyMuPDF.
    """
    if fitz is None or Image is None:
        raise PreviewUnavailable("PyMuPDF/Pillow не встановлено.")

    if target_width < 50:
        target_width = PREVIEW_FALLBACK_WIDTH

    doc = fitz.open(str(pdf_path))
    try:
        total_pages = doc.page_count
        images: List[Image.Image] = []
        pages_to_render = min(total_pages, max_pages)
        for i in range(pages_to_render):
            page = doc.load_page(i)
            rect_w = float(page.rect.width)
            if rect_w <= 0:
                rect_w = 1.0
            zoom = min(MAX_RENDER_ZOOM, max(1.0, (2.0 * target_width) / rect_w))
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            ratio = img.width / img.height if img.height else 1.0
            new_h = max(1, int(target_width / ratio))
            img = img.resize((target_width, new_h), Image.Resampling.LANCZOS)
            images.append(img)
        return images, total_pages
    finally:
        try:
            doc.close()
        except Exception:
            pass
