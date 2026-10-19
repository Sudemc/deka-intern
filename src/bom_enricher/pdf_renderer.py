"""
Smart PDF Page Selection & High-Speed Rendering using pypdfium2.
Renders target pages to base64 PNG images for Multimodal Vision LLM processing.
"""
from __future__ import annotations

import base64
import io
import logging
import re
from typing import Any

import pypdfium2 as pdfium

logger = logging.getLogger(__name__)


def select_target_pages(
    total_pages: int,
    page_texts: list[str],
    target_pn: str | None = None,
    max_pages: int = 2,
) -> list[int]:
    """Select up to max_pages 0-indexed page indices for visual inspection.

    Strategy:
    - Page 0 (cover/overview) is always included.
    - If target_pn is provided, find first page index i > 0 where clean PN
      or series token appears in page_texts[i].
    - If found, add i.
    - If not found and total_pages > 1, add page 1.
    - Cap result list at max_pages.
    """
    if total_pages <= 0 or max_pages <= 0:
        return []

    selected: list[int] = [0]
    if total_pages == 1 or max_pages == 1:
        return selected

    found_page: int | None = None
    if target_pn and target_pn.strip():
        pn_clean = target_pn.strip().lower()
        search_tokens: list[str] = [pn_clean]

        # Suffix removal (parentheses / brackets e.g. packaging)
        no_suffix = re.sub(r"[\(\[].*?[\)\]]", "", target_pn).strip().lower()
        if no_suffix and no_suffix not in search_tokens:
            search_tokens.append(no_suffix)

        # Common packaging / tape-reel suffixes
        suffix_clean = re.sub(r"([-_/](?:tr|reel|tape|t|nd|pbf))+$", "", pn_clean, flags=re.IGNORECASE).strip()
        if suffix_clean and suffix_clean not in search_tokens:
            search_tokens.append(suffix_clean)

        # Prefix before last dash/hyphen
        if "-" in pn_clean:
            prefix = pn_clean.rsplit("-", 1)[0]
            if len(prefix) >= 3 and prefix not in search_tokens:
                search_tokens.append(prefix)

        # First alphanumeric token before delimiter (e.g. series prefix like DG8HS)
        delim_parts = [p for p in re.split(r"[-_/\s]+", pn_clean) if p]
        if delim_parts and len(delim_parts[0]) >= 3 and delim_parts[0] not in search_tokens:
            search_tokens.append(delim_parts[0])

        for token in search_tokens:
            for idx in range(1, total_pages):
                text = page_texts[idx].lower() if idx < len(page_texts) else ""
                if token in text:
                    found_page = idx
                    break
            if found_page is not None:
                break

    if found_page is not None:
        selected.append(found_page)
    elif total_pages > 1:
        selected.append(1)

    # If max_pages allows more and we have remaining pages, fill sequentially
    if len(selected) < max_pages:
        for idx in range(1, total_pages):
            if idx not in selected:
                selected.append(idx)
            if len(selected) >= max_pages:
                break

    return selected[:max_pages]


def select_and_render_pages(
    pdf_bytes: bytes,
    target_pn: str | None = None,
    max_pages: int = 2,
    scale: float = 2.0,
) -> list[dict[str, Any]]:
    """Render selected pages of a PDF to base64 PNG images.

    Args:
        pdf_bytes: Raw PDF content as bytes.
        target_pn: Target manufacturer part number to look for in pages.
        max_pages: Maximum number of pages to render (default: 2).
        scale: Resolution scale factor (default: 2.0 for clear OCR/vision).

    Returns:
        List of dicts with keys:
            - page_number: int (1-indexed)
            - base64_png: str (base64-encoded PNG image data)
            - width: int (pixel width of rendered image)
            - height: int (pixel height of rendered image)
    """
    if not pdf_bytes or max_pages <= 0:
        return []

    try:
        doc = pdfium.PdfDocument(pdf_bytes)
    except Exception as exc:
        logger.warning("Failed to open PDF with pypdfium2: %s", exc)
        return []

    rendered_pages: list[dict[str, Any]] = []
    try:
        total_pages = len(doc)
        if total_pages == 0:
            return []

        # Lightweight text extraction for smart page selection
        page_texts: list[str] = []
        for i in range(total_pages):
            try:
                page = doc[i]
                textpage = page.get_textpage()
                page_texts.append(textpage.get_text_range() or "")
            except Exception as e:
                logger.debug("Failed to extract text from page %d: %s", i, e)
                page_texts.append("")

        target_indices = select_target_pages(
            total_pages=total_pages,
            page_texts=page_texts,
            target_pn=target_pn,
            max_pages=max_pages,
        )

        for page_idx in target_indices:
            try:
                page = doc[page_idx]
                pil_img = page.render(scale=scale).to_pil()
                img_buf = io.BytesIO()
                pil_img.save(img_buf, format="PNG")
                b64_str = base64.b64encode(img_buf.getvalue()).decode("utf-8")

                rendered_pages.append({
                    "page_number": page_idx + 1,  # 1-indexed
                    "base64_png": b64_str,
                    "width": pil_img.width,
                    "height": pil_img.height,
                })
            except Exception as e:
                logger.warning("Failed to render page %d: %s", page_idx, e)
    finally:
        try:
            doc.close()
        except Exception:
            pass

    return rendered_pages
