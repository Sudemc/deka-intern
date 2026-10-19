import io
import pytest
from reportlab.pdfgen import canvas
from bom_enricher.pdf_renderer import select_and_render_pages, select_target_pages


def create_sample_pdf(pages_text: list[str]) -> bytes:
    """Helper to generate a multi-page PDF in memory."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    for text in pages_text:
        c.drawString(100, 750, text)
        c.showPage()
    c.save()
    return buf.getvalue()


def test_select_target_pages_default_two_pages():
    # For a 5-page PDF with no match, returns [0, 1]
    pages = select_target_pages(total_pages=5, page_texts=["", "", "", "", ""], target_pn=None, max_pages=2)
    assert pages == [0, 1]


def test_select_target_pages_matches_target_pn():
    # Page 3 mentions target_pn
    texts = ["cover", "specs", "tables", "DG8HS-1.2-12P drawings", "compliance"]
    pages = select_target_pages(total_pages=5, page_texts=texts, target_pn="DG8HS-1.2-12P", max_pages=2)
    assert pages == [0, 3]


def test_select_target_pages_series_token_match():
    # Target PN has packaging suffix, but page has base series token
    texts = ["cover", "general table", "DG8HS series dimensions", "notes"]
    pages = select_target_pages(total_pages=4, page_texts=texts, target_pn="DG8HS-1.2-12P/TR", max_pages=2)
    assert pages == [0, 2]


def test_select_target_pages_single_page_pdf():
    pages = select_target_pages(total_pages=1, page_texts=["single page"], target_pn="SOMEPN", max_pages=2)
    assert pages == [0]


def test_select_target_pages_empty_or_zero():
    assert select_target_pages(total_pages=0, page_texts=[], target_pn="PN", max_pages=2) == []
    assert select_target_pages(total_pages=5, page_texts=[""] * 5, target_pn="PN", max_pages=0) == []


def test_render_pdf_page_to_base64_with_pypdfium2():
    pdf_bytes = create_sample_pdf(["Test CAD Drawing Page"])

    rendered = select_and_render_pages(pdf_bytes, target_pn="Test", max_pages=2)
    assert len(rendered) == 1
    assert rendered[0]["page_number"] == 1
    assert isinstance(rendered[0]["base64_png"], str)
    assert len(rendered[0]["base64_png"]) > 100
    assert rendered[0]["width"] > 0
    assert rendered[0]["height"] > 0


def test_render_multi_page_pdf_with_target_pn():
    pages_text = [
        "Catalog Overview",
        "Ordering Information Table",
        "DG8HS-1.2-12P Mechanical Dimensions and Pinout",
        "Application Notes",
    ]
    pdf_bytes = create_sample_pdf(pages_text)

    rendered = select_and_render_pages(pdf_bytes, target_pn="DG8HS-1.2-12P", max_pages=2)
    assert len(rendered) == 2
    # page_number is 1-indexed
    assert rendered[0]["page_number"] == 1
    assert rendered[1]["page_number"] == 3
    for item in rendered:
        assert isinstance(item["base64_png"], str)
        assert len(item["base64_png"]) > 100
        assert item["width"] > 0
        assert item["height"] > 0


def test_render_scale_affects_dimensions():
    pdf_bytes = create_sample_pdf(["Single page"])
    rendered_scale1 = select_and_render_pages(pdf_bytes, max_pages=1, scale=1.0)
    rendered_scale2 = select_and_render_pages(pdf_bytes, max_pages=1, scale=2.0)

    assert rendered_scale2[0]["width"] > rendered_scale1[0]["width"]
    assert rendered_scale2[0]["height"] > rendered_scale1[0]["height"]


def test_render_invalid_pdf_bytes_returns_empty():
    assert select_and_render_pages(b"") == []
    assert select_and_render_pages(b"not a valid pdf header") == []
    assert select_and_render_pages(b"", max_pages=0) == []


def test_select_target_pages_max_pages_three():
    texts = ["p0", "p1", "p2", "p3", "p4-target"]
    pages = select_target_pages(total_pages=5, page_texts=texts, target_pn="target", max_pages=3)
    assert pages == [0, 4, 1]


def test_select_target_pages_extended_suffix_and_short_token_boundary():
    # Suffix with #pbf or +t stripped
    texts_suffix = ["cover", "notes", "LT1086 regulator pinout"]
    pages_suffix = select_target_pages(total_pages=3, page_texts=texts_suffix, target_pn="LT1086#PBF", max_pages=2)
    assert pages_suffix == [0, 2]

    # Short token (< 5 chars) should NOT match as substring of longer word (e.g. "CON" inside "ICONIC")
    texts_boundary = ["cover", "This is an ICONIC design", "Pinout for CON connector"]
    pages_boundary = select_target_pages(total_pages=3, page_texts=texts_boundary, target_pn="CON-100", max_pages=2)
    assert pages_boundary == [0, 2]


