"""
Per-party PDF field extraction.

Each party's PO PDF has a different layout, so each gets its own small
extractor function that returns:
    {"po_number": str, "po_date": str, "po_quantity": str}
(any field can be "" if not found - never crash, just report empty)

Only "Myntra" is calibrated against a real sample so far. All others use
GENERIC_EXTRACTOR as a best-effort placeholder until we calibrate them
against a real PDF from that party (send one and I'll tune the exact
extractor - labels/positions vary a lot between parties).
"""

import re
import pdfplumber
import openpyxl
import io


def _get_pdf_text(pdf_bytes: bytes) -> str:
    text_parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text_parts.append(page_text)
    return "\n".join(text_parts)


def _get_excel_text(excel_bytes: bytes) -> str:
    """
    Flattens all cells across all sheets into one text blob, so the same
    label-based regex extractors used for PDFs can also work on Excel POs
    (e.g. Big Basket sends Excel instead of PDF). Good enough for
    label:value style sheets; a tabular/columnar Excel PO may need a
    dedicated per-party extractor later, same as PDFs do.
    """
    text_parts = []
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if cells:
                text_parts.append(" ".join(cells))
    return "\n".join(text_parts)


def _search(pattern, text, group=1, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else ""


# ---------------------------------------------------------------------------
# MYNTRA - calibrated against real sample (PO-MYNJ-ORBI020726-2.pdf)
# ---------------------------------------------------------------------------
def extract_myntra(text: str) -> dict:
    po_number = _search(r"PO\s*#\s*:\s*([A-Za-z0-9\-]+)", text)
    po_date = _search(r"PO Approved Date\s*:\s*([\d/\-]+)", text)
    po_quantity = _search(r"Total Quantity\s*:\s*([\d,]+)", text)
    return {"po_number": po_number, "po_date": po_date, "po_quantity": po_quantity}


# ---------------------------------------------------------------------------
# GENERIC FALLBACK - best-effort for parties not yet calibrated
# Tries a handful of common label variants seen across retail PO formats.
# ---------------------------------------------------------------------------
def extract_generic(text: str) -> dict:
    po_number = _search(
        r"(?:PO\s*(?:No\.?|Number|#)|Purchase Order\s*(?:No\.?|Number)?)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)",
        text,
    )
    po_date = _search(
        r"(?:PO\s*Date|Order Date|Date)\s*[:\-]?\s*([\d/\-]{6,10})",
        text,
    )
    po_quantity = _search(
        r"(?:Total Quantity|Total Qty|Grand Total Qty|Qty Total)\s*[:\-]?\s*([\d,]+)",
        text,
    )
    return {"po_number": po_number, "po_date": po_date, "po_quantity": po_quantity}


EXTRACTORS = {
    "Myntra": extract_myntra,
    # "Reliance": extract_reliance,      # add once we see a sample
    # "D-Mart": extract_dmart,
    # "Manash": extract_manash,
    # "Handy Homes": extract_handy_homes,
    # "Zepto": extract_zepto,
    # "Big Basket": extract_big_basket,
    # "Health & Glow": extract_health_glow,
    # "Blink": extract_blink,
    # "Scootsy": extract_scootsy,
    # "Nykaa": extract_nykaa,
    # "Slikk": extract_slikk,
}


def extract_fields(party_name: str, content_bytes: bytes, file_type: str = "pdf") -> dict:
    """
    Returns {"po_number", "po_date", "po_quantity", "extractor_used"}.
    Never raises - on any failure, returns empty fields with an error note
    so it shows up in the status log instead of crashing the whole batch.

    file_type: "pdf", "xlsx", or "xls" - determines how text is pulled out
    before the party-specific (or generic) label extractor runs on it.
    """
    try:
        if file_type == "pdf":
            text = _get_pdf_text(content_bytes)
        elif file_type in ("xlsx", "xls"):
            text = _get_excel_text(content_bytes)
        else:
            raise ValueError(f"Unsupported file_type: {file_type}")
    except Exception as e:
        return {
            "po_number": "",
            "po_date": "",
            "po_quantity": "",
            "extractor_used": "none",
            "error": f"File read failed ({file_type}): {e}",
        }

    extractor = EXTRACTORS.get(party_name, extract_generic)
    extractor_name = party_name if party_name in EXTRACTORS else "generic-fallback"

    try:
        fields = extractor(text)
        fields["extractor_used"] = extractor_name
        fields["error"] = ""
        return fields
    except Exception as e:
        return {
            "po_number": "",
            "po_date": "",
            "po_quantity": "",
            "extractor_used": extractor_name,
            "error": f"Extraction failed: {e}",
        }
