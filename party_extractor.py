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
import io


def _get_pdf_text(pdf_bytes: bytes) -> str:
    text_parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text_parts.append(page_text)
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


def extract_fields(party_name: str, pdf_bytes: bytes) -> dict:
    """
    Returns {"po_number", "po_date", "po_quantity", "extractor_used"}.
    Never raises - on any failure, returns empty fields with an error note
    so it shows up in the status log instead of crashing the whole batch.
    """
    try:
        text = _get_pdf_text(pdf_bytes)
    except Exception as e:
        return {
            "po_number": "",
            "po_date": "",
            "po_quantity": "",
            "extractor_used": "none",
            "error": f"PDF read failed: {e}",
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