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
import csv
import zipfile

from existing_parsers_bridge import get_summary_from_existing_parser


def _get_csv_text(csv_bytes: bytes) -> str:
    """Flattens CSV rows into the same 'label value' style text used for
    Excel, so the same regex extractors can run on it."""
    text = csv_bytes.decode("utf-8", errors="replace")
    reader = csv.reader(io.StringIO(text))
    lines = []
    for row in reader:
        cells = [c.strip() for c in row if c and c.strip()]
        if cells:
            lines.append(" ".join(cells))
    return "\n".join(lines)


def _get_zip_pdf_text(zip_bytes: bytes) -> str:
    """Opens a ZIP attachment and extracts text from the first PDF found
    inside it (Health & Glow sometimes sends the PO PDF zipped)."""
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        pdf_names = [n for n in zf.namelist() if n.lower().endswith(".pdf")]
        if not pdf_names:
            raise ValueError("no PDF found inside zip attachment")
        with zf.open(pdf_names[0]) as f:
            inner_pdf_bytes = f.read()
    return _get_pdf_text(inner_pdf_bytes)


def extract_all_from_zip(party_name: str, zip_bytes: bytes) -> list:
    """
    For batch emails that bundle multiple POs into one zip (e.g. Myntra's
    "PO_PDF's_....zip" when a subject says "5 New Purchase Orders") - runs
    FULL field extraction (existing-parser bridge, same as a normal single
    PDF) on EVERY PDF found inside, not just the first. Returns a list of
    result dicts, one per inner PDF, each with an added "inner_filename" key.
    Returns an empty list if the zip has no PDFs inside (caller should
    treat that as "no extractable attachment").
    """
    results = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        pdf_names = [n for n in zf.namelist() if n.lower().endswith(".pdf")]
        for name in pdf_names:
            with zf.open(name) as f:
                inner_pdf_bytes = f.read()
            result = extract_fields(party_name, inner_pdf_bytes, "pdf")
            result["inner_filename"] = name
            results.append(result)
    return results


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
        r"(?:PO\s*Date|Order Date|Date)\s*[:\-]?\s*([\d]{1,2}[\-/][A-Za-z]{3,9}[\-/][\d]{2,4}|[\d/\-]{6,10})",
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


# ---------------------------------------------------------------------------
# SUBJECT-LINE EXTRACTION - fallback for parties whose PO reference/date is
# in the email subject itself, not (only) the attachment. Blink and Zomato's
# PO_BCPL subjects follow "...Orb international-<ref> : <date>" consistently
# across every related email for the same PO, even the ones with no
# attachment (reminders/status emails on the same thread) - so this lets us
# fill PO Number + PO Date even when there's nothing attached. Quantity still
# requires the attachment.
# ---------------------------------------------------------------------------
SUBJECT_PATTERNS = {
    "Blink": r"Orb international-(\d+)\s*:\s*(\d{4}-\d{2}-\d{2})",
    "Zomato": r"Orb international-(\d+)\s*:\s*(\d{4}-\d{2}-\d{2})",
}


def extract_subject_fields(party_name: str, subject: str) -> dict:
    """Returns {"po_number": str, "po_date": str} - both "" if no pattern/match."""
    pattern = SUBJECT_PATTERNS.get(party_name)
    if not pattern or not subject:
        return {"po_number": "", "po_date": ""}
    m = re.search(pattern, subject, re.IGNORECASE)
    if not m:
        return {"po_number": "", "po_date": ""}
    return {"po_number": m.group(1), "po_date": m.group(2)}


def extract_fields(party_name: str, content_bytes: bytes, file_type: str = "pdf") -> dict:
    """
    Returns {"po_number", "po_date", "po_quantity", "extractor_used"}.
    Never raises - on any failure, returns empty fields with an error note
    so it shows up in the status log instead of crashing the whole batch.

    file_type: "pdf", "xlsx", or "xls" - determines how text is pulled out
    before the party-specific (or generic) label extractor runs on it.

    PRIORITY: try the existing PO normalization app's calibrated parser
    first (parsers/zepto.py, parsers/nykaa.py, etc. via
    existing_parsers_bridge) - these are already tested against real
    layouts. Only fall back to the lighter internal extractor (Myntra
    regex or generic guesser) if there's no existing-parser mapping for
    this party, or it fails.
    """
    bridge_result = get_summary_from_existing_parser(party_name, content_bytes)
    if bridge_result is not None:
        if not bridge_result["error"]:
            bridge_result["extractor_used"] = f"existing-parser ({party_name})"
            return bridge_result
        # Existing parser errored - fall through to internal extractor as backup

    try:
        if file_type == "pdf":
            text = _get_pdf_text(content_bytes)
        elif file_type in ("xlsx", "xls"):
            text = _get_excel_text(content_bytes)
        elif file_type == "csv":
            text = _get_csv_text(content_bytes)
        elif file_type == "zip":
            text = _get_zip_pdf_text(content_bytes)
        else:
            raise ValueError(f"Unsupported file_type: {file_type}")
    except Exception as e:
        return {
            "po_number": "",
            "po_date": "",
            "po_quantity": "",
            "extractor_used": "none",
            "error": f"File read failed ({file_type}): {e}"
            + (f" [existing parser also failed: {bridge_result['error']}]" if bridge_result else ""),
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
