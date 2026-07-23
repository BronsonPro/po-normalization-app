"""
Reuses the existing, battle-tested parsers from the PO normalization app
(parsers/zepto.py, parsers/nykaa.py, etc.) instead of re-writing extraction
logic from scratch. Every one of those parsers exposes the same function
signature and writes the same output structure:

    convert_pdf_to_excel(input_path, output_path)
      -> writes an Excel with:
           - a header block containing rows "PO No" / "PO Date" (Field, Value)
           - an item-level table with a "Quantity" column
           - a summary block

This module calls that function unchanged (no edits to the parser files
themselves), then reads the resulting Excel back out to pull just the 3
fields the email fetcher needs: PO Number, PO Date, and PO Quantity (the
sum of the Quantity column). Full line-item detail still lives in your
normalization app - this only needs the summary.

Big Basket's parser is intentionally different: its PO arrives as Excel,
not PDF, so its convert_pdf_to_excel() takes an Excel input path instead.
"""

import os
import sys
import tempfile
import importlib
import openpyxl

# The real parsers live in the repo's top-level parsers/ folder, a sibling
# of this po_email_fetcher/ folder. Make it importable without moving/
# duplicating any of that code.
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_PARSERS_DIR = os.path.join(os.path.dirname(_THIS_DIR), "parsers")
if _PARSERS_DIR not in sys.path:
    sys.path.insert(0, _PARSERS_DIR)

# party_name -> (module name in parsers/, expected input file extension)
PARSER_MODULES = {
    "Zepto": ("zepto", "pdf"),
    "Nykaa": ("nykaa", "pdf"),
    "Blink": ("blinkit", "pdf"),
    "Scootsy": ("scootsy", "pdf"),
    "Myntra": ("myntra", "pdf"),
    "D-Mart": ("dmart", "pdf"),
    "Manash": ("manash", "pdf"),
    "Health & Glow": ("healthandglow", "pdf"),
    "Slikk": ("slikk", "pdf"),
    "Big Basket": ("bigbasket", "xlsx"),  # BB PO arrives as Excel, not PDF
    "Reliance": ("tira", "pdf"),  # Reliance POs use the same format as Tira Beauty
}

_module_cache = {}


def _get_parser_module(module_name):
    if module_name not in _module_cache:
        _module_cache[module_name] = importlib.import_module(module_name)
    return _module_cache[module_name]


def _read_summary_from_output(output_path):
    """
    Generic reader for the standard output shape every parser writes:
    header rows (Field, Value - no header), then an items table with a
    "Quantity" column header, then a summary block.
    """
    wb = openpyxl.load_workbook(output_path, data_only=True)
    ws = wb.active

    po_number = ""
    po_date = ""
    quantity_col_idx = None
    quantity_sum = 0.0
    in_items_table = False

    for row in ws.iter_rows(values_only=True):
        cells = list(row)

        if not in_items_table:
            # Header block: look for "PO No" / "PO Date" in column 1
            if len(cells) >= 2 and cells[0]:
                label = str(cells[0]).strip().lower()
                if label == "po no" and cells[1]:
                    po_number = str(cells[1]).strip()
                elif label == "po date" and cells[1]:
                    po_date = str(cells[1]).strip()

            # Look for the items table header row (contains "Quantity")
            for idx, c in enumerate(cells):
                if c and str(c).strip().lower() == "quantity":
                    quantity_col_idx = idx
                    in_items_table = True
                    break
            continue

        # We're past the items header row now - sum the Quantity column
        # until we hit a blank row (end of table / start of summary block)
        if quantity_col_idx is None or quantity_col_idx >= len(cells):
            continue
        if all(c is None or str(c).strip() == "" for c in cells):
            in_items_table = False
            quantity_col_idx = None
            continue
        val = cells[quantity_col_idx]
        try:
            quantity_sum += float(val)
        except (TypeError, ValueError):
            pass

    return po_number, po_date, quantity_sum


def get_summary_from_existing_parser(party_name: str, content_bytes: bytes) -> dict:
    """
    Returns None if this party has no existing-parser mapping (caller should
    fall back to the internal extractor). Otherwise returns:
        {"po_number": str, "po_date": str, "po_quantity": str, "error": str}
    Never raises - errors come back in "error" so they show up in the
    status log instead of crashing the batch.
    """
    if party_name not in PARSER_MODULES:
        return None

    module_name, input_ext = PARSER_MODULES[party_name]

    with tempfile.TemporaryDirectory() as tmp_dir:
        input_path = os.path.join(tmp_dir, f"input.{input_ext}")
        output_path = os.path.join(tmp_dir, "output.xlsx")

        with open(input_path, "wb") as f:
            f.write(content_bytes)

        try:
            module = _get_parser_module(module_name)
            module.convert_pdf_to_excel(input_path, output_path)
            po_number, po_date, quantity_sum = _read_summary_from_output(output_path)

            if not po_number and not po_date and not quantity_sum:
                return {
                    "po_number": "",
                    "po_date": "",
                    "po_quantity": "",
                    "error": "existing parser ran but produced no header/quantity data",
                }

            quantity_str = (
                str(int(quantity_sum)) if quantity_sum == int(quantity_sum) else str(quantity_sum)
            )
            return {
                "po_number": po_number,
                "po_date": po_date,
                "po_quantity": quantity_str,
                "error": "",
            }
        except Exception as e:
            return {
                "po_number": "",
                "po_date": "",
                "po_quantity": "",
                "error": f"existing parser ({module_name}) failed: {e}",
            }


# Default priority when a party has no known expected type (or the
# preferred type isn't among what's attached) - PDF first since that's
# the real PO document for most parties.
_DEFAULT_TYPE_PRIORITY = ["pdf", "xlsx", "xls", "csv", "zip"]


def pick_best_attachment(party_name: str, attachments: list):
    """
    An email can have more than one attachment (e.g. a Zepto email with
    both the real PO PDF and an unrelated CSV) - processing every
    attachment separately was creating a spurious extra row per email, with
    garbage data extracted from whichever file isn't actually the PO. This
    picks the ONE attachment that matches the party's known real PO format
    (from PARSER_MODULES), falling back to a sensible type priority
    (pdf > xlsx > xls > csv > zip) if the party isn't in that mapping or
    none of its attachments match the expected type.

    Returns None if attachments is empty.
    """
    if not attachments:
        return None
    if len(attachments) == 1:
        return attachments[0]

    expected_type = PARSER_MODULES.get(party_name, (None, None))[1]
    if expected_type:
        for a in attachments:
            if a["file_type"] == expected_type:
                return a

    for ft in _DEFAULT_TYPE_PRIORITY:
        for a in attachments:
            if a["file_type"] == ft:
                return a

    return attachments[0]
