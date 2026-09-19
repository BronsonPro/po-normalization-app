import pdfplumber
import pandas as pd
import re
import pytesseract

# ------------------ OCR / GRID HELPERS ------------------
#
# This PDF template's text has been converted to vector outline paths, not
# real font glyphs - pdfplumber, pdfminer and PyMuPDF all extract ZERO
# characters from it, even though the content renders perfectly legibly.
# There is no way to recover the text via any text-extraction library for
# this class of PDF; the only viable approach is OCR on the rendered page
# image. To keep that reliable enough for financial data (barcodes, rates),
# every value is read from a precisely-cropped individual table cell -
# using the PDF's own vector grid lines (not OCR) to find cell boundaries -
# rather than OCR-ing a whole page or box as free-form text. Whole-page OCR
# was tested first and found to misread digits/misalign columns; per-cell
# OCR against the real grid was verified to reproduce every field correctly
# across two full sample POs.

DPI = 300
SCALE = DPI / 72


def _grid_lines(page, y_min, y_max, min_length=5):
    """Vertical and horizontal vector grid-line positions within [y_min, y_max].
    Filters out tiny fragments (occasional stray marks) by a minimum length,
    not to be confused with real text - this PDF has no text objects at all,
    only vector paths, so there's nothing else this could pick up."""
    v_lines, h_lines = set(), set()
    for l in page.lines:
        if l["top"] < y_min - 2 or l["top"] > y_max + 2:
            continue
        if abs(l["x0"] - l["x1"]) < 1:
            if abs(l["bottom"] - l["top"]) >= min_length:
                v_lines.add(round(l["x0"], 1))
        elif abs(l["top"] - l["bottom"]) < 1:
            if abs(l["x1"] - l["x0"]) >= min_length:
                h_lines.add(round(l["top"], 1))
    return sorted(v_lines), sorted(h_lines)


def _ocr_cell(im, x0, y0, x1, y1, psm=6):
    """OCR one cell, cropped a few pixels inside its border lines - including
    the border itself in the crop confuses Tesseract's segmentation for
    cells with little text (verified: this alone fixed several blank reads)."""
    left = int(x0 * SCALE) + 3
    top = int(y0 * SCALE) + 3
    right = int(x1 * SCALE) - 3
    bottom = int(y1 * SCALE) - 3
    if right <= left or bottom <= top:
        return ""
    crop = im.crop((left, top, right, bottom))
    txt = pytesseract.image_to_string(crop, config=f"--psm {psm}").strip()
    return txt.replace("\n", " ").strip()


def _ocr_row(im, cols, y0, y1, psm=6):
    return [_ocr_cell(im, cols[c], y0, cols[c + 1], y1, psm=psm) for c in range(len(cols) - 1)]


def _clean_number(s):
    if not s:
        return 0.0
    s = str(s).strip().replace(" ", "")
    # Tesseract occasionally reads the decimal point as a comma
    s = s.replace(",", ".")
    if s.count(".") > 1:
        parts = s.split(".")
        s = "".join(parts[:-1]) + "." + parts[-1]
    s = re.sub(r"[^0-9.]", "", s)
    try:
        return round(float(s), 2) if s else 0.0
    except Exception:
        return 0.0


def _clean_int(s):
    digits = re.sub(r"[^0-9]", "", str(s or ""))
    try:
        return int(digits) if digits else 0
    except Exception:
        return 0


def _clean_code(s):
    """For SKU/HSN/EAN-type cells: keep only digits (these are always
    numeric on this template)."""
    return re.sub(r"[^0-9]", "", str(s or ""))


# ------------------ HEADER EXTRACTION ------------------

def extract_po_header(pdf_path):
    party_name = "Slikk"
    po_no = ""
    po_date = ""
    po_expiry = ""
    shipping_address = ""
    gst_no = ""

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        im = page.to_image(resolution=DPI).original

        tables = page.find_tables({"vertical_strategy": "lines", "horizontal_strategy": "lines"})
        # Boxes above the product table, top to bottom: [0] PO#/Registered
        # Address, [1] PO#/Nature/PO Expiry + Category/Order Date/Company,
        # [2] Supplier Name/Brand, [3] Billed by/GSTIN/Shipped From,
        # [4] Billed To/GSTIN/Shipped To, [5] Mode of Payment/Credit Term,
        # [6] Approval Details, [7] Taxable/Tax/Total summary, [8] product
        # table. Selected by bbox position (top-to-bottom, stable across
        # POs from this same template) rather than a hand-guessed y-range.
        header_boxes = [t.bbox for t in tables if t.bbox[3] < tables[-1].bbox[1]]

        if len(header_boxes) > 1:
            x0, y0, x1, y1 = header_boxes[1]
            cols, rows = _grid_lines(page, y_min=y0, y_max=y1)
            if len(rows) >= 3 and len(cols) >= 8:
                row0 = _ocr_row(im, cols, rows[0], rows[1])
                row1 = _ocr_row(im, cols, rows[1], rows[2])
                if len(row0) >= 8:
                    po_no = row0[1].replace(" ", "")
                    po_expiry = row0[7]
                if len(row1) >= 4:
                    po_date = row1[3]

        if len(header_boxes) > 4:
            x0, y0, x1, y1 = header_boxes[4]
            cols2, rows2 = _grid_lines(page, y_min=y0, y_max=y1)
            if len(rows2) >= 2 and len(cols2) >= 8:
                row = _ocr_row(im, cols2, rows2[0], rows2[1])
                if len(row) >= 8:
                    gst_no = re.sub(r"[^0-9A-Z]", "", row[3].upper())
                    shipping_address = row[7]

    return {
        "Party Name": party_name,
        "PO No": po_no,
        "PO Date": po_date,
        "PO Expiry Date": po_expiry,
        "Shipping Address": shipping_address,
        "GST #": gst_no,
    }


# ------------------ LINE ITEMS EXTRACTION ------------------

def extract_line_items(pdf_path):
    """Every page's product table (including continuation pages) repeats
    its own header row, so each page is detected and OCR'd independently
    using its own grid."""
    all_rows = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # The product table is always the widest bordered region on the
            # page (its right edge sits further right than the header boxes
            # above it on page 1: x1 ~= 593 vs ~= 574) - on continuation
            # pages it's simply the only table found.
            tables = page.find_tables({"vertical_strategy": "lines", "horizontal_strategy": "lines"})
            prod_table = None
            for t in tables:
                if t.bbox[2] > 580:
                    prod_table = t
                    break
            if prod_table is None:
                continue

            x0, y0, x1, y1 = prod_table.bbox
            cols, row_y = _grid_lines(page, y_min=y0, y_max=y1, min_length=5)
            if len(cols) < 11 or len(row_y) < 2:
                continue

            # A row whose bottom border is cut off by the page boundary (not
            # a genuine end of table) still has its column-divider lines
            # drawn past the last detected horizontal divider. Require most
            # columns to agree closely on where that extension ends, and
            # that it's not far below the last row (roughly one row's worth
            # at most) - otherwise a coincidentally shared left-edge x with
            # an unrelated box further down the page (e.g. Terms &
            # Conditions) can get mistaken for a continuation of this table.
            col_extend_bottoms = []
            for x in cols:
                matching = [
                    l["bottom"] for l in page.lines
                    if abs(l["x0"] - l["x1"]) < 1
                    and abs(l["x0"] - x) < 0.5
                    and l["top"] <= row_y[-1] + 3
                    and l["bottom"] > row_y[-1] + 3
                ]
                if matching:
                    col_extend_bottoms.append(min(matching))
            if len(col_extend_bottoms) >= 8:
                lo, hi = min(col_extend_bottoms), max(col_extend_bottoms)
                if hi - lo < 5 and lo - row_y[-1] < 40:
                    row_y = row_y + [lo]

            im = page.to_image(resolution=DPI).original
            headers = ["S.No", "Product Name", "SKU", "HSN", "MRP", "SP", "Qty",
                       "Purchase Price", "Item Value", "Total Tax", "Total Amount"]

            for r in range(len(row_y) - 1):
                ry0, ry1 = row_y[r], row_y[r + 1]
                if ry1 - ry0 < 8:
                    continue
                vals = _ocr_row(im, cols, ry0, ry1)
                if len(vals) < 11:
                    continue
                row = dict(zip(headers, vals))
                if r == 0:
                    continue  # header row
                all_rows.append(row)

    if not all_rows:
        raise Exception("No line items found in Slikk PO")

    # A row whose product-name cell wraps too tall to fit can have its
    # description pushed onto the FIRST row of the following page, with
    # every other cell in that fragment left blank (verified: this happens
    # exactly at a page break). Detect and merge such fragments back into
    # the row they belong to.
    merged_rows = []
    i = 0
    while i < len(all_rows):
        row = all_rows[i]
        has_data = bool(_clean_code(row["SKU"]))
        has_name = bool(row["Product Name"].strip())
        if has_data and not has_name and i + 1 < len(all_rows):
            nxt = all_rows[i + 1]
            nxt_has_data = bool(_clean_code(nxt["SKU"]))
            nxt_has_name = bool(nxt["Product Name"].strip())
            if nxt_has_name and not nxt_has_data:
                row = dict(row)
                row["Product Name"] = nxt["Product Name"]
                merged_rows.append(row)
                i += 2
                continue
        merged_rows.append(row)
        i += 1

    items = []
    for idx, row in enumerate(merged_rows, start=1):
        qty = _clean_int(row["Qty"])
        ean = _clean_code(row["SKU"])
        if qty <= 0 or not ean:
            continue
        item_value = _clean_number(row["Item Value"])
        total_tax = _clean_number(row["Total Tax"])
        gst_pct = round((total_tax / item_value) * 100, 2) if item_value > 0 else 0.0
        items.append({
            "Sr #": idx,
            "EAN": ean,
            "Product Name": row["Product Name"].strip(),
            "HSN Code": _clean_code(row["HSN"]),
            "Quantity": qty,
            "MRP": _clean_number(row["MRP"]),
            # "SP" (Selling Price) is the master-comparable base rate on this
            # template - "Purchase Price" is SP net of the vendor discount,
            # analogous to Base Rate used elsewhere in this app.
            "Base Rate": _clean_number(row["Purchase Price"]),
            "GST %": gst_pct,
            "Total": _clean_number(row["Total Amount"]),
        })

    return pd.DataFrame(items)


# ------------------ SUMMARY EXTRACTION ------------------

def extract_summary(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        im = page.to_image(resolution=DPI).original
        tables = page.find_tables({"vertical_strategy": "lines", "horizontal_strategy": "lines"})
        header_boxes = [t.bbox for t in tables if t.bbox[3] < tables[-1].bbox[1]]

        row = []
        if len(header_boxes) > 7:
            x0, y0, x1, y1 = header_boxes[7]
            cols, rows = _grid_lines(page, y_min=y0, y_max=y1)
            # rows[0]-rows[1] is the label row ("Taxable Value" / "Tax
            # Amount" / "Total Amount"); rows[1]-rows[2] holds the actual
            # numeric values.
            if len(rows) >= 3 and len(cols) >= 3:
                row = _ocr_row(im, cols, rows[1], rows[2])

    total_base = _clean_number(row[0]) if len(row) > 0 else 0.0
    total_tax = _clean_number(row[1]) if len(row) > 1 else 0.0
    grand_total = _clean_number(row[2]) if len(row) > 2 else 0.0

    return {
        "Total Base Value": f"{total_base:.2f}",
        "Total Tax": f"{total_tax:.2f}",
        "Grand Total": f"{grand_total:.2f}",
    }


# ================== PUBLIC FUNCTION ==================

def convert_pdf_to_excel(pdf_path, output_excel_path):
    header_data = extract_po_header(pdf_path)
    products = extract_line_items(pdf_path)
    summary_data = extract_summary(pdf_path)

    if products.empty:
        raise Exception("No line items found in Slikk PO")

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        row_offset = 0

        header_df = pd.DataFrame({
            "Field": list(header_data.keys()),
            "Value": list(header_data.values()),
        })
        header_df.to_excel(writer, index=False, startrow=row_offset, header=False)
        row_offset += len(header_df) + 2

        products.to_excel(writer, index=False, startrow=row_offset)
        row_offset += len(products) + 2

        summary_df = pd.DataFrame({
            "Field": list(summary_data.keys()),
            "Value": list(summary_data.values()),
        })
        summary_df.to_excel(writer, index=False, startrow=row_offset, header=False)
