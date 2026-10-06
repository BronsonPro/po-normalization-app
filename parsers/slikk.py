import pdfplumber
import pandas as pd
import re
import pytesseract
from PIL import Image
from concurrent.futures import ThreadPoolExecutor

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
#
# PERFORMANCE NOTE: each pytesseract call spawns a separate `tesseract` CLI
# subprocess. Measured startup+recognition cost is ~0.17s PER CELL almost
# entirely from subprocess spin-up, not the tiny crop it's reading - so a
# 52-row PO (11 cells/row) made ~600 sequential subprocess calls, which is
# where nearly all the runtime went. Since each call is an independent
# subprocess wait (not CPU-bound in this process), running many of them
# concurrently via a thread pool overlaps that spin-up cost instead of
# paying it serially - measured ~2.2x speedup even on a 2-core machine.
# `page.find_tables()` and `page.to_image()` are also each call-once,
# ~0.15-0.6s - both were previously being re-run 2-3x per PDF (once each
# from extract_po_header/extract_line_items/extract_summary, which each
# used to open and re-render the PDF independently), so this version opens
# the PDF once and caches each page's rendered image and detected tables
# for reuse across all three extraction steps.

DPI = 300
SCALE = DPI / 72
MAX_OCR_WORKERS = 8


def _get_page_image(page, img_cache):
    if page.page_number not in img_cache:
        img_cache[page.page_number] = page.to_image(resolution=DPI).original
    return img_cache[page.page_number]


def _get_page_tables(page, tables_cache):
    if page.page_number not in tables_cache:
        tables_cache[page.page_number] = page.find_tables(
            {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
        )
    return tables_cache[page.page_number]


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


def _ocr_row_batched(im, cols, y0, y1, psm=6, gap=24):
    """OCR a whole row in ONE tesseract call instead of one call per cell
    (11 for a product row), while still giving Tesseract each cell as its
    own visually isolated block - exactly like the per-cell version, just
    laid out in a single composite image instead of dispatched as 11
    separate subprocess calls.

    Why this exists: the thread-pool parallelism in _ocr_row only helps
    when multiple CPU cores are actually free to overlap on - Streamlit
    Community Cloud containers are commonly limited to a single vCPU,
    where running several "concurrent" tesseract subprocesses just
    time-slices one core and barely beats doing them one at a time. The
    dominant cost is per-call subprocess spin-up (~0.17s), which scales
    with the NUMBER of calls, not the core count - cutting 11 calls/row
    down to 1 removes that cost outright regardless of cores available.

    An earlier version of this function cropped the entire row width
    (~2400px, an extreme ~26:1 aspect ratio) and OCR'd it as one wide
    line. That measurably corrupted numeric values on some rows even
    though the crop was visually clean (reproduced independently via
    plain pytesseract.image_to_string on the saved crop) - Tesseract's
    own layout analysis mis-segmenting that unusual shape, not a bug in
    the column-bucketing here. Stacking each cell into its own generously
    padded vertical slot instead keeps every cell the same shape/size
    pytesseract already handles reliably (a small isolated block of
    text), so this is the per-cell approach's accuracy in a single call,
    not a differently-shaped OCR problem.
    """
    from pytesseract import Output

    n_cols = len(cols) - 1
    crops = []
    for c in range(n_cols):
        left = int(cols[c] * SCALE) + 3
        top = int(y0 * SCALE) + 3
        right = int(cols[c + 1] * SCALE) - 3
        bottom = int(y1 * SCALE) - 3
        crops.append(im.crop((left, top, right, bottom)) if right > left and bottom > top else None)

    real_crops = [c for c in crops if c is not None]
    if not real_crops:
        return ["" for _ in range(n_cols)]
    max_w = max(c.width for c in real_crops)
    slot_h = max(c.height for c in real_crops)
    stride = slot_h + gap

    canvas = Image.new("RGB", (max_w, stride * n_cols), "white")
    for c, crop in enumerate(crops):
        if crop is not None:
            canvas.paste(crop, (0, c * stride))

    data = pytesseract.image_to_data(canvas, config=f"--psm {psm}", output_type=Output.DICT)

    buckets = [[] for _ in range(n_cols)]
    for i, word in enumerate(data["text"]):
        word = word.strip()
        if not word:
            continue
        # A faint remnant of a cell's own border line, still visible
        # despite the inward margin, occasionally gets OCR'd as a lone "|"
        # and would otherwise leak into whichever slot it landed in - it's
        # never a real character in this template's codes, rates or
        # names, so drop it.
        if re.fullmatch(r"[|]+", word):
            continue
        slot = max(0, min(n_cols - 1, data["top"][i] // stride))
        # Within a slot, a wrapped multi-line cell (e.g. a 2-line Product
        # Name) needs its lines read top-to-bottom, then left-to-right on
        # each. Raw pixel `top` isn't a reliable primary sort key for
        # that: same-line words can differ by a few px of `top` purely
        # from font-metric noise (a capital letter vs. a lowercase one),
        # which is enough to reorder same-line words if sorted on `top`
        # before `left`. Tesseract's own `line_num` (stable per detected
        # text line within this isolated, padded slot) is the reliable
        # signal for which physical line a word is on; `left` then orders
        # words within that same line.
        buckets[slot].append((data["line_num"][i], data["left"][i], word))

    result = []
    for b in buckets:
        b.sort(key=lambda t: (t[0], t[1]))
        result.append(" ".join(w for _, _, w in b))
    return result


def _ocr_row(im, cols, y0, y1, executor=None, psm=6, batch=False):
    """OCR every cell of one row.
    - batch=True: one tesseract call for the whole row (see
      _ocr_row_batched) - the big win on single-core deployments.
    - executor given (and not batching): the 11 per-cell calls are
      dispatched concurrently - helps only when multiple cores are free.
    - neither: original sequential per-cell calls.
    Order is preserved in all three, so downstream parsing is unaffected."""
    if batch:
        return _ocr_row_batched(im, cols, y0, y1, psm=psm)
    cells = [(cols[c], y0, cols[c + 1], y1) for c in range(len(cols) - 1)]
    if executor is not None and len(cells) > 1:
        return list(executor.map(lambda c: _ocr_cell(im, c[0], c[1], c[2], c[3], psm=psm), cells))
    return [_ocr_cell(im, *c, psm=psm) for c in cells]


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



def _clean_date(s):
    """OCR can misread characters in dates (e.g. '01 Oct 2026' read as
    'O01 Oct 2026', or a 0 read as O / 1 as l). Pull out day / month name /
    year, fix look-alike characters in the numeric parts, and return a
    normalized 'DD Mon YYYY'. If it can't be understood, return the text
    unchanged so nothing is silently invented."""
    raw = str(s or "").strip()
    m = re.search(r"([0-9OoIl|]{1,3})\s*([A-Za-z]{3,9})\.?,?\s*([0-9OoIl|]{4})", raw)
    if not m:
        return raw
    fix = lambda t: re.sub(r"[Oo]", "0", re.sub(r"[Il|]", "1", t))
    try:
        day = int(fix(m.group(1)))
        year = fix(m.group(3))
        mon = m.group(2)[:3].title()
        if not (1 <= day <= 31):
            day = int(str(day)[-2:])
        return f"{day:02d} {mon} {year}"
    except Exception:
        return raw

# ------------------ HEADER EXTRACTION ------------------

def extract_po_header(pdf_path, _pdf=None, _executor=None, _img_cache=None, _tables_cache=None):
    party_name = "Slikk"
    po_no = ""
    po_date = ""
    po_expiry = ""
    shipping_address = ""
    gst_no = ""

    def _run(pdf, executor, img_cache, tables_cache):
        nonlocal po_no, po_date, po_expiry, shipping_address, gst_no
        page = pdf.pages[0]
        im = _get_page_image(page, img_cache)

        tables = _get_page_tables(page, tables_cache)
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
                row0 = _ocr_row(im, cols, rows[0], rows[1], executor=executor)
                row1 = _ocr_row(im, cols, rows[1], rows[2], executor=executor)
                if len(row0) >= 8:
                    po_no = row0[1].replace(" ", "")
                    po_expiry = _clean_date(row0[7])
                if len(row1) >= 4:
                    po_date = _clean_date(row1[3])

        if len(header_boxes) > 4:
            x0, y0, x1, y1 = header_boxes[4]
            cols2, rows2 = _grid_lines(page, y_min=y0, y_max=y1)
            if len(rows2) >= 2 and len(cols2) >= 8:
                row = _ocr_row(im, cols2, rows2[0], rows2[1], executor=executor)
                if len(row) >= 8:
                    gst_no = re.sub(r"[^0-9A-Z]", "", row[3].upper())
                    shipping_address = row[7]

    if _pdf is not None:
        _run(_pdf, _executor, _img_cache, _tables_cache)
    else:
        with pdfplumber.open(pdf_path) as pdf, ThreadPoolExecutor(max_workers=MAX_OCR_WORKERS) as executor:
            _run(pdf, executor, {}, {})

    return {
        "Party Name": party_name,
        "PO No": po_no,
        "PO Date": po_date,
        "PO Expiry Date": po_expiry,
        "Shipping Address": shipping_address,
        "GST #": gst_no,
    }


# ------------------ LINE ITEMS EXTRACTION ------------------

def extract_line_items(pdf_path, _pdf=None, _executor=None, _img_cache=None, _tables_cache=None):
    """Every page's product table (including continuation pages) repeats
    its own header row, so each page is detected and OCR'd independently
    using its own grid."""
    all_rows = []

    def _run(pdf, executor, img_cache, tables_cache):
        for page in pdf.pages:
            # The product table is always the widest bordered region on the
            # page (its right edge sits further right than the header boxes
            # above it on page 1: x1 ~= 593 vs ~= 574) - on continuation
            # pages it's simply the only table found.
            tables = _get_page_tables(page, tables_cache)
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

            im = _get_page_image(page, img_cache)
            headers = ["S.No", "Product Name", "SKU", "HSN", "MRP", "SP", "Qty",
                       "Purchase Price", "Item Value", "Total Tax", "Total Amount"]

            for r in range(len(row_y) - 1):
                ry0, ry1 = row_y[r], row_y[r + 1]
                if ry1 - ry0 < 8:
                    continue
                # batch=True: 1 tesseract call for this row instead of 11 -
                # see _ocr_row_batched for why this (not the thread pool
                # above) is the real fix for a single-core deployment.
                vals = _ocr_row(im, cols, ry0, ry1, batch=True)
                if len(vals) < 11:
                    continue
                row = dict(zip(headers, vals))
                if r == 0:
                    continue  # header row
                all_rows.append(row)

    if _pdf is not None:
        _run(_pdf, _executor, _img_cache, _tables_cache)
    else:
        with pdfplumber.open(pdf_path) as pdf, ThreadPoolExecutor(max_workers=MAX_OCR_WORKERS) as executor:
            _run(pdf, executor, {}, {})

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

def extract_summary(pdf_path, _pdf=None, _executor=None, _img_cache=None, _tables_cache=None):
    row = []

    def _run(pdf, executor, img_cache, tables_cache):
        nonlocal row
        page = pdf.pages[0]
        im = _get_page_image(page, img_cache)
        tables = _get_page_tables(page, tables_cache)
        header_boxes = [t.bbox for t in tables if t.bbox[3] < tables[-1].bbox[1]]

        if len(header_boxes) > 7:
            x0, y0, x1, y1 = header_boxes[7]
            cols, rows = _grid_lines(page, y_min=y0, y_max=y1)
            # rows[0]-rows[1] is the label row ("Taxable Value" / "Tax
            # Amount" / "Total Amount"); rows[1]-rows[2] holds the actual
            # numeric values.
            if len(rows) >= 3 and len(cols) >= 3:
                row = _ocr_row(im, cols, rows[1], rows[2], executor=executor)

    if _pdf is not None:
        _run(_pdf, _executor, _img_cache, _tables_cache)
    else:
        with pdfplumber.open(pdf_path) as pdf, ThreadPoolExecutor(max_workers=MAX_OCR_WORKERS) as executor:
            _run(pdf, executor, {}, {})

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
    # Open the PDF once and share one thread pool + per-page image/table
    # caches across all three extraction passes, instead of each one
    # re-opening the file and re-rendering/re-detecting page 0 from
    # scratch (previously page 0 alone was rendered and table-detected up
    # to 3 times). See the PERFORMANCE NOTE above _get_page_image.
    with pdfplumber.open(pdf_path) as pdf, ThreadPoolExecutor(max_workers=MAX_OCR_WORKERS) as executor:
        img_cache = {}
        tables_cache = {}
        header_data = extract_po_header(pdf_path, _pdf=pdf, _executor=executor,
                                         _img_cache=img_cache, _tables_cache=tables_cache)
        products = extract_line_items(pdf_path, _pdf=pdf, _executor=executor,
                                       _img_cache=img_cache, _tables_cache=tables_cache)
        summary_data = extract_summary(pdf_path, _pdf=pdf, _executor=executor,
                                        _img_cache=img_cache, _tables_cache=tables_cache)

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
