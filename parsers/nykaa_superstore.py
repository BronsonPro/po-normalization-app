import pdfplumber
import pandas as pd
import re
from datetime import datetime

# ------------------ HEADER EXTRACTION ------------------

def extract_po_header(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    # Nykaa Superstore POs prepend a Terms & Conditions block before the
    # real Supplier/PO Details section on page 1. That block's own text
    # (e.g. point 6 mentions "Shipping Address") pollutes the regexes below
    # if left in, so anchor to the real header section first.
    anchor = re.search(r"Supplier Details", text, re.IGNORECASE)
    if anchor:
        text = text[anchor.start():]

    def find(pattern, src=text):
        m = re.search(pattern, src, re.DOTALL | re.IGNORECASE)
        return m.group(1).strip() if m else ""

    # Superstore POs show "Pan : AAFCN5072P" (colon, not the "PAN - ..."
    # dash format the regular Nykaa PO uses), so the buyer-name regex used
    # for the regular Nykaa parser doesn't match here. The buyer is always
    # Nykaa E-Retail Limited (Superstore) for this party, so hardcode it.
    buyer_name = "Nykaa E-Retail Limited (Superstore)"

    buyer_gstin = find(r"GSTN:\s*([A-Z0-9]+)") or find(r"Shipping Address.*?GSTIN\s*:\s*([A-Z0-9]+)")

    # -------- SHIPPING ADDRESS --------
    ship_block = find(r"Shipping Address\s*(.*?)GSTIN", text)
    shipping_address = " ".join(ship_block.split()) if ship_block else ""

    po_date_str = find(r"PO Date\s+([A-Za-z]+\s\d{1,2},\s\d{4})")
    po_expiry_str = find(r"PO Expiry Date\s+([A-Za-z]+\s\d{1,2},\s\d{4})")

    # Nykaa Superstore POs often don't state an expiry date at all - default
    # to 6 months from the PO date so the exported PO (and anything reading
    # its "PO Expiry Date" field downstream) always has a real date rather
    # than a blank.
    if not po_expiry_str and po_date_str:
        try:
            from dateutil.relativedelta import relativedelta
            po_date_parsed = datetime.strptime(po_date_str, "%b %d, %Y")
            expiry_dt = po_date_parsed + relativedelta(months=6)
            po_expiry_str = expiry_dt.strftime("%b %d, %Y")
        except Exception:
            pass

    data = {
        "Party Name": buyer_name,
        "PO No": find(r"PO No\s+([A-Z0-9]+)"),
        "PO Date": po_date_str,
        "PO Expiry Date": po_expiry_str,
        "Shipping Address": shipping_address,
        "GST #": buyer_gstin,
    }

    return data


# ------------------ LINE ITEMS EXTRACTION ------------------

FINAL_COLUMNS = [
    "#", "Item Code", "EAN", "Vendor Sku", "SKU Name", "HSN", "Qty", "MRP",
    "Unit Price", "Taxable Value",
    "CGST Rate", "CGST Amt",
    "SGST Rate", "SGST Amt",
    "IGST Rate", "IGST Amt",
    "Total"
]


def extract_line_items_and_text_totals(pdf_path):
    all_rows = []
    total_text_rows = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):

            table = page.extract_table({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5,
                "snap_tolerance": 5,
                "join_tolerance": 5,
            })

            if not table:
                continue

            for row in table:
                r = [(c.strip() if c else "") for c in row]

                while len(r) < 19:
                    r.append("")

                row_text = " ".join(r).lower()

                is_total_text = (
                    "total amount" in row_text
                    or "total tax" in row_text
                    or "grand total" in row_text
                )

                out = []

                out.append(r[0])
                out.append(r[1])
                out.append(r[2])
                out.append(r[3])
                out.append(r[4])
                out.append(r[5])
                out.append(r[6])
                out.append(r[7])

                if page_idx == 0:
                    unit_price = f"{r[8] or ''} {r[9] or ''}".strip()
                    taxable = f"{r[10] or ''} {r[11] or ''}".strip()
                    tax_start = 12
                else:
                    unit_price = r[8]
                    taxable = r[9]
                    tax_start = 10

                out.append(unit_price)
                out.append(taxable)

                out.append(r[tax_start])
                out.append(r[tax_start + 1])
                out.append(r[tax_start + 2])
                out.append(r[tax_start + 3])
                out.append(r[tax_start + 4])
                out.append(r[tax_start + 5])
                out.append(r[tax_start + 6])

                if is_total_text:
                    total_text_rows.append(out)
                    continue

                if str(out[0]).strip().isdigit():
                    all_rows.append(out)

    if not all_rows:
        raise Exception("No line items detected.")

    line_df = pd.DataFrame(all_rows, columns=FINAL_COLUMNS)
    total_text_df = pd.DataFrame(total_text_rows, columns=FINAL_COLUMNS)

    return line_df.reset_index(drop=True), total_text_df.reset_index(drop=True)


# ------------------ CLEAN ------------------

def clean_and_validate_line_items(df):
    df = df.copy()

        # ✅ ADD THIS: Clean HSN Code - extract only numeric part
    def clean_hsn(x):
        if pd.isna(x):
            return ""
        s = str(x).strip()
        # Extract only the first 4-8 digits
        m = re.match(r'^(\d{4,8})', s)
        return m.group(1) if m else ""
    
    df["HSN"] = df["HSN"].apply(clean_hsn)
    
    money_cols = [
        "MRP", "Unit Price", "Taxable Value",
        "CGST Amt", "SGST Amt", "IGST Amt", "Total"
    ]

    rate_cols = ["CGST Rate", "SGST Rate", "IGST Rate"]

    def extract_number(x):
        if pd.isna(x):
            return 0.00
        s = str(x).replace(",", "")
        m = re.search(r"\d+(\.\d+)?", s)
        return round(float(m.group()) if m else 0.00, 2)

    def extract_int(x):
        if pd.isna(x):
            return 0
        m = re.search(r"\d+", str(x))
        return int(m.group()) if m else 0

    def extract_rate(x):
        if pd.isna(x):
            return 0.00
        m = re.search(r"\d+(\.\d+)?", str(x))
        return round(float(m.group()) if m else 0.00, 2)

    df["SKU Name"] = (
        df["SKU Name"]
        .astype(str)
        .str.replace("\n", " ", regex=False)
        .str.replace(r"Colour\s*:\s*.*", "", regex=True)
        .str.replace(r"Size\s*:\s*.*", "", regex=True)
        .str.replace(r"\s{2,}", " ", regex=True)
        .str.strip()
    )

    df["Qty"] = df["Qty"].apply(extract_int)

    for c in money_cols:
        df[c] = df[c].apply(extract_number)

    for c in rate_cols:
        df[c] = df[c].apply(extract_rate)

    return df


# ------------------ SUMMARY FROM TEXT ------------------

def extract_summary_from_text(total_text_df):
    if total_text_df.empty:
        return {}

    text_blob = " ".join(
        total_text_df.astype(str).fillna("").values.flatten().tolist()
    )

    def find(pattern):
        m = re.search(pattern, text_blob, re.IGNORECASE)
        return m.group(1).strip() if m else "0.00"

    def fmt2_text(x):
        try:
            val = float(str(x).replace(",", ""))
            return f"{val:.2f}"   # force 2 decimals as TEXT
        except:
            return "0.00"

    summary = {
        "Total Base Value": fmt2_text(find(r"Total Amount\(\+\)\s*([0-9,.]+)")),
        "Total Tax": fmt2_text(find(r"Total Tax\(\+\)\s*([0-9,.]+)")),
        "Grand Total": fmt2_text(find(r"Grand Total\s*([0-9,.]+)")),
    }

    return summary


# ================== PUBLIC FUNCTION FOR GUI ==================

def convert_pdf_to_excel(pdf_file_path, output_excel_path):

    header_data = extract_po_header(pdf_file_path)

    items_df, total_text_df = extract_line_items_and_text_totals(pdf_file_path)

    items_df = clean_and_validate_line_items(items_df)

    # -------- DERIVE GST % --------
    items_df["GST %"] = items_df.apply(
        lambda r: r["IGST Rate"] if r["IGST Rate"] > 0 else (r["CGST Rate"] + r["SGST Rate"]),
        axis=1
    )

    # -------- STANDARD OUTPUT --------
    final_df = pd.DataFrame({
        "EAN": items_df["EAN"],
        "Product Name": items_df["SKU Name"],
        "HSN Code": items_df["HSN"],
        "Quantity": items_df["Qty"],
        "MRP": items_df["MRP"],
        "Base Rate": items_df["Unit Price"],
        "GST %": items_df["GST %"],
        "Total": items_df["Total"],
    })

    # -------- ADD SR # --------
    final_df.insert(0, "Sr #", range(1, len(final_df) + 1))

    summary_data = extract_summary_from_text(total_text_df)

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:

        row = 0

        header_df = pd.DataFrame({
            "Field": list(header_data.keys()),
            "Value": list(header_data.values()),
        })

        header_df.to_excel(writer, index=False, startrow=row, header=False)
        row += len(header_df) + 2

        final_df.to_excel(writer, index=False, startrow=row)
        row += len(final_df) + 2

        summary_df = pd.DataFrame({
            "Field": list(summary_data.keys()),
            "Value": list(summary_data.values()),
        })

        summary_df.to_excel(writer, index=False, startrow=row, header=False)
