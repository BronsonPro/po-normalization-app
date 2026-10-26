"""
Flipkart Minutes PO Parser

Flipkart sends the PO as an Excel file (.xls / .xlsx), not a PDF.
The app passes the uploaded Excel straight to convert_pdf_to_excel() below.

Notes on the Flipkart format:
- "Supplier Price" is the landing price INCLUDING GST, so it is NOT used as Base Rate.
  Base Rate = Taxable Value / Quantity  (matches master "Cost Price (Excl GST)").
- Some rows have a blank "Ean". For those rows the FSN (e.g. BARHRYH8JMWYGEHV) is written
  into the EAN column, and app.py swaps it for the real EAN using the master's "FSN" column.
- Dates are 2-digit-year (23-10-26); they are written as 23-10-2026.
"""

import re
import pandas as pd


PARTY_NAME = "Flipkart"   # must match the "Party Name" used in PartyCode.xlsx


# ------------------ HELPERS ------------------

def clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def num(v):
    """'450.0 INR' -> 450.0, '18.0%' -> 18.0, '1,068' -> 1068.0"""
    s = clean(v).replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else 0.0


def fix_date(d):
    """'23-10-26' -> '23-10-2026'; anything already 4-digit is left as is."""
    d = clean(d)
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2})", d)
    if m:
        return f"{int(m.group(1)):02d}-{int(m.group(2)):02d}-20{m.group(3)}"
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2}).*", d)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return d


def value_after(raw, label):
    """Find a cell whose text equals/starts with `label` and return the next non-empty cell to its right."""
    lab = label.lower()
    for i in range(len(raw)):
        row = raw.iloc[i].tolist()
        for j, cell in enumerate(row):
            if clean(cell).lower().startswith(lab):
                for k in range(j + 1, len(row)):
                    if clean(row[k]):
                        return clean(row[k])
    return ""


# ================== MAIN ==================

def convert_pdf_to_excel(input_path, output_excel_path):

    raw = pd.read_excel(input_path, header=None)   # .xls needs the 'xlrd' package

    # -------- HEADER --------
    po_no = value_after(raw, "PO#")
    po_date = fix_date(value_after(raw, "ORDER DATE"))
    po_expiry = fix_date(value_after(raw, "PO Expiry"))
    shipping_address = value_after(raw, "SHIPPED TO ADDRESS")

    # Flipkart's GSTIN = the GSTIN printed on the SHIPPED TO row
    gst_no = ""
    for i in range(len(raw)):
        row = [clean(c) for c in raw.iloc[i].tolist()]
        if any(c.lower().startswith("shipped to address") for c in row):
            gstins = [c for c in row if re.fullmatch(r"\d{2}[A-Z0-9]{13}", c)]
            if gstins:
                gst_no = gstins[-1]
            break

    # -------- ITEM TABLE --------
    header_idx = None
    for i in range(len(raw)):
        row = [clean(c).lower() for c in raw.iloc[i].tolist()]
        if any(c.startswith("s. no") or c == "s.no" for c in row) and any("fsn" in c for c in row):
            header_idx = i
            break
    if header_idx is None:
        raise Exception("Could not find the item table (S. no. / FSN header) in Flipkart PO")

    headers = [clean(c).lower() for c in raw.iloc[header_idx].tolist()]

    def col(*keys, exclude=()):
        for idx, h in enumerate(headers):
            if all(k in h for k in keys) and not any(x in h for x in exclude):
                return idx
        return None

    c_sno = 0
    c_hsn = col("hsn")
    c_fsn = col("fsn")
    c_qty = col("quantity", exclude=("pending",))
    c_title = col("title")
    c_ean = col("ean")
    c_mrp = col("mrp")
    c_taxable = col("taxable")
    c_igst = col("igst rate")
    c_sgst = col("sgst", "rate")
    c_cgst = col("cgst rate")
    c_total = col("total amount")

    items = []
    for i in range(header_idx + 1, len(raw)):
        row = raw.iloc[i].tolist()
        sno = clean(row[c_sno])
        if not re.fullmatch(r"\d+(\.0)?", sno):
            break   # reached "Total Quantity=" / notes

        qty = int(num(row[c_qty]))
        taxable = num(row[c_taxable])
        base_rate = round(taxable / qty, 2) if qty else 0.0

        gst = num(row[c_igst]) if c_igst is not None else 0.0
        if gst == 0:
            gst = (num(row[c_sgst]) if c_sgst is not None else 0) + (num(row[c_cgst]) if c_cgst is not None else 0)

        ean = re.sub(r"\.0$", "", clean(row[c_ean])) if c_ean is not None else ""
        fsn = clean(row[c_fsn]) if c_fsn is not None else ""
        if not re.fullmatch(r"\d{8,14}", ean):
            ean = fsn   # placeholder - mapped to real EAN from master in app.py

        items.append({
            "Sr #": len(items) + 1,
            "EAN": ean,
            "Product Name": clean(row[c_title]),
            "HSN Code": re.sub(r"\.0$", "", clean(row[c_hsn])),
            "Quantity": qty,
            "MRP": num(row[c_mrp]),
            "Base Rate": base_rate,
            "GST %": round(gst, 2),
            "Total": round(num(row[c_total]), 2),
        })

    if not items:
        raise Exception("No line items detected in Flipkart PO")

    df = pd.DataFrame(items)

    # -------- SUMMARY (from the "Total=" row) --------
    total_base = total_tax = grand_total = None
    for i in range(header_idx + 1, len(raw)):
        row = [clean(c) for c in raw.iloc[i].tolist()]
        if any(c.lower().startswith("total=") for c in row):
            nums = [num(c) for c in row if "inr" in c.lower() or re.fullmatch(r"[\d.,]+", c)]
            nums = [n for n in nums if n > 0]
            if len(nums) >= 3:
                total_base, total_tax, grand_total = nums[-3], nums[-2], nums[-1]
            break

    if total_base is None:
        total_base = round((df["Base Rate"] * df["Quantity"]).sum(), 2)
        grand_total = round(df["Total"].sum(), 2)
        total_tax = grand_total - total_base

    summary = {
        "Total Base Value": f"{total_base:.2f}",
        "Total Tax": f"{total_tax:.2f}",
        "Grand Total": f"{grand_total:.2f}",
    }

    header_data = {
        "Party Name": PARTY_NAME,
        "PO No": po_no,
        "PO Date": po_date,
        "PO Expiry Date": po_expiry,
        "Shipping Address": shipping_address,
        "GST #": gst_no,
    }

    # -------- WRITE (same layout as other parsers) --------
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        row = 0
        pd.DataFrame({"Field": list(header_data.keys()), "Value": list(header_data.values())}) \
            .to_excel(writer, index=False, startrow=row, header=False)
        row += len(header_data) + 2

        df.to_excel(writer, index=False, startrow=row)
        row += len(df) + 2

        pd.DataFrame({"Field": list(summary.keys()), "Value": list(summary.values())}) \
            .to_excel(writer, index=False, startrow=row, header=False)

    return output_excel_path
