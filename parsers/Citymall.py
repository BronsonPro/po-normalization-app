"""
CityMall (CMUNITY INNOVATIONS PRIVATE LIMITED) PO Parser

CityMall POs do not carry an EAN - only CityMall's own "Article Id" (e.g. CM03449498).
The parser writes the Article Id into the EAN column of the converted PO.
app.py then swaps it for the real EAN using the "Article ID" column of the CityMall master.
"""

import pdfplumber
import pandas as pd
import re


PARTY_NAME = "CMUNITY INNOVATIONS PRIVATE LIMITED"

# Item line, e.g.
# 1 CM03449498 Bronson Professional 4 In 1 Travel 96032900 500.00 226.61 12 2719.32 18.00 489.48 3208.80
ITEM_RE = re.compile(
    r"^(\d+)\s+"              # S.No
    r"(CM[A-Z0-9]+)\s+"       # Article Id
    r"(.*?)\s*"               # Article Name (first line)
    r"(\d{8})\s+"             # HSN (8 digits)
    r"([\d,]+\.\d{2})\s+"     # MRP
    r"([\d,]+\.\d{2})\s+"     # Base Cost Price
    r"([\d,]+)\s+"            # Quantity
    r"([\d,]+\.\d{2})\s+"     # Base Amount
    r"((?:[\d,]+\.\d{2}\s*)+)$"  # Tax rate/amount pairs ... Total Amount
)

# Cess values printed at the end of the continuation line: "... 0.00 0.00"
CESS_TAIL_RE = re.compile(r"\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}\s*$")


def to_float(x):
    try:
        return float(str(x).replace(",", "").strip())
    except Exception:
        return 0.0


# ------------------ HEADER ------------------

def extract_header(first_page_text, full_text):

    def find(pattern, src=first_page_text, flags=re.IGNORECASE):
        m = re.search(pattern, src, flags)
        return m.group(1).strip() if m else ""

    po_no = find(r"Purchase Order\s+(PO-\d+)")
    po_date = find(r"Purchase Order Date\s+(\d{2}-\d{2}-\d{4})")
    po_expiry = find(r"Purchase Order Expiry Date\s+(\d{2}-\d{2}-\d{4})")

    # Buyer (CityMall) GST is the first GST on the page; vendor GST comes later
    gst_no = find(r"GST\s+([0-9A-Z]{15})")

    # Delivery address spans 2 lines and is interleaved with vendor-column text
    shipping_address = ""
    lines = first_page_text.split("\n")
    for i, line in enumerate(lines):
        if line.strip().lower().startswith("delivery"):
            part1 = re.sub(r"^Delivery\s+", "", line.strip(), flags=re.I)
            part1 = re.sub(r"\s+Contact.*$", "", part1, flags=re.I)
            part2 = lines[i + 1].strip() if i + 1 < len(lines) else ""
            part2 = re.sub(r"^Address\s+", "", part2, flags=re.I)
            part2 = re.sub(r"\s+Person\s*$", "", part2, flags=re.I)
            shipping_address = f"{part1} {part2}".strip()
            break

    return {
        "Party Name": PARTY_NAME,
        "PO No": po_no,
        "PO Date": po_date,
        "PO Expiry Date": po_expiry,
        "Shipping Address": shipping_address,
        "GST #": gst_no,
    }


# ------------------ LINE ITEMS ------------------

def extract_items(all_lines):
    items = []
    current = None

    for raw in all_lines:
        line = raw.strip()
        if not line:
            continue

        # End of item table
        if re.match(r"^Total\s+[\d,]+\.\d{2}", line):
            break

        m = ITEM_RE.match(line)
        if m:
            if current:
                items.append(current)

            tail = [to_float(v) for v in m.group(9).split()]
            total = tail[-1]
            pairs = tail[:-1]             # [rate, amt] (IGST) or [rate, amt, rate, amt] (CGST+SGST)
            gst_pct = sum(pairs[0::2])    # sum of the rates

            current = {
                "Sr #": int(m.group(1)),
                "EAN": m.group(2),        # Article Id - replaced with real EAN in app.py
                "Product Name": m.group(3).strip(),
                "HSN Code": m.group(4),
                "Quantity": int(m.group(7).replace(",", "")),
                "MRP": to_float(m.group(5)),
                "Base Rate": to_float(m.group(6)),
                "GST %": round(gst_pct, 2),
                "Total": total,
            }
            continue

        # Continuation of the product name (may also carry the cess "0.00 0.00")
        if current:
            text = CESS_TAIL_RE.sub("", line).strip()
            if text and not re.fullmatch(r"[\d.,\s]+", text):
                current["Product Name"] += " " + text

    if current:
        items.append(current)

    return pd.DataFrame(items)


# ------------------ SUMMARY ------------------

def extract_summary(full_text, items_df):
    base = ""
    grand = ""

    m = re.search(r"^Total\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", full_text, re.M)
    if m:
        base = m.group(1).replace(",", "")
        grand = m.group(2).replace(",", "")

    m = re.search(r"Net Amount\s+([\d,]+\.\d{2})", full_text)
    if m:
        grand = m.group(1).replace(",", "")

    if not base:
        base = f"{(items_df['Base Rate'] * items_df['Quantity']).sum():.2f}"
    if not grand:
        grand = f"{items_df['Total'].sum():.2f}"

    tax = to_float(grand) - to_float(base)

    return {
        "Total Base Value": f"{to_float(base):.2f}",
        "Total Tax": f"{tax:.2f}",
        "Grand Total": f"{to_float(grand):.2f}",
    }


# ================== PUBLIC FUNCTION ==================

def convert_pdf_to_excel(pdf_path, output_excel_path):

    with pdfplumber.open(pdf_path) as pdf:
        page_texts = [p.extract_text() or "" for p in pdf.pages]

    first_page_text = page_texts[0]
    full_text = "\n".join(page_texts)

    header_data = extract_header(first_page_text, full_text)

    # Start reading items after the table header
    all_lines = full_text.split("\n")
    start = 0
    for i, l in enumerate(all_lines):
        if l.strip().lower().startswith("ordered items"):
            start = i + 1
            break

    items_df = extract_items(all_lines[start:])

    if items_df.empty:
        raise Exception("No line items detected in CityMall PO")

    # Renumber Sr # sequentially
    items_df["Sr #"] = range(1, len(items_df) + 1)

    summary_data = extract_summary(full_text, items_df)

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        row = 0

        header_df = pd.DataFrame({
            "Field": list(header_data.keys()),
            "Value": list(header_data.values()),
        })
        header_df.to_excel(writer, index=False, startrow=row, header=False)
        row += len(header_df) + 2

        items_df.to_excel(writer, index=False, startrow=row)
        row += len(items_df) + 2

        summary_df = pd.DataFrame({
            "Field": list(summary_data.keys()),
            "Value": list(summary_data.values()),
        })
        summary_df.to_excel(writer, index=False, startrow=row, header=False)

    return output_excel_path
