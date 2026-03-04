import pdfplumber
import pandas as pd
import re


# ------------------ HELPERS ------------------

def clean(x):
    return x.strip() if isinstance(x, str) else ""


def num(x):
    if not x:
        return ""
    m = re.search(r"\d+(\.\d+)?", str(x).replace(",", ""))
    return m.group() if m else ""


def fmt2(x):
    try:
        return f"{float(str(x).replace(',', '')):.2f}"
    except:
        return x


# ================== MAIN ==================

def convert_pdf_to_excel(pdf_path, output_excel_path):

    items = []

    po_no = ""
    po_date = ""
    po_expiry = ""
    gst_no = ""
    shipping_address = ""

    total_base_value = ""
    total_tax = ""
    grand_total = ""

    with pdfplumber.open(pdf_path) as pdf:

        # =====================================================
        # FULL TEXT (HEADER + SUMMARY)
        # =====================================================
        full_text = "\n".join([p.extract_text() or "" for p in pdf.pages])

        # -------- PO DETAILS --------
        m = re.search(r"PO\s*No\s*:\s*([A-Z0-9]+)", full_text)
        if m:
            po_no = m.group(1)

        m = re.search(r"PO\s*Date\s*:\s*([\d\-]+)", full_text)
        if m:
            po_date = m.group(1)

        m = re.search(r"PO\s*Expiry\s*Date\s*:\s*([\d\-]+)", full_text)
        if m:
            po_expiry = m.group(1)

        # -------- SHIPPING ADDRESS + GSTIN --------
        shipping_address = ""
        gst_no = ""

        lines = [l.strip() for l in full_text.split("\n") if l.strip()]

        for i, line in enumerate(lines):
            if "shipping" in line.lower() and "address" in line.lower():

                addr_lines = []

                for j in range(i + 1, min(i + 25, len(lines))):

                    if "gstin" in lines[j].lower():
                        gst_no = lines[j].split(":")[-1].strip()
                        break

                    if lines[j].lower().startswith("address"):
                        continue

                    addr_lines.append(lines[j])

                # remove duplicates
                cleaned = []
                for a in addr_lines:
                    if not cleaned or cleaned[-1] != a:
                        cleaned.append(a)

                shipping_address = "\n".join(cleaned)

                # remove duplicate block if repeated
                parts = shipping_address.split("\n")
                half = len(parts) // 2
                if parts[:half] == parts[half:]:
                    shipping_address = "\n".join(parts[:half])

                break

        # -------- SUMMARY --------
        m = re.search(r"Total\s+Amount\s*\(INR\)\s*([\d,]+\.\d{2})", full_text, re.I)
        if m:
            total_base_value = m.group(1)

        m = re.search(r"Total\s+Tax\s*\(INR\)\s*([\d,]+\.\d{2})", full_text, re.I)
        if m:
            total_tax = m.group(1)

        m = re.search(r"Grand\s+Total\s*\(INR\)\s*([\d,]+\.\d{2})", full_text, re.I)
        if m:
            grand_total = m.group(1)

        # =====================================================
        # LINE ITEMS (COLUMN POSITION BASED)
        # =====================================================
        header = None

        for page in pdf.pages:

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
                r = [clean(c) for c in row]
                joined = " ".join(r).lower()

                if "material code" in joined and "ean" in joined and "quantity" in joined:
                    header = r
                    continue

                if not header:
                    continue

                if not r or not r[0].isdigit():
                    continue

                def col(name):
                    for i, h in enumerate(header):
                        if name in h.lower():
                            return r[i]
                    return ""

                cgst = num(col("cgst"))
                sgst = num(col("sgst"))
                igst = num(col("igst"))

                if igst and float(igst) > 0:
                    gst_pct = igst
                else:
                    try:
                        gst_pct = str(round(float(cgst or 0) + float(sgst or 0), 2))
                    except:
                        gst_pct = ""

                items.append({
                    "EAN": col("ean"),
                    "Product Name": col("item description"),
                    "HSN Code": col("hsn"),
                    "Quantity": num(col("quantity")),
                    "MRP": num(col("mrp")),
                    "Base Rate": num(col("unit base")),
                    "GST %": gst_pct,
                    "Total": num(r[-1]),
                })

    if not items:
        raise Exception("No line items detected in Zepto PDF")

    df = pd.DataFrame(items)

    # -------- ADD SR # COLUMN --------
    df.insert(0, "Sr #", range(1, len(df) + 1))

    # =====================================================
    # WRITE FINAL OUTPUT FORMAT
    # =====================================================
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:

        row = 0

        header_df = pd.DataFrame({
            "Field": [
                "Party Name",
                "PO No",
                "PO Date",
                "PO Expiry Date",
                "Shipping Address",
                "GST #",
            ],
            "Value": [
                "Zepto",
                po_no,
                po_date,
                po_expiry,
                shipping_address,
                gst_no,
            ],
        })

        header_df.to_excel(writer, index=False, startrow=row, header=False)
        row += len(header_df) + 2

        df.to_excel(writer, index=False, startrow=row)
        row += len(df) + 2

        summary_df = pd.DataFrame({
            "Field": ["Total Base Value", "Total Tax", "Grand Total"],
            "Value": [fmt2(total_base_value), fmt2(total_tax), fmt2(grand_total)],
        })

        summary_df.to_excel(writer, index=False, startrow=row, header=False)