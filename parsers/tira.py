import pdfplumber
import pandas as pd
import re
from decimal import Decimal


# ------------------ HELPERS ------------------

def clean(x):
    return x.strip() if isinstance(x, str) else ""


def num(x):
    if not x:
        return ""
    m = re.search(r"\d+(\.\d+)?", str(x).replace(",", ""))
    return m.group() if m else ""


# ================== MAIN ==================

def convert_pdf_to_excel(pdf_path, output_excel_path):

    items = []

    po_no = ""
    po_date = ""
    po_expiry_date = ""
    gst_no = ""
    shipping_address = ""

    # keep as TEXT exactly
    total_basic_value = ""
    total_tax = ""
    grand_total = ""

    with pdfplumber.open(pdf_path) as pdf:

        # =====================================================
        # FULL TEXT (HEADER + SUMMARY)
        # =====================================================
        full_text = "\n".join([p.extract_text() or "" for p in pdf.pages])

        # -------- PO NO --------
        m = re.search(r"PO\s*NO\.?\s*[:\-]?\s*(\d+)", full_text, re.I)
        if m:
            po_no = m.group(1)

        # -------- PO DATE --------
        m = re.search(r"PO\s*Date\s*[:\-]?\s*([\d\.]+)", full_text, re.I)
        if m:
            po_date = m.group(1)

        # -------- DELIVERY DATE → PO EXPIRY DATE --------
        m = re.search(r"Delivery\s*Date\s*[:\-]?\s*([\d\.]+)", full_text, re.I)
        if m:
            po_expiry_date = m.group(1)

        # -------- SHIPPING ADDRESS + GST --------
        lines = [l.strip() for l in full_text.split("\n") if l.strip()]

        for i, line in enumerate(lines):
            if "delivery address" in line.lower():

                addr = []
                for j in range(i + 1, min(i + 15, len(lines))):

                    if "gstn" in lines[j].lower() or "gstin" in lines[j].lower():
                        gst_no = lines[j].split(":")[-1].strip()
                        break

                    addr.append(lines[j])

                shipping_address = "\n".join(addr)
                break

        # ================== SUMMARY — COPY EXACT ==================

        m = re.search(r"TOTAL\s+BASIC\s+VALUE\s*[:\-]?\s*INR\s*([\d,]+\.\d{2})", full_text, re.I)
        if m:
            total_basic_value = m.group(1)   # keep commas

        # try TOTAL TAX directly
        m = re.search(r"TOTAL\s+TAX\s*[:\-]?\s*INR\s*([\d,]+\.\d{2})", full_text, re.I)
        if m:
            total_tax = m.group(1)
        else:
            m1 = re.search(r"TOTAL\s+CGST\s*[:\-]?\s*INR\s*([\d,]+\.\d{2})", full_text, re.I)
            m2 = re.search(r"TOTAL\s+SGST\s*[:\-]?\s*INR\s*([\d,]+\.\d{2})", full_text, re.I)

            if m1 and m2:
                try:
                    cgst = Decimal(m1.group(1).replace(",", ""))
                    sgst = Decimal(m2.group(1).replace(",", ""))
                    total_tax = f"{(cgst + sgst):.2f}"
                except:
                    total_tax = ""

        m = re.search(r"Total\s+Order\s+Value\s*[:\-]?\s*INR\s*([\d,]+\.\d{2})", full_text, re.I)
        if m:
            grand_total = m.group(1)

        # =====================================================
        # LINE ITEMS TABLE
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

                if ("ean" in joined and "material" in joined and "quantity" in joined):
                    header = r
                    continue

                if not header:
                    continue

                def col(name):
                    for i, h in enumerate(header):
                        if name in h.lower():
                            if i < len(r):
                                return r[i]
                    return ""

                ean_val = num(col("ean"))
                if not ean_val:
                    continue

                qty = int(float(num(col("quantity")) or 0))
                mrp = num(col("mrp"))
                base = num(col("base")) or num(col("price"))
                cgst = num(col("cgst"))
                sgst = num(col("sgst"))

                try:
                    gst_pct = int(float(cgst or 0) + float(sgst or 0))
                except:
                    gst_pct = ""

                name_parts = []
                for i, h in enumerate(header):
                    if any(k in h.lower() for k in ["material", "description", "product"]):
                        if i < len(r) and r[i]:
                            name_parts.append(r[i])

                full_name = " ".join(name_parts)
                full_name = re.sub(r"\s{2,}", " ", full_name).strip()

                # ---------- SKIP DC / LOCATION ROWS (TIRA ONLY) ----------
                dc_words = ["dc", "warehouse", "bhiwandi", "kukse", "rrl", "bpc"]

                if any(w in full_name.lower() for w in dc_words):
                    continue

                items.append({
                    "EAN": ean_val,
                    "Product Name": full_name,
                    "HSN Code": col("hsn"),
                    "Quantity": qty,
                    "MRP": mrp,
                    "Base Rate": base,
                    "GST %": gst_pct,
                    "Total": num(col("total")),
                })

    if not items:
        raise Exception("No line items detected in Tira PO")

    df = pd.DataFrame(items)
    df.insert(0, "Sr #", range(1, len(df) + 1))

    # =====================================================
    # WRITE OUTPUT
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
                "Reliance Retail Limited (Tira)",
                po_no,
                po_date,
                po_expiry_date,
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
            "Value": [total_basic_value, total_tax, grand_total],
        })

        summary_df.to_excel(writer, index=False, startrow=row, header=False)
