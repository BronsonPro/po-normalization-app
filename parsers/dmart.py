import pdfplumber
import pandas as pd
import re


# ------------------ HEADER EXTRACTION ------------------

def extract_po_header(pdf_path):

    party_name = "DMart"
    po_no = ""
    po_date = ""
    po_expiry = ""
    shipping_address = ""
    gst_no = ""

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    for line in text.split("\n"):

        # PO Number in title line: "AvenueE-CommerceLtd PurchaseOrder 4501879572"
        m = re.search(r'PurchaseOrder\s+(\d+)', line)
        if m:
            po_no = m.group(1).strip()

        # PO Date and Validity: "PurchaseOrderDate:27.12.2025 POValidity:27.12.2025to27.01.2026"
        m = re.search(r'PurchaseOrderDate:([\d.]+)', line)
        if m:
            po_date = m.group(1).strip()

        m = re.search(r'POValidity:[\d.]+to([\d.]+)', line)
        if m:
            po_expiry = m.group(1).strip()

        # GST - Ship To GST (buyer GST)
        m = re.search(r'GST#([A-Z0-9]{15})', line)
        if m and not gst_no:
            gst_no = m.group(1).strip()

    # Shipping address - collect all lines from ShipTo to PurchaseOrderDate
    lines = text.split("\n")
    ship_section = ""
    in_ship = False
    for line in lines:
        if "ShipTo" in line:
            in_ship = True
        if in_ship:
            ship_section += line + " "
        if "PurchaseOrderDate" in line and in_ship:
            break
    
    # Find all 6-digit pincodes
    all_pins = re.findall(r'\b(\d{6})\b', ship_section)
    # The shipping pincode typically appears twice (Bill To and Ship To columns)
    # Count occurrences and take the one that appears most (excluding vendor code 103006)
    from collections import Counter
    pin_counts = Counter([pin for pin in all_pins if not pin.startswith('103')])
    if pin_counts:
        # Take the most common pincode
        shipping_address = pin_counts.most_common(1)[0][0]

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

    all_lines = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            all_lines.extend(text.split("\n"))

    items = []

    i = 0
    while i < len(all_lines):
        line = all_lines[i].strip()

        # Match product line 1:
        # "1 7452225170304 9603 BronsonPSelfClean EA 60 325.00 67.81 9.00 9.00 - - - 80.02 4,800.90"
        m = re.match(
            r'^(\d+)\s+'                        # SR No
            r'(\d{13})\s+'                      # EAN (13 digits)
            r'(\d{4,8})\s+'                     # HSN (partial - first part)
            r'(.+?)\s+'                         # Description part 1
            r'EA\s+'                            # UOM
            r'(\d+)\s+'                        # Qty
            r'([\d,]+\.?\d*)\s+'              # MRP
            r'([\d,]+\.?\d*)\s+'              # Basic Price
            r'([\d.]+)\s+'                     # CGST%
            r'([\d.]+)\s+'                     # SGST%
            r'.*?([\d,]+\.?\d*)\s+'           # Landed Price
            r'([\d,]+\.?\d*)$',               # Total Value
            line
        )

        if m:
            sr_no = m.group(1)
            ean = m.group(2)
            hsn_part1 = m.group(3)
            qty = m.group(5)
            mrp = m.group(6).replace(",", "")
            basic = m.group(7).replace(",", "")
            cgst = float(m.group(8))
            sgst = float(m.group(9))
            landed = m.group(10).replace(",", "")
            total = m.group(11).replace(",", "")

            gst_pct = round(cgst + sgst, 2)

            # Line 2 has: article_no + hsn_part2 + desc_part2 + 1.00 + cgst_val + sgst_val + ...
            # e.g. "140006766 2900 ingHairBrush-1N 1.00 366.17 366.17 - - -"
            desc_part2 = ""
            hsn_part2 = ""
            if i + 1 < len(all_lines):
                next_line = all_lines[i + 1].strip()
                m2 = re.match(r'^(\d+)\s+(\d+)\s+(.+?)\s+1\.00', next_line)
                if m2:
                    hsn_part2 = m2.group(2)
                    desc_part2 = m2.group(3).strip()

            # Full HSN code
            hsn = hsn_part1 + hsn_part2

            # Full description - clean up concatenated words
            desc_part1 = m.group(4).strip()
            product_name = (desc_part1 + " " + desc_part2).strip()

            items.append({
                "Sr #": int(sr_no),
                "EAN": ean,
                "Product Name": product_name,
                "HSN Code": hsn,
                "Quantity": int(qty),
                "MRP": float(mrp),
                "Base Rate": float(basic),
                "GST %": gst_pct,
                "Total": float(total),
            })

        i += 1

    return pd.DataFrame(items)


# ------------------ SUMMARY EXTRACTION ------------------

def extract_summary(pdf_path):

    grand_total = 0.0

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                m = re.match(r'^Total\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)$', line)
                if m:
                    grand_total = float(m.group(2).replace(",", ""))
                    break

    return grand_total


# ================== PUBLIC FUNCTION ==================

def convert_pdf_to_excel(pdf_path, output_excel_path):

    header_data = extract_po_header(pdf_path)
    products = extract_line_items(pdf_path)

    if products.empty:
        raise Exception("No line items found in DMart PO")

    grand_total = extract_summary(pdf_path)
    total_base = round((products["Base Rate"] * products["Quantity"]).sum(), 2)
    total_tax = round(grand_total - total_base, 2)

    summary_data = {
        "Total Base Value": f"{total_base:.2f}",
        "Total Tax": f"{total_tax:.2f}",
        "Grand Total": f"{grand_total:.2f}",
    }

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