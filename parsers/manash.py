import pdfplumber
import pandas as pd
import re


# ------------------ HEADER EXTRACTION ------------------

def extract_po_header(pdf_path):

    party_name = "Manash"
    po_no = ""
    po_date = ""
    po_expiry = ""
    shipping_address = ""
    gst_no = ""

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    for line in text.split("\n"):

        if "PO Number" in line:
            m = re.search(r'PO Number\s*:\s*([\w\d]+)', line)
            if m:
                po_no = m.group(1).strip()

        if "Date" in line and "Validity" not in line and "PO Number" not in line:
            m = re.search(r'Date\s*:\s*([\d.]+)', line)
            if m:
                po_date = m.group(1).strip()

        if "Validity End Date" in line:
            m = re.search(r'Validity End Date\s*:\s*([\d.]+)', line)
            if m:
                po_expiry = m.group(1).strip()

        if "GST No:" in line and "Supplier" not in line and "GST No." not in line:
            m = re.search(r'GST No:\s*([A-Z0-9]{15})', line)
            if m:
                gst_no = m.group(1).strip()

    # Shipping Address - extract pincode from delivery address section
    # Look for lines mentioning delivery/shipping and extract last 6-digit pincode
    delivery_lines = ""
    for line in text.split("\n"):
        if any(keyword in line for keyword in ["Delivery Address", "Ship To", "Warehouse", "Village", "Gala"]):
            delivery_lines += line + " "
    
    # Find all 6-digit numbers in delivery section and take the last one
    # (Last one is typically the actual delivery address, not vendor address)
    all_pins = re.findall(r'\b(\d{6})\b', delivery_lines)
    if all_pins:
        # Take the last pincode found (delivery address typically appears last)
        shipping_address = all_pins[-1]

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

        # Match SR No at start of line - e.g. "1 PPLB89722NM1 360.00 96033020 EA 2 137.34 ..."
        m = re.match(
            r'^(\d+)\s+(PPLB\S*)\s*'           # SR No + SKU
            r'(\d{12,14})?\s*'                  # Optional EAN (12-14 digits)
            r'([\d,]+\.?\d*)\s+'               # MRP
            r'(\d{6,8})\s+'                    # HSN Code
            r'EA\s+'                            # UOM
            r'(\d+)\s+'                        # Qty
            r'([\d,]+\.?\d*)\s+'              # Rate/Unit
            r'[\d.]+\s+'                       # SGST%
            r'[\d,]+\s+'                       # SGST Value
            r'[\d.]+\s+'                       # CGST%
            r'[\d,]+\s+'                       # CGST Value
            r'([\d,]+\.?\d*)',                 # Amount
            line
        )

        if m:
            sr_no = m.group(1)
            sku = m.group(2)
            ean = m.group(3) or ""
            mrp = m.group(4).replace(",", "")
            hsn = m.group(5)
            qty = m.group(6)
            rate = m.group(7).replace(",", "")
            amount = m.group(8).replace(",", "")

            # Extract GST % from line
            gst_m = re.search(r'EA\s+\d+\s+[\d,.]+\s+([\d.]+)\s+', line)
            gst_pct = float(gst_m.group(1)) * 2 if gst_m else 18.0

            # Product name is on NEXT line
            product_name = ""
            if i + 1 < len(all_lines):
                next_line = all_lines[i + 1].strip()
                # Next line is product name if it doesn't start with a number
                if next_line and not re.match(r'^\d+\s+PPLB', next_line) and \
                   not next_line.startswith("SR No") and \
                   not next_line.startswith("This is") and \
                   not next_line.startswith("Page"):
                    product_name = next_line

            # If EAN not in line, try to extract from SKU code
            if not ean:
                # Extract full digit sequence from SKU (12-14 digits)
                ean_in_sku = re.search(r'(\d{12,14})', sku)
                if ean_in_sku:
                    ean = ean_in_sku.group(1)

            items.append({
                "Sr #": int(sr_no),
                "EAN": ean,
                "Product Name": product_name,
                "HSN Code": hsn,
                "Quantity": int(qty),
                "MRP": float(mrp),
                "Base Rate": float(rate),
                "GST %": gst_pct,
                "Total": float(amount),
            })

        i += 1

    return pd.DataFrame(items)


# ------------------ SUMMARY EXTRACTION ------------------

def extract_summary(pdf_path):

    total_qty = 0
    total_tax = 0.0
    grand_total = 0.0

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                # Total line: "Total 129 492.68 492.68 6,934.54"
                m = re.match(r'^Total\s+(\d+)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)', line)
                if m:
                    total_qty = int(m.group(1))
                    sgst = float(m.group(2).replace(",", ""))
                    cgst = float(m.group(3).replace(",", ""))
                    grand_total = float(m.group(4).replace(",", ""))
                    total_tax = round(sgst + cgst, 2)
                    break

    total_base = round(grand_total - total_tax, 2)

    return {
        "Total Base Value": f"{total_base:.2f}",
        "Total Tax": f"{total_tax:.2f}",
        "Grand Total": f"{grand_total:.2f}",
    }


# ================== PUBLIC FUNCTION ==================

def convert_pdf_to_excel(pdf_path, output_excel_path):

    # Extract all sections
    header_data = extract_po_header(pdf_path)
    products = extract_line_items(pdf_path)
    summary_data = extract_summary(pdf_path)

    if products.empty:
        raise Exception("No line items found in Manash PO")

    # Write normalized output Excel
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        row_offset = 0

        # Header section
        header_df = pd.DataFrame({
            "Field": list(header_data.keys()),
            "Value": list(header_data.values()),
        })
        header_df.to_excel(writer, index=False, startrow=row_offset, header=False)
        row_offset += len(header_df) + 2

        # Products table
        products.to_excel(writer, index=False, startrow=row_offset)
        row_offset += len(products) + 2

        # Summary section
        summary_df = pd.DataFrame({
            "Field": list(summary_data.keys()),
            "Value": list(summary_data.values()),
        })
        summary_df.to_excel(writer, index=False, startrow=row_offset, header=False)