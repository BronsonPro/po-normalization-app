import pdfplumber
import pandas as pd
import re


# ------------------ HEADER EXTRACTION ------------------

def extract_po_header(pdf_path):

    party_name = "D-MART"
    po_no = ""
    po_date = ""
    po_expiry = ""
    shipping_address = ""
    gst_no = ""

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    for line in text.split("\n"):

        m = re.search(r'PurchaseOrder\s+(\d+)', line)
        if m:
            po_no = m.group(1).strip()

        m = re.search(r'PurchaseOrderDate:([\d.]+)', line)
        if m:
            po_date = m.group(1).strip()

        m = re.search(r'POValidity:[\d.]+to([\d.]+)', line)
        if m:
            po_expiry = m.group(1).strip()

        m = re.search(r'GST#([A-Z0-9]{15})', line)
        if m and not gst_no:
            gst_no = m.group(1).strip()

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
    
    all_pins = re.findall(r'\b(\d{6})\b', ship_section)
    from collections import Counter
    pin_counts = Counter([pin for pin in all_pins if not pin.startswith('103')])
    if pin_counts:
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

    items = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            
            # Extract table
            tables = page.extract_tables()
            
            if not tables:
                continue
            
            table = tables[0]
            
            # Find header row (contains "Sr" and "EAN/Article")
            header_row_idx = None
            for idx, row in enumerate(table):
                if row and any(cell and "Sr" in str(cell) and "No" in str(cell) for cell in row[:2]):
                    header_row_idx = idx
                    break
            
            if header_row_idx is None:
                continue
            
            headers = table[header_row_idx]
            
            # Find column indices
            col_sr = None
            col_ean = None
            col_hsn = None
            col_desc = None
            col_qty = None
            col_mrp = None
            col_basic = None
            col_cgst = None
            col_sgst = None
            col_igst = None
            col_landed = None
            col_total = None
            
            for idx, h in enumerate(headers):
                if not h:
                    continue
                h_lower = str(h).lower().replace("\n", "").replace(" ", "")
                
                if "sr" in h_lower and "no" in h_lower:
                    col_sr = idx
                elif "ean" in h_lower or "article" in h_lower:
                    col_ean = idx
                elif "hsn" in h_lower and "code" in h_lower:
                    col_hsn = idx
                elif "description" in h_lower:
                    col_desc = idx
                elif "poqty" in h_lower or ("qty" in h_lower and "po" in h_lower):
                    col_qty = idx
                elif "mrp" in h_lower:
                    col_mrp = idx
                elif "basic" in h_lower and "price" in h_lower:
                    col_basic = idx
                elif "cgst%" in h_lower or h_lower == "cgst%":
                    col_cgst = idx
                elif "sgst%" in h_lower or "sgst%cess%" in h_lower:
                    col_sgst = idx
                elif "igst%" in h_lower or "igst%ugst%" in h_lower:
                    col_igst = idx
                elif "landed" in h_lower:
                    col_landed = idx
                elif "total" in h_lower and "value" in h_lower:
                    col_total = idx
            
            # Process data rows
            for row_idx in range(header_row_idx + 1, len(table)):
                row = table[row_idx]
                
                if not row or not row[col_sr]:
                    continue
                
                sr_text = str(row[col_sr]).strip()
                
                # Check if this is a data row (starts with number)
                if not sr_text or not sr_text[0].isdigit():
                    continue
                
                # Extract SR number
                sr_match = re.match(r'(\d+)', sr_text)
                if not sr_match:
                    continue
                
                sr_no = sr_match.group(1)
                
                # Extract values
                ean = str(row[col_ean] or "").strip().split()[0] if col_ean is not None else ""
                hsn = str(row[col_hsn] or "").strip().replace("\n", "") if col_hsn is not None else ""
                desc = str(row[col_desc] or "").strip().replace("\n", " ") if col_desc is not None else ""
                qty = str(row[col_qty] or "").strip() if col_qty is not None else "0"
                mrp = str(row[col_mrp] or "0").strip().replace(",", "") if col_mrp is not None else "0"
                basic = str(row[col_basic] or "0").strip().replace(",", "") if col_basic is not None else "0"
                landed = str(row[col_landed] or "0").strip().replace(",", "") if col_landed is not None else "0"
                total = str(row[col_total] or "0").strip().replace(",", "") if col_total is not None else "0"
                
                # Extract GST % - check which format
                cgst_text = str(row[col_cgst] or "").strip() if col_cgst is not None else "-"
                sgst_text = str(row[col_sgst] or "").strip().split()[0] if col_sgst is not None else "-"  # Split to get just SGST, not CESS
                igst_text = str(row[col_igst] or "").strip().split()[0] if col_igst is not None else "-"  # Split to get just IGST, not UGST
                
                # Determine GST %
                gst_pct = 0.0
                
                # Check if IGST format (IGST has value, CGST/SGST are dashes)
                if igst_text != "-" and igst_text:
                    try:
                        gst_pct = float(igst_text)
                    except:
                        gst_pct = 0.0
                # Check if CGST+SGST format
                elif cgst_text != "-" and sgst_text != "-":
                    try:
                        cgst_val = float(cgst_text)
                        sgst_val = float(sgst_text)
                        gst_pct = cgst_val + sgst_val
                    except:
                        gst_pct = 0.0
                
                # Clean up values
                try:
                    qty_int = int(qty)
                    mrp_float = float(mrp) if mrp else 0.0
                    basic_float = float(basic) if basic else 0.0
                    total_float = float(total) if total else 0.0
                except:
                    continue
                
                items.append({
                    "Sr #": int(sr_no),
                    "EAN": ean,
                    "Product Name": desc,
                    "HSN Code": hsn,
                    "Quantity": qty_int,
                    "MRP": mrp_float,
                    "Base Rate": basic_float,
                    "GST %": gst_pct,
                    "Total": total_float,
                })

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
