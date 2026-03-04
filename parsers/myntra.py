import pdfplumber
import pandas as pd
import re


# ------------------ HEADER EXTRACTION ------------------

def extract_po_header(pdf_path):

    party_name = "Myntra"
    po_no = ""
    po_date = ""
    po_expiry = ""
    shipping_address = ""
    gst_no = ""

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    for line in text.split("\n"):

        if "PO #:" in line:
            m = re.search(r'PO #:\s*([A-Z0-9\-]+)', line)
            if m:
                po_no = m.group(1).strip()

        if "PO Approved Date:" in line:
            m = re.search(r'PO Approved Date:\s*([\d\-]+)', line)
            if m:
                date_str = m.group(1).strip()
                # Convert YYYY-MM-DD to DD-MM-YYYY
                try:
                    from datetime import datetime
                    dt = datetime.strptime(date_str, "%Y-%m-%d")
                    po_date = dt.strftime("%d-%m-%Y")
                except:
                    po_date = date_str

        if "Estimated Shipment Date:" in line:
            m = re.search(r'Estimated Shipment Date:\s*([\d/]+)', line)
            if m:
                date_str = m.group(1).strip()
                # Convert DD/MM/YYYY to DD-MM-YYYY
                try:
                    from datetime import datetime
                    dt = datetime.strptime(date_str, "%d/%m/%Y")
                    po_expiry = dt.strftime("%d-%m-%Y")
                except:
                    po_expiry = date_str

        # GSTIN from SHIP TO section
        if "GSTIN#" in line:
            m = re.search(r'GSTIN#\s*([A-Z0-9]{15})', line)
            if m:
                gst_no = m.group(1).strip()

    # Shipping Address - extract pincode from SHIP TO section
    # Look for 6-digit pincode in the text
    ship_to_section = ""
    lines = text.split("\n")
    in_ship_to = False
    for line in lines:
        if "SHIP TO:" in line:
            in_ship_to = True
        elif "GSTIN#" in line and in_ship_to:
            break
        elif in_ship_to:
            ship_to_section += line + " "
    
    # Extract 6-digit pincode from ship to section
    pin_match = re.search(r'\b(\d{6})\b', ship_to_section)
    if pin_match:
        shipping_address = pin_match.group(1)

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

    for i, line in enumerate(all_lines):
        line = line.strip()

        # Match lines starting with BNPL (SKU code)
        if line.startswith("BNPL"):
            parts = line.split()
            
            if len(parts) < 13:
                continue
            
            try:
                sku_code = parts[0]
                hsn = parts[1]
                
                # Find EAN - Myntra has 13 or 14 digit EANs with various split patterns
                ean = ""
                ean_idx = 0
                
                # Pattern 1: Complete 13-14 digit on this line
                for j, p in enumerate(parts):
                    if p.isdigit() and len(p) in [13, 14]:
                        ean = p
                        ean_idx = j
                        break
                
                if not ean and i + 1 < len(all_lines):
                    next_line = all_lines[i + 1].strip()
                    next_parts = next_line.split()
                    
                    if next_parts:
                        # Pattern 2: 11-digit + 2-3 digit suffix from LAST word of next line (13-digit EANs)
                        for j, p in enumerate(parts):
                            if p.isdigit() and len(p) == 11:
                                suffix = next_parts[-1]
                                if suffix.isdigit() and len(suffix) in [2, 3]:
                                    potential_ean = p + suffix
                                    if len(potential_ean) in [13, 14]:
                                        # But check if there's a 9-digit after this 11-digit
                                        # If yes, this might be a 14-digit EAN pattern instead
                                        if j + 1 < len(parts) and parts[j + 1].isdigit() and len(parts[j + 1]) == 9:
                                            # This is a 14-digit pattern, skip this match
                                            continue
                                        ean = potential_ean
                                        ean_idx = j
                                        break
                        
                        # Pattern 2B: 11-digit + 2-3 digit from ANY position in next line
                        # (when last word is not a digit, look for digits elsewhere in the line)
                        if not ean:
                            for j, p in enumerate(parts):
                                if p.isdigit() and len(p) == 11:
                                    # Find any 2-3 digit number in next line
                                    for word in next_parts:
                                        if word.isdigit() and len(word) in [2, 3]:
                                            potential_ean = p + word
                                            if len(potential_ean) in [13, 14]:
                                                # Check if there's a 9-digit after on main line
                                                if j + 1 < len(parts) and parts[j + 1].isdigit() and len(parts[j + 1]) == 9:
                                                    continue
                                                ean = potential_ean
                                                ean_idx = j
                                                break
                                    if ean:
                                        break
                        
                        # Pattern 3: 11-digit + SECOND word of next line for 14-digit EANs
                        # (when there's a 9-digit number after the 11-digit on main line)
                        if not ean and len(next_parts) >= 2:
                            for j, p in enumerate(parts):
                                if p.isdigit() and len(p) == 11:
                                    # Check if next part is 9 digits
                                    if j + 1 < len(parts) and parts[j + 1].isdigit() and len(parts[j + 1]) == 9:
                                        second_word = next_parts[1]
                                        if second_word.isdigit() and len(second_word) == 5:
                                            # EAN = 11-digit + last 3 digits of 5-digit word = 14 digits
                                            # Actually: EAN = 11-digit + first 2 digits + last digit of second_word
                                            # Wait - let me check: 15289321452 + 52530
                                            # The pattern is: take 11-digit + full 5-digit from next line!
                                            # But some might need truncation...
                                            # Looking at master: should be exactly 14 digits
                                            # 15289321452 (11) + last 3 of 52530 (530) = 15289321452530 (14)
                                            potential_ean = p + second_word[-3:]
                                            if len(potential_ean) == 14:
                                                ean = potential_ean
                                                ean_idx = j
                                                break
                        
                        # Pattern 4: 11-digit + FIRST word of next line (for 2-digit suffix, 13-digit EANs)
                        if not ean:
                            for j, p in enumerate(parts):
                                if p.isdigit() and len(p) == 11:
                                    first_word = next_parts[0]
                                    if first_word.isdigit() and len(first_word) == 2:
                                        potential_ean = p + first_word
                                        if len(potential_ean) == 13:
                                            ean = potential_ean
                                            ean_idx = j
                                            break
                
                if not ean or len(ean) < 13:
                    continue
                
                # After EAN: Color Size StyleID Qty MRP List Landed IGST% IGST_Amt Total
                # Get last 7 numeric values
                try:
                    qty = int(float(parts[-7]))
                    mrp = float(parts[-6])
                    landed = float(parts[-4])
                    igst_pct = float(parts[-3])
                    total = float(parts[-1])
                    
                    # Validate values are reasonable
                    if qty <= 0 or mrp <= 0 or total <= 0:
                        continue
                        
                except (ValueError, IndexError):
                    # Skip if can't extract numeric values
                    continue
                
                # Product name is words between HSN and EAN (will be partial - updated from master)
                desc_parts = parts[2:ean_idx]
                product_name = " ".join(desc_parts).strip()
                
                items.append({
                    "Sr #": len(items) + 1,
                    "EAN": ean,  # Will be formatted as text by openpyxl
                    "Product Name": product_name,
                    "HSN Code": hsn,
                    "Quantity": qty,
                    "MRP": mrp,
                    "Base Rate": landed,
                    "GST %": igst_pct,
                    "Total": total,
                })
            
            except Exception as e:
                print(f"Error parsing line {i}: {e}")
                continue

    return pd.DataFrame(items)


# ------------------ SUMMARY EXTRACTION ------------------

def extract_summary(pdf_path):

    grand_total = 0.0

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                m = re.search(r'Grand Total:\s*([\d,]+\.?\d*)', line)
                if m:
                    grand_total = float(m.group(1).replace(",", ""))
                    break

    return grand_total


# ================== PUBLIC FUNCTION ==================

def convert_pdf_to_excel(pdf_path, output_excel_path):

    header_data = extract_po_header(pdf_path)
    products = extract_line_items(pdf_path)

    if products.empty:
        raise Exception("No line items found in Myntra PO")

    grand_total = extract_summary(pdf_path)
    total_base = round((products["Base Rate"] * products["Quantity"]).sum(), 2)
    total_tax = round(grand_total - total_base, 2)

    summary_data = {
        "Total Base Value": f"{total_base:.2f}",
        "Total Tax": f"{total_tax:.2f}",
        "Grand Total": f"{grand_total:.2f}",
    }

    # Write to Excel with proper text formatting for EAN
    from openpyxl import Workbook
    from openpyxl.styles import numbers
    
    wb = Workbook()
    ws = wb.active
    
    row_offset = 1
    
    # Header section
    for field, value in header_data.items():
        ws.cell(row=row_offset, column=1, value=field)
        ws.cell(row=row_offset, column=2, value=value)
        row_offset += 1
    
    row_offset += 2
    
    # Products table with headers
    headers = ["Sr #", "EAN", "Product Name", "HSN Code", "Quantity", "MRP", "Base Rate", "GST %", "Total"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=row_offset, column=col, value=header)
    row_offset += 1
    
    # Products data
    for _, row in products.iterrows():
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row_offset, column=col)
            value = row[header]
            
            # Format EAN as text
            if header == "EAN":
                cell.value = str(value)
                cell.number_format = '@'  # Text format
            else:
                cell.value = value
        row_offset += 1
    
    row_offset += 2
    
    # Summary section
    for field, value in summary_data.items():
        ws.cell(row=row_offset, column=1, value=field)
        ws.cell(row=row_offset, column=2, value=value)
        row_offset += 1
    
    wb.save(output_excel_path)