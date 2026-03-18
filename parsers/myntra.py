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
    items = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Extract tables instead of text
            tables = page.extract_tables()
            
            if not tables:
                continue
            
            for table in tables:
                for row in table:
                    if not row or len(row) < 10:
                        continue
                    
                    # Check if first cell starts with BNPL (SKU code)
                    first_cell = str(row[0] or "").strip()
                    if not first_cell.startswith("BNPL"):
                        continue
                    
                    try:
                        # Column structure based on your sample:
                        # 0=SKU, 1=HSN, 2=ProductName, 3=EAN(8-digit), 4=Another#, 5=Color, 
                        # 6=Size, 7=StyleID, 8=Qty, 9=MRP, 10=Rate1, 11=Rate2, 
                        # 12=CGST%, 13=CGST_Amt, 14=SGST%, 15=SGST_Amt, 16=Total
                        
                        sku_code = first_cell
                        hsn = str(row[1] or "").strip()
                        product_name = str(row[2] or "").strip()
                        ean = str(row[3] or "").strip()
                        
                        # Get numeric values from end
                        qty = int(float(str(row[8] or "0").strip()))
                        mrp = float(str(row[9] or "0").strip())
                        base_rate = float(str(row[11] or "0").strip())  # Use Rate2
                        cgst_pct = float(str(row[12] or "0").strip())
                        sgst_pct = float(str(row[14] or "0").strip())
                        total = float(str(row[16] or "0").strip())
                        
                        gst_pct = cgst_pct + sgst_pct
                        
                        # Validate
                        if qty <= 0 or mrp <= 0 or total <= 0 or not ean:
                            continue
                        
                        items.append({
                            "Sr #": len(items) + 1,
                            "EAN": ean,
                            "Product Name": product_name,
                            "HSN Code": hsn,
                            "Quantity": qty,
                            "MRP": mrp,
                            "Base Rate": base_rate,
                            "GST %": gst_pct,
                            "Total": total,
                        })
                        
                    except (ValueError, IndexError) as e:
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
