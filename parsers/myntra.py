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
                    try:
                        from datetime import datetime
                        dt = datetime.strptime(date_str, "%d/%m/%Y")
                        po_expiry = dt.strftime("%d-%m-%Y")
                    except:
                        po_expiry = date_str

            if "GSTIN#" in line:
                m = re.search(r'GSTIN#\s*([A-Z0-9]{15})', line)
                if m:
                    gst_no = m.group(1).strip()

        # Extract pincode
        pin_match = re.search(r'\b(\d{6})\b', text)
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
        # Get all text from page 1
        text = pdf.pages[0].extract_text() or ""
        
        lines = text.split('\n')
        
        for line in lines:
            if 'BNPL' not in line:
                continue
            
            parts = line.strip().split()
            
            if len(parts) < 15:
                continue
            
            try:
                # Find SKU starting with BNPL
                sku_idx = -1
                for i, part in enumerate(parts):
                    if part.startswith('BNPL'):
                        sku_idx = i
                        break
                
                if sku_idx == -1:
                    continue
                
                sku_code = parts[sku_idx]
                hsn = parts[sku_idx + 1] if sku_idx + 1 < len(parts) else ""
                
                # Find 8-digit EAN
                ean = ""
                ean_idx = -1
                for i in range(sku_idx + 2, len(parts)):
                    if parts[i].isdigit() and len(parts[i]) == 8:
                        ean = parts[i]
                        ean_idx = i
                        break
                
                if not ean:
                    continue
                
                # Product name between HSN and EAN
                product_name = " ".join(parts[sku_idx + 2:ean_idx]).strip()
                
                # Numeric values from end
                total = float(parts[-1])
                sgst_amt = float(parts[-2])
                sgst_pct = float(parts[-3])
                cgst_amt = float(parts[-4])
                cgst_pct = float(parts[-5])
                base_rate2 = float(parts[-6])
                base_rate1 = float(parts[-7])
                mrp = float(parts[-8])
                qty = int(float(parts[-9]))
                
                gst_pct = cgst_pct + sgst_pct
                
                if qty <= 0 or mrp <= 0 or total <= 0:
                    continue
                
                items.append({
                    "Sr #": len(items) + 1,
                    "EAN": ean,
                    "Product Name": product_name,
                    "HSN Code": hsn,
                    "Quantity": qty,
                    "MRP": mrp,
                    "Base Rate": base_rate2,
                    "GST %": gst_pct,
                    "Total": total,
                })
                
            except (ValueError, IndexError):
                continue

    return pd.DataFrame(items)

# ------------------ SUMMARY EXTRACTION ------------------

def extract_summary(pdf_path):
    grand_total = 0.0
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            m = re.search(r'Grand Total:\s*([\d,]+\.?\d*)', text)
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

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active

    row_offset = 1

    for field, value in header_data.items():
        ws.cell(row=row_offset, column=1, value=field)
        ws.cell(row=row_offset, column=2, value=value)
        row_offset += 1

    row_offset += 2

    headers = ["Sr #", "EAN", "Product Name", "HSN Code", "Quantity", "MRP", "Base Rate", "GST %", "Total"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=row_offset, column=col, value=header)

    row_offset += 1

    for _, row in products.iterrows():
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row_offset, column=col)
            value = row[header]

            if header == "EAN":
                cell.value = str(value)
                cell.number_format = '@'
            else:
                cell.value = value

        row_offset += 1

    row_offset += 2

    for field, value in summary_data.items():
        ws.cell(row=row_offset, column=1, value=field)
        ws.cell(row=row_offset, column=2, value=value)
        row_offset += 1

    wb.save(output_excel_path)
