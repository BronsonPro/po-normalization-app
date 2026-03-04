import pdfplumber
import pandas as pd
import re


# ------------------ HEADER EXTRACTION ------------------

def extract_po_header(pdf_path):

    party_name = "Sliksync"
    po_no = ""
    po_date = ""
    po_expiry = ""
    shipping_address = ""
    gst_no = ""

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    lines = text.split("\n")

    # Po number: SKPO/25-26/393
    for line in lines:
        if "Po number:" in line or "PO number:" in line:
            m = re.search(r'Po number:\s*([A-Z0-9\-/]+)', line, re.IGNORECASE)
            if m:
                po_no = m.group(1).strip()

        # Date : 12-Feb-2026
        if "Date :" in line and not "Deliver" in line:
            m = re.search(r'Date\s*:\s*([\d\-A-Za-z]+)', line)
            if m:
                date_str = m.group(1).strip()
                # Convert "12-Feb-2026" to "12-02-2026"
                try:
                    from datetime import datetime
                    dt = datetime.strptime(date_str, "%d-%b-%Y")
                    po_date = dt.strftime("%d-%m-%Y")
                except:
                    po_date = date_str

    # Shipping Address - extract from "Deliver To" section
    deliver_section = ""
    in_deliver = False
    for line in lines:
        if "Deliver To" in line:
            in_deliver = True
        elif in_deliver and "GSTIN" in line:
            deliver_section += line
            break
        elif in_deliver:
            deliver_section += line + " "
    
    # Extract pincode from deliver section (6-digit number)
    pin_match = re.search(r'\b(\d{6})\b', deliver_section)
    if pin_match:
        shipping_address = pin_match.group(1)

    # GST from deliver section (GSTIN in deliver to section)
    gst_match = re.search(r'GSTIN\s*:\s*([A-Z0-9]{15})', deliver_section)
    if gst_match:
        gst_no = gst_match.group(1)

    # PO Expiry - Slikk doesn't have expiry in PO, leave blank
    po_expiry = ""

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

    with pdfplumber.open(pdf_path) as pdf:
        # Products are on page 2
        if len(pdf.pages) < 2:
            return pd.DataFrame()
        
        # Try table extraction first
        tables = pdf.pages[1].extract_tables()
        
        if not tables:
            return pd.DataFrame()
        
        # Find the main product table (has Description, SKU ID, HSN Code columns)
        product_table = None
        for table in tables:
            if table and len(table) > 1:
                # Check if first row has key headers
                header_row = [str(cell).lower() if cell else "" for cell in table[0]]
                if "description" in " ".join(header_row) and "sku" in " ".join(header_row):
                    product_table = table
                    break
        
        if not product_table:
            return pd.DataFrame()
        
        # Extract products
        items = []
        
        # Find column indices from header
        header = product_table[0]
        desc_idx = None
        sku_idx = None
        hsn_idx = None
        qty_idx = None
        mrp_idx = None
        purchase_idx = None
        gst_rate_idx = None
        total_purchase_idx = None
        
        for i, cell in enumerate(header):
            cell_str = str(cell).lower() if cell else ""
            if "description" in cell_str:
                desc_idx = i
            elif "sku" in cell_str:
                sku_idx = i
            elif "hsn" in cell_str:
                hsn_idx = i
            elif "allocate" in cell_str or ("qty" in cell_str and "allocate" not in cell_str):
                # Find "Allocate qty" column
                if "allocate" in cell_str:
                    qty_idx = i
            elif "mrp" == cell_str:
                mrp_idx = i
            elif "purchase price/pu" in cell_str or "purchase" in cell_str and "price" in cell_str and "total" not in cell_str:
                purchase_idx = i
            elif "gst rate on purchase" in cell_str:
                gst_rate_idx = i
            elif "total purchase price with gst" in cell_str:
                total_purchase_idx = i
        
        # If qty_idx not found, look for column after HSN
        if qty_idx is None and hsn_idx is not None:
            qty_idx = hsn_idx + 1
        
        # Process data rows
        for row_idx, row in enumerate(product_table[1:], 1):
            if not row or len(row) < 5:
                continue
            
            try:
                # Extract values
                description = str(row[desc_idx]).strip() if desc_idx is not None and row[desc_idx] else ""
                
                # SKU ID - extract digits and concatenate (PDF may have text mixed in)
                sku_id = ""
                if sku_idx is not None and row[sku_idx]:
                    sku_str = str(row[sku_idx]).strip()
                    # First try to find 13-14 digit number directly
                    digits = re.findall(r'\d{13,14}', sku_str)
                    if digits:
                        sku_id = digits[0]
                    else:
                        # Extract all digits and concatenate
                        all_digits = ''.join(re.findall(r'\d', sku_str))
                        if len(all_digits) >= 13:
                            # Prefer 13 digits (standard EAN-13)
                            # Only use 14 if the first 13 don't match master
                            sku_id = all_digits[:13]
                
                # HSN Code - extract only digits (8 digits)
                hsn = ""
                if hsn_idx is not None and row[hsn_idx]:
                    hsn_str = str(row[hsn_idx]).strip()
                    # Extract 8-digit HSN code
                    hsn_match = re.search(r'\b(\d{8})\b', hsn_str)
                    if hsn_match:
                        hsn = hsn_match.group(1)
                    else:
                        # Fallback: take any 6-8 digit number
                        digits = re.findall(r'\d{6,8}', hsn_str)
                        if digits:
                            hsn = digits[0]
                
                # Qty
                qty = 0
                if qty_idx is not None and row[qty_idx]:
                    try:
                        qty = int(float(str(row[qty_idx]).strip()))
                    except:
                        pass
                
                if qty <= 0 or not sku_id or len(sku_id) < 13:
                    continue
                
                # MRP
                mrp = 0.0
                if mrp_idx is not None and row[mrp_idx]:
                    try:
                        mrp = float(str(row[mrp_idx]).replace(",", "").strip())
                    except:
                        pass
                
                # Purchase price per unit (Base Rate) - column 13
                base_rate = 0.0
                if len(row) > 13 and row[13]:
                    try:
                        base_rate = float(str(row[13]).replace(",", "").strip())
                    except:
                        pass
                
                # GST Rate - extract from column 15
                gst_pct = 18.0
                if len(row) > 15 and row[15]:
                    gst_str = str(row[15]).strip()
                    m = re.search(r'(\d+)', gst_str)
                    if m:
                        gst_pct = float(m.group(1))
                
                # Total purchase with GST - column 17
                total = 0.0
                if len(row) > 17 and row[17]:
                    try:
                        total = float(str(row[17]).replace(",", "").replace(" ", "").strip())
                    except:
                        pass
                
                items.append({
                    "Sr #": len(items) + 1,
                    "EAN": sku_id,
                    "Product Name": description,
                    "HSN Code": hsn,
                    "Quantity": qty,
                    "MRP": mrp,
                    "Base Rate": base_rate,
                    "GST %": gst_pct,
                    "Total": total,
                })
            
            except Exception as e:
                print(f"Error parsing row {row_idx}: {e}")
                continue
        
        return pd.DataFrame(items)


# ------------------ SUMMARY EXTRACTION ------------------

def extract_summary(pdf_path):

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    # Extract from summary table
    # Total 46,562
    # IGST 7,103
    # Sub Total 39,459
    
    grand_total = 0.0
    total_tax = 0.0
    
    for line in text.split("\n"):
        if line.strip().startswith("Total") and not "Sub Total" in line:
            m = re.search(r'Total\s+([\d,]+)', line)
            if m:
                grand_total = float(m.group(1).replace(",", ""))
        
        if "IGST" in line:
            m = re.search(r'IGST\s+([\d,]+)', line)
            if m:
                total_tax = float(m.group(1).replace(",", ""))
    
    total_base = round(grand_total - total_tax, 2)

    return {
        "Total Base Value": f"{total_base:.2f}",
        "Total Tax": f"{total_tax:.2f}",
        "Grand Total": f"{grand_total:.2f}",
    }


# ================== PUBLIC FUNCTION ==================

def convert_pdf_to_excel(pdf_path, output_excel_path):

    header_data = extract_po_header(pdf_path)
    products = extract_line_items(pdf_path)

    if products.empty:
        raise Exception("No line items found in Slikk PO")

    summary_data = extract_summary(pdf_path)

    # Write to Excel with proper formatting
    from openpyxl import Workbook
    
    wb = Workbook()
    ws = wb.active
    
    row_offset = 1
    
    # Header section
    for field, value in header_data.items():
        ws.cell(row=row_offset, column=1, value=field)
        ws.cell(row=row_offset, column=2, value=value)
        row_offset += 1
    
    row_offset += 2
    
    # Products table
    headers = ["Sr #", "EAN", "Product Name", "HSN Code", "Quantity", "MRP", "Base Rate", "GST %", "Total"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=row_offset, column=col, value=header)
    row_offset += 1
    
    # Products data
    for _, row in products.iterrows():
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row_offset, column=col)
            value = row[header]
            
            # Format EAN and HSN as text
            if header == "EAN" or header == "HSN Code":
                cell.value = str(value)
                cell.number_format = '@'
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