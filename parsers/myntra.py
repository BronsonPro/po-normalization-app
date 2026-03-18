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
        # Extract table from page 1 (diagnostic showed table headers exist)
        tables = pdf.pages[0].extract_tables()
        
        if not tables:
            return pd.DataFrame(items)
        
        # Find the product table (the one with column headers)
        header_row = None
        data_table = None
        
        for table in tables:
            if not table:
                continue
            # Look for header row containing SKU Code
            for row in table:
                if row and any('SKU' in str(cell) for cell in row):
                    header_row = row
                    data_table = table
                    break
            if header_row:
                break
        
        if not header_row:
            # Fallback to text-based parsing
            return extract_line_items_from_text(pdf_path)
        
        # Clean header names
        headers = [str(cell).strip().replace('\n', ' ') for cell in header_row]
        
        # Find column indices by name
        def find_col(keywords):
            for i, h in enumerate(headers):
                if any(kw.lower() in h.lower() for kw in keywords):
                    return i
            return -1
        
        sku_col = find_col(['SKU'])
        hsn_col = find_col(['HSN'])
        product_col = find_col(['Article', 'Name', 'Product'])
        ean_col = find_col(['Article Number', 'EAN', 'Vendor Article Number'])
        qty_col = find_col(['Quantity', 'Qty'])
        mrp_col = find_col(['MRP'])
        
        # For base rate and GST, we need to check what columns exist
        cgst_pct_col = find_col(['CGST %', 'CGST%'])
        sgst_pct_col = find_col(['SGST %', 'SGST%'])
        igst_pct_col = find_col(['IGST %', 'IGST%'])
        
        # Find rate columns - could be "Landed Cost", "Base Rate", etc.
        rate_cols = []
        for i, h in enumerate(headers):
            if any(kw in h.lower() for kw in ['cost', 'rate', 'price']) and 'mrp' not in h.lower():
                rate_cols.append(i)
        
        total_col = find_col(['Total', 'Amount'])
        
        # Now extract from text since table structure doesn't preserve all data properly
        # Use table only to understand structure, extract from text
        return extract_line_items_from_text(pdf_path)

def extract_line_items_from_text(pdf_path):
    """Extract using text with intelligent column detection"""
    items = []
    debug_info = []  # Collect debug info to show in UI
    
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""
        
        lines = text.split('\n')
        
        # Find header line to understand column order
        header_line = None
        for line in lines:
            if 'SKU Code' in line and 'HSN' in line:
                header_line = line
                break
        
        # Detect if IGST or CGST+SGST based on header
        is_igst = 'IGST' in header_line if header_line else False
        is_cgst_sgst = 'CGST' in header_line and 'SGST' in header_line if header_line else False
        
        debug_info.append(f"Header line: {header_line[:100] if header_line else 'NOT FOUND'}")
        debug_info.append(f"Detected IGST: {is_igst}, CGST+SGST: {is_cgst_sgst}")
        
        for line in lines:
            if 'BNPL' not in line:
                continue
            
            parts = line.strip().split()
            
            if len(parts) < 10:
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
                
                # HSN is right after SKU
                hsn = parts[sku_idx + 1] if sku_idx + 1 < len(parts) else ""
                
                # Find EAN - look for 8-digit or longer number (could be 8, 11, or 13 digits)
                ean = ""
                ean_idx = -1
                for i in range(sku_idx + 2, len(parts)):
                    if parts[i].isdigit() and len(parts[i]) >= 8:
                        ean = parts[i]
                        ean_idx = i
                        break
                
                if not ean:
                    continue
                
                # Product name is between HSN and EAN
                product_name = " ".join(parts[sku_idx + 2:ean_idx]).strip()
                
                # Now intelligently find numeric columns from the end
                # Total is always last
                total = float(parts[-1])
                
                debug_info.append(f"\n=== Product: {sku_code} ===")
                debug_info.append(f"Total parts: {len(parts)}")
                debug_info.append(f"Last 10 parts: {parts[-10:]}")
                
                # Determine format and extract accordingly
                if is_igst or (not is_cgst_sgst and len(parts) < sku_idx + 20):
                    # IGST format: ...Qty MRP Rate1 Rate2 IGST% IGST_Amt Total
                    debug_info.append("Using IGST format")
                    igst_amt = float(parts[-2])
                    gst_pct = float(parts[-3])
                    base_rate2 = float(parts[-4])
                    base_rate1 = float(parts[-5])
                    mrp = float(parts[-6])
                    qty = int(float(parts[-7]))
                    debug_info.append(f"Qty={qty}, MRP={mrp}, Rate={base_rate2}, GST%={gst_pct}, Total={total}")
                else:
                    # CGST+SGST format: ...Qty MRP Rate1 Rate2 CGST% CGST_Amt SGST% SGST_Amt Total
                    debug_info.append("Using CGST+SGST format")
                    sgst_amt = float(parts[-2])
                    sgst_pct = float(parts[-3])
                    cgst_amt = float(parts[-4])
                    cgst_pct = float(parts[-5])
                    base_rate2 = float(parts[-6])
                    base_rate1 = float(parts[-7])
                    mrp = float(parts[-8])
                    qty = int(float(parts[-9]))
                    gst_pct = cgst_pct + sgst_pct
                    debug_info.append(f"Qty={qty}, MRP={mrp}, Rate={base_rate2}, GST%={gst_pct}, Total={total}")
                
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
                
            except (ValueError, IndexError) as e:
                debug_info.append(f"ERROR parsing: {str(e)}")
                continue
    
    # Show debug in Streamlit
    import streamlit as st
    with st.expander("🔍 Myntra Parser Debug Info"):
        st.text("\n".join(debug_info))
    
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
