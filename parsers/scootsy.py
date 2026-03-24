"""
Scootsy Parser - EXPERT VERSION
Extracts data exactly as it appears in the PDF
"""

import pdfplumber
import pandas as pd
import re


def convert_pdf_to_excel(pdf_path, output_path):
    """
    Convert Scootsy PDF to Excel format.
    Maps Item Code to EAN from master file if available.
    """
    
    all_rows = []
    
    # Extract header info
    header_data = {
        "PO No": "",
        "PO Date": "",
        "PO Expiry Date": "",
        "Shipping Address": "",
        "GST No": ""
    }
    
    # Try to load master file for EAN mapping
    master_ean_map = {}
    try:
        import os
        master_dir = os.path.dirname(pdf_path)
        master_path = os.path.join(master_dir, "Scootsy Master.xlsx")
        if not os.path.exists(master_path):
            master_path = os.path.join(master_dir, "Masterfile_Scootsy.xlsx")
        if not os.path.exists(master_path):
            master_path = os.path.join(master_dir, "Scootsy_master.xlsx")
        
        if os.path.exists(master_path):
            master_df = pd.read_excel(master_path)
            # Try both column name possibilities
            if 'Brand SKU Code' in master_df.columns and 'Item Code' in master_df.columns:
                for _, row in master_df.iterrows():
                    item_code = str(int(float(row['Item Code']))) if pd.notna(row['Item Code']) else None
                    ean = str(int(float(row['Brand SKU Code']))) if pd.notna(row['Brand SKU Code']) else None
                    if item_code and ean:
                        master_ean_map[item_code] = ean
            elif 'EAN' in master_df.columns and 'Item Code' in master_df.columns:
                for _, row in master_df.iterrows():
                    item_code = str(int(float(row['Item Code']))) if pd.notna(row['Item Code']) else None
                    ean = str(int(float(row['EAN']))) if pd.notna(row['EAN']) else None
                    if item_code and ean:
                        master_ean_map[item_code] = ean
    except Exception as e:
        # If master not found, EAN will be empty
        pass
    
    with pdfplumber.open(pdf_path) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""
        
        # Extract PO Number
        po_match = re.search(r'PO No\s*:\s*([A-Z0-9]+)', first_page_text)
        if po_match:
            header_data["PO No"] = po_match.group(1)
        
        # Extract PO Date
        date_match = re.search(r'PO Date\s*:\s*([A-Za-z]+\s+\d+,\s+\d{4})', first_page_text)
        if date_match:
            header_data["PO Date"] = date_match.group(1)
        
        # Extract PO Expiry Date
        expiry_match = re.search(r'PO Expiry Date:\s*([A-Za-z]+\s+\d+,\s+\d{4})', first_page_text)
        if expiry_match:
            header_data["PO Expiry Date"] = expiry_match.group(1)
        
        # Extract table data from all pages
        for page_num, page in enumerate(pdf.pages):
            # For page 2+, use more lenient table extraction settings
            # to capture rows at the top without borders
            if page_num > 0:
                # Use explicit vertical lines and snap tolerance to catch borderless rows
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "snap_tolerance": 5,
                    "join_tolerance": 5,
                    "edge_min_length": 10
                })
            else:
                tables = page.extract_tables()
            
            if not tables:
                continue
            
            # Extract shipping address from table (cleaner than text extraction)
            for row in tables[0]:
                if row and len(row) > 9 and row[9] and 'PJTJ' in str(row[9]) and not header_data["Shipping Address"]:
                    addr_lines = str(row[9]).split('\n')
                    # Take lines before Contact and GSTIN
                    clean_lines = []
                    for line in addr_lines:
                        if 'GSTIN' in line or 'GST' in line:
                            # Extract GSTIN number - try multiple patterns
                            gstin_match = re.search(r'(?:GSTIN|GST)[:\s-]*([A-Z0-9]{15})', line, re.IGNORECASE)
                            if gstin_match:
                                header_data["GST No"] = gstin_match.group(1)
                            break
                        if 'Contact' in line:
                            break
                        clean_lines.append(line.strip())
                    header_data["Shipping Address"] = ', '.join(clean_lines)[:250]
            
            # Fallback: Try to extract GSTIN from first page text if not found
            if not header_data["GST No"]:
                gstin_match = re.search(r'(?:GSTIN|GST)[:\s-]*([A-Z0-9]{15})', first_page_text, re.IGNORECASE)
                if gstin_match:
                    header_data["GST No"] = gstin_match.group(1)
            
            # Track which serial numbers we've processed
            processed_sr_nos = set()
            
            for table in tables:
                for row_idx, row in enumerate(table):
                    first_cell = str(row[0] or "").strip()
                    
                    if not row or len(row) < 10:
                        continue
                    
                    # Check if it's a data row (first column is a number)
                    if not first_cell or not first_cell.isdigit():
                        continue
                    
                    # Skip if already processed (avoid duplicates)
                    if first_cell in processed_sr_nos:
                        continue
                    processed_sr_nos.add(first_cell)
                    
                    # Detect table format based on row length
                    # Page 1: 19 columns with None at index 9
                    # Page 2+: 18 columns without the None
                    has_extra_column = len(row) >= 19
                    
                    # Extract data from correct columns with defensive length checks
                    sr_no = first_cell
                    item_code = str(row[1] or "").strip() if len(row) > 1 else ""
                    product_name = str(row[2] or "").strip().replace('\n', ' ') if len(row) > 2 else ""
                    hsn = str(row[3] or "").strip() if len(row) > 3 else ""
                    qty = str(row[4] or "").strip() if len(row) > 4 else ""
                    mrp = str(row[5] or "").strip() if len(row) > 5 else ""
                    base_rate = str(row[6] or "").strip() if len(row) > 6 else ""
                    
                    # GST Rate - indices depend on table format
                    if has_extra_column:
                        # 19-column format (page 1): None at index 9, everything shifted by 1
                        cgst_rate = str(row[8] or "").strip() if len(row) > 8 else ""
                        sgst_rate = str(row[11] or "").strip() if len(row) > 11 else ""
                        igst_rate = str(row[13] or "").strip() if len(row) > 13 else ""
                        total = str(row[18] or "").strip() if len(row) > 18 else ""
                    else:
                        # 18-column format (page 2+): No extra None column
                        cgst_rate = str(row[8] or "").strip() if len(row) > 8 else ""
                        sgst_rate = str(row[10] or "").strip() if len(row) > 10 else ""
                        igst_rate = str(row[12] or "").strip() if len(row) > 12 else ""
                        total = str(row[17] or "").strip() if len(row) > 17 else ""
                    
                    # Use IGST if present, otherwise CGST+SGST
                    gst_rate = ""
                    try:
                        # Convert to float and check
                        igst_val = float(igst_rate) if igst_rate and igst_rate.replace('.','').replace('-','').isdigit() else 0
                        cgst_val = float(cgst_rate) if cgst_rate and cgst_rate.replace('.','').replace('-','').isdigit() else 0
                        sgst_val = float(sgst_rate) if sgst_rate and sgst_rate.replace('.','').replace('-','').isdigit() else 0
                        
                        if igst_val > 0:
                            gst_rate = str(int(igst_val))  # IGST
                        elif cgst_val > 0 or sgst_val > 0:
                            gst_rate = str(int(cgst_val + sgst_val))  # CGST + SGST
                        else:
                            # All are 0, leave blank
                            gst_rate = ""
                    except Exception as e:
                        # If conversion fails, leave blank
                        gst_rate = ""
                    
                    # Map Item Code to EAN if available from master
                    item_code_clean = str(int(float(item_code))) if item_code and item_code.replace('.','').replace('-','').isdigit() else ""
                    ean = master_ean_map.get(item_code_clean, "")  # Get EAN from master map
                    
                    all_rows.append([
                        sr_no,
                        ean,  # EAN from master file lookup
                        item_code,
                        product_name,
                        hsn,
                        qty,
                        mrp,
                        base_rate,
                        gst_rate,
                        total
                    ])
            
            # FALLBACK: Extract rows from text that weren't captured in tables
            # This handles rows at page tops without borders
            page_text = page.extract_text() or ""
            text_lines = page_text.split('\n')
            
            for line in text_lines:
                # Look for lines that start with a number followed by item code pattern
                # Pattern: "5 380147 Product Name 96032900 48 550.00 ..."
                match = re.match(r'^(\d+)\s+(\d{6})\s+(.+)', line)
                if match:
                    sr_no = match.group(1)
                    
                    # Skip if already processed from table extraction
                    if sr_no in processed_sr_nos:
                        continue
                    
                    # Try to parse the entire line as space-separated values
                    parts = line.split()
                    if len(parts) < 10:
                        continue
                    
                    # Extract fields (this is approximate - adjust indices as needed)
                    item_code = parts[1] if len(parts) > 1 else ""
                    
                    # Product name spans multiple parts until HSN (8-digit number)
                    product_parts = []
                    hsn_idx = -1
                    for i in range(2, len(parts)):
                        if len(parts[i]) == 8 and parts[i].isdigit():
                            hsn = parts[i]
                            hsn_idx = i
                            break
                        product_parts.append(parts[i])
                    
                    if hsn_idx == -1:
                        continue
                    
                    product_name = ' '.join(product_parts)
                    
                    # After HSN: Qty, MRP, Base Cost, Taxable Value, GST rates, Total
                    remaining = parts[hsn_idx + 1:]
                    if len(remaining) < 8:
                        continue
                    
                    qty = remaining[0] if len(remaining) > 0 else ""
                    mrp = remaining[1] if len(remaining) > 1 else ""
                    base_rate = remaining[2] if len(remaining) > 2 else ""
                    
                    # Try to find IGST rate (usually around index 6-8 in remaining)
                    # This is approximate - the exact index depends on format
                    igst_rate = ""
                    for val in remaining[5:10]:
                        try:
                            if float(val) > 0 and float(val) <= 28:  # GST rates are 0-28%
                                igst_rate = val
                                break
                        except:
                            continue
                    
                    gst_rate = str(int(float(igst_rate))) if igst_rate else ""
                    
                    # Total is usually the last value
                    total = remaining[-1] if remaining else ""
                    
                    # Map Item Code to EAN
                    item_code_clean = str(int(float(item_code))) if item_code and item_code.replace('.','').replace('-','').isdigit() else ""
                    ean = master_ean_map.get(item_code_clean, "")
                    
                    all_rows.append([
                        sr_no,
                        ean,
                        item_code,
                        product_name,
                        hsn,
                        qty,
                        mrp,
                        base_rate,
                        gst_rate,
                        total
                    ])
                    
                    processed_sr_nos.add(sr_no)
    
    # Add summary rows - extract from PDF
    # Summary can be on any page, so search all pages
    total_amount = ""
    total_tax = ""
    grand_total = ""
    
    with pdfplumber.open(pdf_path) as pdf:
        # Search all pages for summary
        for page_num, page in enumerate(pdf.pages):
            page_text = page.extract_text() or ""
            
            # Look for summary keywords on this page
            if 'Total Amount' in page_text or 'Grand Total' in page_text:
                # Extract summary values
                amt_match = re.search(r'Total\s+Amount\s*\(INR\)\s*:?\s*([\d,.]+)', page_text, re.IGNORECASE)
                if amt_match:
                    total_amount = amt_match.group(1)
                
                tax_match = re.search(r'Total\s+Tax\s*\(INR\)\s*:?\s*([\d,.]+)', page_text, re.IGNORECASE)
                if tax_match:
                    total_tax = tax_match.group(1)
                
                grand_match = re.search(r'Grand\s+Total\s*\(INR\)\s*:?\s*([\d,.]+)', page_text, re.IGNORECASE)
                if grand_match:
                    grand_total = grand_match.group(1)
                
                # If we found at least one value, stop searching
                if total_amount or total_tax or grand_total:
                    break
    
    # Add summary rows
    all_rows.append(["", "", "", "", "", "", "", "", "", ""])
    all_rows.append(["Total Base Value", total_amount, "", "", "", "", "", "", "", ""])
    all_rows.append(["Total Tax", total_tax, "", "", "", "", "", "", "", ""])
    all_rows.append(["Grand Total", grand_total, "", "", "", "", "", "", "", ""])
    
    # Create DataFrame
    df = pd.DataFrame(all_rows, columns=[
        "Sr #",
        "EAN",
        "Item Code",
        "Product Name",
        "HSN Code",
        "Quantity",
        "MRP",
        "Base Rate",
        "GST %",
        "Total"
    ])
    
    # Write to Excel with headers
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Header info rows (Excel rows 1-8)
        pd.DataFrame([
            ["Party Name", "Scootsy"],
            ["PO No", header_data["PO No"]],
            ["PO Date", header_data["PO Date"]],
            ["PO Expiry Date", header_data["PO Expiry Date"]],
            ["Shipping Address", header_data["Shipping Address"]],
            ["GST No", header_data["GST No"]],
            ["", ""],
            ["", ""],
        ]).to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=0)
        
        # Empty rows 8-9
        pd.DataFrame([[""], [""]]).to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=7)
        
        # Table headers at row 10 (0-indexed row 9)
        pd.DataFrame([df.columns.tolist()]).to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=9)
        
        # Data starting at row 11 (0-indexed row 10)
        df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=10)
    
    return output_path


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 2:
        convert_pdf_to_excel(sys.argv[1], sys.argv[2])
        print(f"✓ Converted {sys.argv[1]} to {sys.argv[2]}")
    else:
        print("Usage: python scootsy.py <input.pdf> <output.xlsx>")
