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
    
    # Try to load master file for EAN and Product Name mapping
    master_ean_map = {}
    master_product_map = {}
    try:
        import os
        import streamlit as st
        master_dir = os.path.dirname(pdf_path)
        master_path = os.path.join(master_dir, "Scootsy Master.xlsx")
        if not os.path.exists(master_path):
            master_path = os.path.join(master_dir, "Masterfile_Scootsy.xlsx")
        if not os.path.exists(master_path):
            master_path = os.path.join(master_dir, "Scootsy_master.xlsx")
        
        if os.path.exists(master_path):
            master_df = pd.read_excel(master_path)
            
            # Try both column name possibilities for EAN
            if 'Brand SKU Code' in master_df.columns and 'Item Code' in master_df.columns:
                for _, row in master_df.iterrows():
                    item_code = str(int(float(row['Item Code']))) if pd.notna(row['Item Code']) else None
                    ean = str(int(float(row['Brand SKU Code']))) if pd.notna(row['Brand SKU Code']) else None
                    # Product name is in 'SKU Name' column
                    product_name = str(row['SKU Name']) if pd.notna(row.get('SKU Name')) else None
                    if item_code:
                        if ean:
                            master_ean_map[item_code] = ean
                        if product_name:
                            master_product_map[item_code] = product_name
            elif 'EAN' in master_df.columns and 'Item Code' in master_df.columns:
                for _, row in master_df.iterrows():
                    item_code = str(int(float(row['Item Code']))) if pd.notna(row['Item Code']) else None
                    ean = str(int(float(row['EAN']))) if pd.notna(row['EAN']) else None
                    # Product name is in 'SKU Name' column
                    product_name = str(row['SKU Name']) if pd.notna(row.get('SKU Name')) else None
                    if item_code:
                        if ean:
                            master_ean_map[item_code] = ean
                        if product_name:
                            master_product_map[item_code] = product_name
    except Exception as e:
        # If master not found, EAN and Product Name will be empty/from PO
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
            
            # Extract shipping address from first page table only
            if page_num == 0:
                for row in tables[0]:
                    if row and len(row) > 9 and row[9] and 'PJTJ' in str(row[9]) and not header_data["Shipping Address"]:
                        addr_lines = str(row[9]).split('\n')
                        # Take lines before Contact and GSTIN
                        clean_lines = []
                        for line in addr_lines:
                            if 'GSTIN' in line or 'GST' in line:
                                # Extract GSTIN number - use proper GSTIN format pattern
                                gstin_match = re.search(r'(?:GSTIN|GST\s*No|GST)[:\s-]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})', line, re.IGNORECASE)
                                if gstin_match:
                                    header_data["GST No"] = gstin_match.group(1)
                                break
                            if 'Contact' in line:
                                break
                            clean_lines.append(line.strip())
                        # Store full address - don't truncate
                        header_data["Shipping Address"] = ', '.join(clean_lines)
        
        # Fallback: Try to extract GSTIN from shipping address section of first page text
        if not header_data["GST No"]:
            # Only search in shipping address section, NOT vendor section
            # Try multiple end markers to capture the full shipping block
            ship_section = None
            for end_marker in [r'Vendor\s+Details', r'Vendor', r'Terms\s+and\s+Conditions', r'PO\s+No']:
                ship_match = re.search(r'Shipping Address[:\s]+(.*?)' + end_marker, first_page_text, re.DOTALL | re.IGNORECASE)
                if ship_match:
                    ship_section = ship_match.group(1)
                    break
            
            # If still not found, try getting everything after "Shipping Address" for next 500 chars
            if not ship_section:
                ship_match = re.search(r'Shipping Address[:\s]+(.{1,500})', first_page_text, re.DOTALL | re.IGNORECASE)
                if ship_match:
                    ship_section = ship_match.group(1)
            
            if ship_section:
                # Look for GSTIN in the shipping section
                gstin_match = re.search(r'(?:GSTIN|GST\s*No|GST)[:\s-]*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})', ship_section, re.IGNORECASE)
                if gstin_match:
                    header_data["GST No"] = gstin_match.group(1)
        
        # Fallback: Extract shipping address from first page text if not found in table
        if not header_data["Shipping Address"]:
            # Look for "Shipping Address" or "Ship To:" or "Consignee:" section
            ship_match = re.search(r'Shipping Address\s+(.*?)(?:GSTIN|Contact|GST No|PO\s+No)', first_page_text, re.DOTALL | re.IGNORECASE)
            if not ship_match:
                ship_match = re.search(r'(?:Ship To|Consignee)[:.\s]+(.*?)(?:GSTIN|Contact|PO)', first_page_text, re.DOTALL | re.IGNORECASE)
            
            if ship_match:
                addr_text = ship_match.group(1).strip()
                # Store full address - don't truncate
                # Later processing in app.py will extract pincode from this
                header_data["Shipping Address"] = addr_text
        
        # Extract table data from all pages (second pass for data rows)
        all_sr_numbers_found = []  # Track all serial numbers we find
        
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
            
            # Track which serial numbers we've processed
            processed_sr_nos = set()
            
            for table in tables:
                for row_idx, row in enumerate(table):
                    first_cell = str(row[0] or "").strip()
                    
                    if not row or len(row) < 8:  # Lowered from 10 to 8 to catch last rows
                        continue
                    
                    # Check if it's a data row (first column is a number)
                    if not first_cell or not first_cell.isdigit():
                        continue
                    
                    # Skip if already processed (avoid duplicates)
                    if first_cell in processed_sr_nos:
                        continue
                    processed_sr_nos.add(first_cell)
                    all_sr_numbers_found.append((first_cell, page_num + 1))  # Track with page number
                    
                    # Detect table format and check if HSN is missing
                    row_len = len(row)
                    
                    # Check if this is a page without HSN:
                    # If col[4] is an 8-digit number (HSN pattern), then col[3] should be HSN but might be empty
                    # In that case, the actual HSN is in col[4] and we're missing the HSN column
                    col3_val = str(row[3] or "").strip() if row_len > 3 else ""
                    col4_val = str(row[4] or "").strip() if row_len > 4 else ""
                    
                    # If col[3] is empty/short and col[4] looks like HSN (8 digits), HSN column is missing
                    has_no_hsn = (row_len == 19 and 
                                 len(col3_val) < 5 and 
                                 len(col4_val) == 8 and 
                                 col4_val.isdigit())
                    
                    # Extract data from correct columns
                    sr_no = first_cell
                    item_code = str(row[1] or "").strip() if row_len > 1 else ""
                    product_name = str(row[2] or "").strip().replace('\n', ' ') if row_len > 2 else ""
                    
                    if has_no_hsn:
                        # 19-column format WITHOUT HSN - everything shifts left by 1
                        hsn = col4_val  # HSN is actually in col[4]
                        qty = str(row[5] or "").strip() if row_len > 5 else ""  # Qty shifted to col[5]
                        mrp = str(row[6] or "").strip() if row_len > 6 else ""  # MRP shifted to col[6]
                        base_rate = str(row[7] or "").strip() if row_len > 7 else ""  # Base shifted to col[7]
                    else:
                        # Normal formats with HSN in correct position
                        hsn = col3_val
                        qty = col4_val
                        
                        # Handle MRP and Base Rate - they might be merged in column 5 for 17-column format
                        if row_len == 17:
                            col5 = str(row[5] or "").strip() if row_len > 5 else ""
                            col6 = str(row[6] or "").strip() if row_len > 6 else ""
                            
                            # Check if col5 contains two space-separated numbers (MRP + Base Rate)
                            col5_parts = col5.split()
                            if len(col5_parts) >= 2:
                                mrp = col5_parts[0]
                                base_rate = col5_parts[1]
                            else:
                                mrp = col5
                                base_rate = col6
                        else:
                            # Normal format
                            mrp = str(row[5] or "").strip() if row_len > 5 else ""
                            base_rate = str(row[6] or "").strip() if row_len > 6 else ""
                    
                    # GST Rate - indices depend on table format
                    if row_len == 19:
                        # 19-column format (page 1): None at index 9, everything shifted by 1
                        cgst_rate = str(row[8] or "").strip() if row_len > 8 else ""
                        sgst_rate = str(row[11] or "").strip() if row_len > 11 else ""
                        igst_rate = str(row[13] or "").strip() if row_len > 13 else ""
                        total = str(row[18] or "").strip() if row_len > 18 else ""
                    elif row_len == 18:
                        # 18-column format (page 2): No extra None column
                        cgst_rate = str(row[8] or "").strip() if row_len > 8 else ""
                        sgst_rate = str(row[10] or "").strip() if row_len > 10 else ""
                        igst_rate = str(row[12] or "").strip() if row_len > 12 else ""
                        total = str(row[17] or "").strip() if row_len > 17 else ""
                    elif row_len == 17:
                        # 17-column format (pages 3, 5): Columns merged
                        # IGST might be in column 11 or 12 with merged values
                        col11 = str(row[11] or "").strip() if row_len > 11 else ""
                        col12 = str(row[12] or "").strip() if row_len > 12 else ""
                        
                        # Check col12 first (might have "0 18.00 153.76" format)
                        col12_parts = col12.split()
                        if len(col12_parts) >= 2:
                            # Format: "0 18.00 ..." -> IGST rate is second value
                            igst_rate = col12_parts[1] if len(col12_parts) > 1 else ""
                        else:
                            igst_rate = col11
                        
                        cgst_rate = "0"  # Assume IGST format
                        sgst_rate = "0"
                        total = str(row[16] or "").strip() if row_len > 16 else ""
                    else:
                        # Fallback for other formats
                        cgst_rate = str(row[8] or "").strip() if row_len > 8 else ""
                        sgst_rate = str(row[10] or "").strip() if row_len > 10 else ""
                        igst_rate = str(row[12] or "").strip() if row_len > 12 else ""
                        total = str(row[-1] or "").strip()  # Last column
                    
                    # Use IGST if present, otherwise CGST+SGST
                    gst_rate = ""
                    try:
                        # Clean rate strings - extract just the numeric value
                        # Handle formats like '9.00', '0 9.00', '0.00', etc.
                        def clean_rate(rate_str):
                            if not rate_str:
                                return 0
                            # Split on space and take the last numeric part
                            parts = rate_str.strip().split()
                            for part in reversed(parts):  # Check from right to left
                                try:
                                    val = float(part)
                                    if val > 0:  # Return first positive value found
                                        return val
                                except:
                                    continue
                            # If no positive value found, try to convert the whole string
                            try:
                                return float(rate_str.strip())
                            except:
                                return 0
                        
                        igst_val = clean_rate(igst_rate)
                        cgst_val = clean_rate(cgst_rate)
                        sgst_val = clean_rate(sgst_rate)
                        
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
                    
                    # Map Item Code to EAN and Product Name from master if available
                    item_code_clean = str(int(float(item_code))) if item_code and item_code.replace('.','').replace('-','').isdigit() else ""
                    ean = master_ean_map.get(item_code_clean, "")  # Get EAN from master map
                    master_product_name = master_product_map.get(item_code_clean, "")  # Get Product Name from master
                    
                    # Use master product name if available, otherwise use PO product name
                    final_product_name = master_product_name if master_product_name else product_name
                    
                    all_rows.append([
                        sr_no,
                        ean,  # EAN from master file lookup
                        item_code,
                        final_product_name,  # Product name from master (preferred) or PO
                        hsn,
                        qty,
                        mrp,
                        base_rate,
                        gst_rate,
                        total
                    ])
            
            # FALLBACK: Extract rows from text that weren't captured in tables
            # This handles rows at page tops without borders AND fixes truncated product names
            # Use layout=True to preserve positioning and multi-line text
            page_text = page.extract_text(layout=True) or ""
            text_lines = page_text.split('\n')
            
            for line_idx, line in enumerate(text_lines):
                # Look for lines that start with a number followed by item code pattern
                # Pattern: "5 255618 Beautiliss Professional Classic Eyelash Curler ... 82142090 12 200.00 ..."
                # Allow leading whitespace
                match = re.match(r'^\s*(\d+)\s+(\d{6})\s+(.+)', line)
                if match:
                    sr_no = match.group(1)
                    item_code = match.group(2)
                    rest = match.group(3)
                    
                    # For page 2, product name might be on multiple lines
                    # Collect next few lines until we hit the HSN (8-digit number)
                    if page_num > 0:  # Page 2+
                        full_rest = rest
                        for next_line in text_lines[line_idx + 1:line_idx + 10]:  # Check next 10 lines
                            # Stop if we hit a line with HSN or another serial number
                            if re.search(r'\d{8}', next_line) or re.match(r'^\d+\s+\d{6}', next_line):
                                break
                            # Add this line to the product name
                            full_rest += ' ' + next_line.strip()
                        rest = full_rest
                    
                    # Check if this row was already processed
                    # Product name from master will be used regardless, so no need to update
                    row_found = False
                    for idx, existing_row in enumerate(all_rows):
                        if existing_row[0] == sr_no:  # Match by serial number
                            row_found = True
                            break
                    
                    # If row wasn't found in table extraction, add it now
                    if not row_found:  # Removed the processed_sr_nos check - only check if row is actually in all_rows
                        # Try to parse the entire line as space-separated values
                        parts = line.split()
                        if len(parts) < 8:  # Lowered from 10 to 8 to catch last rows
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
                        if len(remaining) < 6:  # Lowered from 8 to 6 to catch shorter rows
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
                        
                        # Map Item Code to EAN and Product Name from master
                        item_code_clean = str(int(float(item_code))) if item_code and item_code.replace('.','').replace('-','').isdigit() else ""
                        ean = master_ean_map.get(item_code_clean, "")
                        master_product_name = master_product_map.get(item_code_clean, "")
                        
                        # Use master product name if available
                        final_product_name = master_product_name if master_product_name else product_name
                        
                        all_rows.append([
                            sr_no,
                            ean,
                            item_code,
                            final_product_name,
                            hsn,
                            qty,
                            mrp,
                            base_rate,
                            gst_rate,
                            total
                        ])
                        
                        processed_sr_nos.add(sr_no)
                        all_sr_numbers_found.append((sr_no, page_num + 1))  # Track in summary too!
    
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
