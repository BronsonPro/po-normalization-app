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
        
        # Extract table data
        for page in pdf.pages:
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
            
            for table in tables:
                for row in table:
                    if not row or len(row) < 10:
                        continue
                    
                    # Check if it's a data row (first column is a number)
                    first_cell = str(row[0] or "").strip()
                    if not first_cell or not first_cell.isdigit():
                        continue
                    
                    # Extract data from correct columns with defensive length checks
                    # Column mapping from PDF:
                    # 0=Sr, 1=Item Code, 2=Product, 3=HSN, 4=Qty, 5=MRP, 6=Base Cost, 7=Taxable Value, 13=IGST Rate, 14=IGST Amt, 18=Total
                    
                    sr_no = first_cell
                    item_code = str(row[1] or "").strip() if len(row) > 1 else ""
                    product_name = str(row[2] or "").strip().replace('\n', ' ') if len(row) > 2 else ""
                    hsn = str(row[3] or "").strip() if len(row) > 3 else ""
                    qty = str(row[4] or "").strip() if len(row) > 4 else ""
                    mrp = str(row[5] or "").strip() if len(row) > 5 else ""
                    base_rate = str(row[6] or "").strip() if len(row) > 6 else ""
                    gst_rate = str(row[13] or "").strip() if len(row) > 13 else ""  # IGST Rate column
                    total = str(row[18] or "").strip() if len(row) > 18 else ""  # Total column (may not exist in some formats)
                    
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
    
    # Add summary rows - extract from PDF if possible
    summary_text = first_page_text
    
    total_amount = ""
    total_tax = ""
    grand_total = ""
    
    # Extract summary values
    amt_match = re.search(r'Total Amount.*?\(INR\)\s+([\d,.]+)', summary_text)
    if amt_match:
        total_amount = amt_match.group(1)
    
    tax_match = re.search(r'Total Tax.*?\(INR\)\s+([\d,.]+)', summary_text)
    if tax_match:
        total_tax = tax_match.group(1)
    
    grand_match = re.search(r'Grand Total.*?\(INR\)\s+([\d,.]+)', summary_text)
    if grand_match:
        grand_total = grand_match.group(1)
    
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
