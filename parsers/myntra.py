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
        # PO Number
        if "Purchase Order Number" in line or "PO Number" in line:
            m = re.search(r':\s*(\S+)', line)
            if m:
                po_no = m.group(1).strip()
        
        # PO Date
        if "PO Date" in line:
            m = re.search(r':\s*([\d/\-\.]+)', line)
            if m:
                po_date = m.group(1).strip()
        
        # Expiry Date
        if "Expiry Date" in line or "Valid" in line:
            m = re.search(r':\s*([\d/\-\.]+)', line)
            if m:
                po_expiry = m.group(1).strip()
        
        # GST Number
        if "GSTIN" in line or "GST" in line:
            m = re.search(r'([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})', line)
            if m:
                gst_no = m.group(1).strip()
        
        # Shipping Address (pincode)
        m = re.search(r'\b(\d{6})\b', line)
        if m and not shipping_address:
            shipping_address = m.group(1)

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
            
            for table in tables:
                
                # Find header row
                header_row_idx = None
                for idx, row in enumerate(table):
                    if row and any(cell and "SKU Code" in str(cell) for cell in row):
                        header_row_idx = idx
                        break
                
                if header_row_idx is None:
                    continue
                
                headers = table[header_row_idx]
                
                # Find column indices
                col_sku = None
                col_ean = None
                col_hsn = None
                col_product_name = None
                col_qty = None
                col_mrp = None
                col_list_price = None
                col_landed_price = None
                col_cgst_pct = None
                col_sgst_pct = None
                col_igst_pct = None
                col_total = None
                
                for idx, h in enumerate(headers):
                    if not h:
                        continue
                    h_clean = str(h).lower().replace("\n", " ").replace("  ", " ").strip()
                    
                    if "sku code" in h_clean or "sku" == h_clean:
                        col_sku = idx
                    elif "vendor article number" in h_clean or "article number" in h_clean:
                        col_ean = idx
                    elif "hsn code" in h_clean or "hsn" == h_clean:
                        col_hsn = idx
                    elif "vendor article name" in h_clean or "product name" in h_clean or "article name" in h_clean:
                        col_product_name = idx
                    elif h_clean == "qty" or "quantity" in h_clean:
                        col_qty = idx
                    elif h_clean == "mrp":
                        col_mrp = idx
                    elif "list price" in h_clean:
                        col_list_price = idx
                    elif "landed price" in h_clean:
                        col_landed_price = idx
                    elif "cgst tax percent" in h_clean or h_clean == "cgst%":
                        col_cgst_pct = idx
                    elif "sgst tax percent" in h_clean or h_clean == "sgst%":
                        col_sgst_pct = idx
                    elif "igst tax percent" in h_clean or h_clean == "igst%" or "igst" in h_clean:
                        col_igst_pct = idx
                    elif "total" in h_clean and ("plus" in h_clean or "tax" in h_clean):
                        col_total = idx
                
                # Validate we have minimum required columns
                if col_sku is None or col_ean is None or col_qty is None:
                    continue
                
                # Process data rows
                for row_idx in range(header_row_idx + 1, len(table)):
                    row = table[row_idx]
                    
                    if not row or not row[col_sku]:
                        continue
                    
                    sku = str(row[col_sku] or "").strip()
                    ean = str(row[col_ean] or "").strip()
                    
                    # Skip empty or header-like rows
                    if not sku or "SKU" in sku or len(sku) < 3:
                        continue
                    
                    if not ean or len(ean) < 3:
                        continue
                    
                    # Extract values
                    hsn = str(row[col_hsn] or "").strip() if col_hsn is not None else ""
                    product_name = str(row[col_product_name] or "").strip() if col_product_name is not None else ""
                    qty_text = str(row[col_qty] or "0").strip() if col_qty is not None else "0"
                    mrp_text = str(row[col_mrp] or "0").strip() if col_mrp is not None else "0"
                    
                    # Use LIST PRICE for Base Rate (not Landed Price!)
                    list_price_text = str(row[col_list_price] or "0").strip() if col_list_price is not None else "0"
                    
                    total_text = str(row[col_total] or "0").strip() if col_total is not None else "0"
                    
                    # Clean numeric values
                    qty_text = re.sub(r'[^\d.]', '', qty_text)
                    mrp_text = re.sub(r'[^\d.]', '', mrp_text)
                    list_price_text = re.sub(r'[^\d.]', '', list_price_text)
                    total_text = re.sub(r'[^\d.]', '', total_text)
                    
                    # Extract GST %
                    gst_pct = 0.0
                    
                    # Check for IGST format
                    if col_igst_pct is not None:
                        igst_text = str(row[col_igst_pct] or "").strip()
                        igst_text = re.sub(r'[^\d.]', '', igst_text)
                        if igst_text and igst_text != "0":
                            try:
                                gst_pct = float(igst_text)
                            except:
                                pass
                    
                    # Check for CGST+SGST format if IGST is 0 or not found
                    if gst_pct == 0.0 and col_cgst_pct is not None and col_sgst_pct is not None:
                        cgst_text = str(row[col_cgst_pct] or "").strip()
                        sgst_text = str(row[col_sgst_pct] or "").strip()
                        cgst_text = re.sub(r'[^\d.]', '', cgst_text)
                        sgst_text = re.sub(r'[^\d.]', '', sgst_text)
                        
                        try:
                            cgst_val = float(cgst_text) if cgst_text else 0.0
                            sgst_val = float(sgst_text) if sgst_text else 0.0
                            gst_pct = cgst_val + sgst_val
                        except:
                            pass
                    
                    # Convert to proper types
                    try:
                        qty_int = int(float(qty_text)) if qty_text else 0
                        mrp_float = float(mrp_text) if mrp_text else 0.0
                        list_price_float = float(list_price_text) if list_price_text else 0.0
                        total_float = float(total_text) if total_text else 0.0
                    except:
                        continue
                    
                    # Validate data
                    if qty_int <= 0 or mrp_float <= 0:
                        continue
                    
                    items.append({
                        "Sr #": len(items) + 1,
                        "EAN": ean,  # Using Vendor Article Number
                        "Product Name": product_name,
                        "HSN Code": hsn,
                        "Quantity": qty_int,
                        "MRP": mrp_float,
                        "Base Rate": list_price_float,  # Using List Price as requested
                        "GST %": gst_pct,
                        "Total": total_float,
                    })

    return pd.DataFrame(items)


# ------------------ SUMMARY EXTRACTION ------------------

def extract_summary(pdf_path):
    grand_total = 0.0

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[-1].extract_text() or ""
        
        for line in text.split("\n"):
            # Look for grand total
            if "Grand Total" in line or "Total Amount" in line:
                m = re.search(r'([\d,]+\.?\d*)', line)
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
    total_tax = round(grand_total - total_base, 2) if grand_total > 0 else 0.0

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
