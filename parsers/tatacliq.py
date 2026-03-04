import pdfplumber
import pandas as pd
import re


# ------------------ HEADER EXTRACTION ------------------

def extract_po_header(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

    def find(pattern, src=text):
        m = re.search(pattern, src, re.DOTALL | re.IGNORECASE)
        return m.group(1).strip() if m else ""

    buyer_name = "Tata UniStore Limited"
    po_no = find(r"Purchase Order\s*:\s*([0-9]+)")
    po_date = find(r"PO Date\s*:\s*([0-9.]+)")
    po_expiry = find(r"Shipment Date\s*:\s*([0-9.]+)")
    
    # Get all GST numbers - last one is shipping address GST
    gst_matches = re.findall(r"GST No[:\s]+([A-Z0-9]+)", text)
    gst_no = gst_matches[-1] if gst_matches else ""
    
    # Extract Shipping Address - TataCliq format
    # The shipping section appears after "Shipping Address:" and before "GST No:"
    shipping_address = ""
    
    # Method 1: Try to find shipping address block
    ship_match = re.search(r"Shipping Address:\s*(.*?)\s*GST No", text, re.DOTALL)
    if ship_match:
        addr_text = ship_match.group(1)
        # Remove any PAN/TIN references
        addr_text = re.sub(r"PAN No:.*", "", addr_text)
        addr_text = re.sub(r"TIN No:.*", "", addr_text)
        # Clean and join
        shipping_address = " ".join(addr_text.split())
    
    # Method 2: If method 1 didn't work, extract line by line
    if not shipping_address or len(shipping_address) < 10:
        lines = text.split('\n')
        in_shipping_section = False
        shipping_lines = []
        
        for i, line in enumerate(lines):
            if 'Shipping Address:' in line:
                in_shipping_section = True
                # Sometimes the first line of address is on same line
                after_label = line.split('Shipping Address:')[-1].strip()
                if after_label:
                    shipping_lines.append(after_label)
                continue
            
            if in_shipping_section:
                # Stop at GST No, PAN No, or next section
                if any(marker in line for marker in ['GST No:', 'PAN No:', 'TIN No:', 'Page ']):
                    break
                
                # Add non-empty lines
                if line.strip():
                    shipping_lines.append(line.strip())
        
        if shipping_lines:
            shipping_address = " ".join(shipping_lines)

    data = {
        "Party Name": buyer_name,
        "PO No": po_no,
        "PO Date": po_date,
        "PO Expiry Date": po_expiry,
        "Shipping Address": shipping_address,
        "GST #": gst_no,
    }

    return data

# ------------------ LINE ITEMS EXTRACTION ------------------

def extract_line_items_and_text_totals(pdf_path):
    all_rows = []
    total_text_rows = []

    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 2:
            raise Exception("PDF does not have enough pages")
        
        page = pdf.pages[1]
        text = page.extract_text()
        
        if not text:
            raise Exception("Could not extract text from page 2")
        
        lines = text.split('\n')
        
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # Check if this line starts with article code (80100...)
            if re.match(r'^80100\d+', line):
                parts = line.split()
                
                if len(parts) < 10:
                    i += 1
                    continue
                
                # Article Code (first part)
                article_code = parts[0]
                
                # Extract EAN (13 digit number, but not the article code)
                ean = ""
                for part in parts[1:]:  # Skip first part (article code)
                    if (len(part) == 13 or len(part) == 12) and part.isdigit():
                        ean = part
                        break 
                # Extract HSN (8 digit number)
                hsn = ""
                for part in parts[1:]:
                    if len(part) == 8 and part.isdigit():
                        hsn = part
                        break
                
                if not ean:
                    i += 1
                    continue
                
                # Get product name - between article code and EAN
                product_parts = []
                for j in range(1, len(parts)):
                    if parts[j] == ean:
                        break
                    # Skip HSN and other numeric codes
                    if not (parts[j].isdigit() and len(parts[j]) >= 8):
                        product_parts.append(parts[j])
                
                # Add next line for product name continuation
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if not next_line.startswith('80100'):
                        next_parts = next_line.split()
                        for part in next_parts[:3]:
                            if part.upper() not in ['TRANSPARENT', 'WHITE', 'BLACK', 'GREY', 'NA', 'ENT']:
                                if not ('GRAM' in part.upper() or part.upper().endswith('G') or part.upper().endswith('AM')):
                                    product_parts.append(part)
                
                product_name = " ".join(product_parts)
                
                # Now extract the specific fields we need
                # Line format after product name and codes:
                # ... Colour Size QTY_Net UoM Unit_Cost Taxable CGST_Rate CGST_Amt SGST_Rate SGST_Amt IGST_Rate IGST_Amt Total
                
                # Find position after HSN
                hsn_index = -1
                for j, part in enumerate(parts):
                    if part == hsn:
                        hsn_index = j
                        break
                
                if hsn_index == -1:
                    i += 1
                    continue
                
                # After HSN, we have: Colour, Size, QTY, UoM, Unit Cost, ...
                # Skip to numeric values after HSN
                remaining_parts = parts[hsn_index + 1:]
                
                # Filter only numeric parts (skip Colour, Size, UoM like "PC")
                numeric_parts = []
                for part in remaining_parts:
                    clean = part.replace(',', '').replace('.', '')
                    if clean.isdigit() or '.' in part:
                        try:
                            float(part.replace(',', ''))
                            numeric_parts.append(part)
                        except:
                            pass
                
                # Now numeric_parts should be: [QTY, Unit_Cost, Taxable, CGST_Rate, CGST_Amt, SGST_Rate, SGST_Amt, IGST_Rate, IGST_Amt, Total]
                # That's 10 values minimum
                
                if len(numeric_parts) >= 10:
                    qty = numeric_parts[0]           # QTY Net
                    unit_cost = numeric_parts[1]     # Unit Cost (INR)
                    taxable = numeric_parts[2]       # Taxable Amount
                    cgst_rate = numeric_parts[3]     # CGST Rate
                    cgst_amt = numeric_parts[4]      # CGST Amount
                    sgst_rate = numeric_parts[5]     # SGST Rate
                    sgst_amt = numeric_parts[6]      # SGST Amount
                    igst_rate = numeric_parts[7]     # IGST Rate
                    igst_amt = numeric_parts[8]      # IGST Amount
                    total = numeric_parts[9]         # Total Gross Cost
                    
                    # Calculate GST %
                    try:
                        cgst = float(cgst_rate)
                        sgst = float(sgst_rate)
                        igst = float(igst_rate)
                        if igst > 0:
                            gst_pct = igst
                        else:
                            gst_pct = cgst + sgst
                    except:
                        gst_pct = 18.0
                    
                    all_rows.append({
                        "EAN": ean,
                        "Product Name": product_name,
                        "HSN Code": hsn,
                        "Quantity": qty,
                        "MRP": "",
                        "Base Rate": unit_cost,
                        "GST %": gst_pct,
                        "Total": total,
                    })
            
            # Check for total row
            if "Total" in line and "PC" in line and not line.startswith('80100'):
                total_text_rows.append(line)
            
            i += 1

    if not all_rows:
        raise Exception("No line items detected.")

    return pd.DataFrame(all_rows), total_text_rows


# ------------------ CLEAN ------------------

def clean_and_validate_line_items(df):
    df = df.copy()

    def extract_number(x):
        if pd.isna(x) or x == "":
            return 0.00
        s = str(x).replace(",", "")
        m = re.search(r"\d+(\.\d+)?", s)
        return round(float(m.group()) if m else 0.00, 2)

    def extract_int(x):
        if pd.isna(x):
            return 0
        m = re.search(r"\d+", str(x))
        return int(m.group()) if m else 0

    df["Product Name"] = df["Product Name"].astype(str).str.replace(r"\s{2,}", " ", regex=True).str.strip()
    df["Quantity"] = df["Quantity"].apply(extract_int)
    df["MRP"] = df["MRP"].apply(extract_number)
    df["Base Rate"] = df["Base Rate"].apply(extract_number)
    df["Total"] = df["Total"].apply(extract_number)
    
    try:
        df["GST %"] = df["GST %"].astype(float).round(2)
    except:
        df["GST %"] = 18.0

    return df


# ------------------ SUMMARY ------------------

def extract_summary_from_text(total_text_list):
    if not total_text_list:
        return {"Total Base Value": "0.00", "Total Tax": "0.00", "Grand Total": "0.00"}

    text_blob = " ".join(total_text_list)
    
    # Extract all numbers with decimals
    numbers = re.findall(r'[\d,]+\.\d{2}', text_blob)
    
    if len(numbers) >= 4:
        # Total line format: Total QTY PC Taxable_Total Tax1 Tax2 Tax3 Grand_Total
        base_value = numbers[0]
        grand_total = numbers[-1]
        
        # Tax is sum of CGST and SGST
        if len(numbers) >= 3:
            try:
                tax1 = float(numbers[1].replace(',', ''))
                tax2 = float(numbers[2].replace(',', ''))
                total_tax = f"{(tax1 + tax2):.2f}"
            except:
                total_tax = "0.00"
        else:
            total_tax = "0.00"
        
        return {
            "Total Base Value": base_value,
            "Total Tax": total_tax,
            "Grand Total": grand_total,
        }
    
    return {"Total Base Value": "0.00", "Total Tax": "0.00", "Grand Total": "0.00"}


# ================== PUBLIC FUNCTION ==================

def convert_pdf_to_excel(pdf_file_path, output_excel_path):
    header_data = extract_po_header(pdf_file_path)
    items_df, total_text_list = extract_line_items_and_text_totals(pdf_file_path)
    items_df = clean_and_validate_line_items(items_df)
    items_df.insert(0, "Sr #", range(1, len(items_df) + 1))
    summary_data = extract_summary_from_text(total_text_list)

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        row = 0
        header_df = pd.DataFrame({
            "Field": list(header_data.keys()),
            "Value": list(header_data.values()),
        })
        header_df.to_excel(writer, index=False, startrow=row, header=False)
        row += len(header_df) + 2

        items_df.to_excel(writer, index=False, startrow=row)
        row += len(items_df) + 2

        summary_df = pd.DataFrame({
            "Field": list(summary_data.keys()),
            "Value": list(summary_data.values()),
        })
        summary_df.to_excel(writer, index=False, startrow=row, header=False)