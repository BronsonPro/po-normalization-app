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

    # Extract buyer name
    buyer_name = "Blink Commerce Private Limited"
    
    # Extract PO Number
    po_no = find(r"P\.O\.\s*Number\s*:?\s*([0-9]+)")
    
    # Extract PO Date
    po_date = find(r"Date\s*:?\s*([A-Za-z]+\.\s*\d{1,2},\s*\d{4})")
    
    # Extract PO Expiry Date
    po_expiry = find(r"PO expiry date\s*:?\s*([A-Za-z]+\.\s*\d{1,2},\s*\d{4})")
    
    # Extract GST Number (there are two - we want the "Delivered To" one which is 30...)
    gst_matches = re.findall(r"GST No\.\s*:\s*([A-Z0-9]+)", text)
    gst_no = gst_matches[-1] if gst_matches else ""  # Take the last one (Delivered To)
    
    # Extract Shipping Address (after "Delivered To", before the product table)
    ship_match = re.search(r"To\s+(.*?)(?=#\s+Item\s+HSN|#\s+Item\s+Code)", text, re.DOTALL)
    if ship_match:
        # Extract just the address lines, remove GST No and Reference
        addr_text = ship_match.group(1)
        # Remove GST No line
        addr_text = re.sub(r"GST No\..*", "", addr_text)
        # Remove Reference line  
        addr_text = re.sub(r"Reference\s*:.*", "", addr_text)
        shipping_address = " ".join(addr_text.split())
    else:
        shipping_address = ""

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
        for page in pdf.pages:
            
            table = page.extract_table({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5,
                "snap_tolerance": 5,
                "join_tolerance": 5,
            })

            if not table:
                continue

            for row in table:
                r = [(c.strip() if c else "") for c in row]
                
                # Skip header rows
                row_text = " ".join(r).lower()
                if "item code" in row_text or "#" == r[0]:
                    continue
                
                # Check if this is a total row
                is_total_text = (
                    "total quantity" in row_text
                    or "total amount" in row_text
                    or "net amount" in row_text
                )
                
                # Extract product rows (rows that start with a number)
                if r[0] and r[0].strip().isdigit():
                    
                    # BlinkIt columns: #, Item Code, HSN Code, Product UPC, Product Description, 
                    # Basic Cost Price, IGST %, CESS %, ADD T. CESS, Tax Amt, Landing Rate, 
                    # Qty, MRP, Margin %, Total Amt
                    
                    item_code = r[1].replace("\n", "") if r[1] else ""
                    hsn = r[2].replace("\n", "") if r[2] else ""
                    upc = r[3].replace("\n", "") if r[3] else ""  # This is EAN
                    product_name = r[4].replace("\n", " ") if r[4] else ""
                    basic_cost = r[5] if r[5] else ""
                    igst_rate = r[6].replace("\n", "") if r[6] else ""
                    landing_rate = r[10].replace("\n", "") if r[10] else ""
                    qty = r[11] if r[11] else ""
                    mrp = r[12] if r[12] else ""
                    total = r[14] if len(r) > 14 and r[14] else ""
                    
                    # Calculate GST % from IGST
                    gst_pct = igst_rate
                    
                    out = {
                        "EAN": upc,
                        "Product Name": product_name,
                        "HSN Code": hsn,
                        "Quantity": qty,
                        "MRP": mrp,
                        "Base Rate": landing_rate,
                        "GST %": gst_pct,
                        "Total": total,
                    }
                    
                    all_rows.append(out)
                
                if is_total_text:
                    # Store total rows for summary
                    total_text_rows.append(" ".join(r))

    if not all_rows:
        raise Exception("No line items detected.")

    line_df = pd.DataFrame(all_rows)

    return line_df.reset_index(drop=True), total_text_rows


# ------------------ CLEAN ------------------

def clean_and_validate_line_items(df):
    df = df.copy()

    money_cols = ["MRP", "Base Rate", "Total"]

    def extract_number(x):
        if pd.isna(x):
            return 0.00
        s = str(x).replace(",", "")
        m = re.search(r"\d+(\.\d+)?", s)
        return round(float(m.group()) if m else 0.00, 2)

    def extract_int(x):
        if pd.isna(x):
            return 0
        m = re.search(r"\d+", str(x))
        return int(m.group()) if m else 0

    def extract_rate(x):
        if pd.isna(x):
            return 0.00
        m = re.search(r"\d+(\.\d+)?", str(x))
        return round(float(m.group()) if m else 0.00, 2)

    # Clean product name
    df["Product Name"] = (
        df["Product Name"]
        .astype(str)
        .str.replace(r"\s{2,}", " ", regex=True)
        .str.strip()
    )

    df["Quantity"] = df["Quantity"].apply(extract_int)

    for c in money_cols:
        df[c] = df[c].apply(extract_number)

    df["GST %"] = df["GST %"].apply(extract_rate)

    return df


# ------------------ SUMMARY FROM TEXT ------------------

def extract_summary_from_text(total_text_list):
    if not total_text_list:
        return {}

    text_blob = " ".join(total_text_list)

    def find(pattern):
        m = re.search(pattern, text_blob, re.IGNORECASE)
        return m.group(1).strip() if m else "0.00"

    def fmt2_text(x):
        try:
            val = float(str(x).replace(",", ""))
            return f"{val:.2f}"
        except:
            return "0.00"

    # BlinkIt shows: Total Amount, Cart Discount, Net amount
    # We'll use Total Amount as Total Base Value and Net amount as Grand Total
    
    summary = {
        "Total Base Value": fmt2_text(find(r"Total Amount\s*([0-9,.]+)")),
        "Total Tax": "0.00",  # Not shown separately in BlinkIt PO
        "Grand Total": fmt2_text(find(r"Net amount\s*([0-9,.]+)")),
    }

    return summary


# ================== PUBLIC FUNCTION FOR GUI ==================

def convert_pdf_to_excel(pdf_file_path, output_excel_path):

    header_data = extract_po_header(pdf_file_path)

    items_df, total_text_list = extract_line_items_and_text_totals(pdf_file_path)

    items_df = clean_and_validate_line_items(items_df)

    # -------- ADD SR # --------
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