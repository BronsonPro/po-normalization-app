import pdfplumber
import pandas as pd
import re


# ------------------ HEADER EXTRACTION ------------------

def extract_po_header(pdf_path):

    party_name = "D-MART"
    po_no = ""
    po_date = ""
    po_expiry = ""
    shipping_address = ""
    gst_no = ""

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    for line in text.split("\n"):

        # PO Number in title line
        m = re.search(r'PurchaseOrder\s+(\d+)', line)
        if m:
            po_no = m.group(1).strip()

        # PO Date and Validity
        m = re.search(r'PurchaseOrderDate:([\d.]+)', line)
        if m:
            po_date = m.group(1).strip()

        m = re.search(r'POValidity:[\d.]+to([\d.]+)', line)
        if m:
            po_expiry = m.group(1).strip()

        # GST - Ship To GST (buyer GST)
        m = re.search(r'GST#([A-Z0-9]{15})', line)
        if m and not gst_no:
            gst_no = m.group(1).strip()

    # Shipping address - collect all lines from ShipTo to PurchaseOrderDate
    lines = text.split("\n")
    ship_section = ""
    in_ship = False
    for line in lines:
        if "ShipTo" in line:
            in_ship = True
        if in_ship:
            ship_section += line + " "
        if "PurchaseOrderDate" in line and in_ship:
            break
    
    # Find all 6-digit pincodes
    all_pins = re.findall(r'\b(\d{6})\b', ship_section)
    from collections import Counter
    pin_counts = Counter([pin for pin in all_pins if not pin.startswith('103')])
    if pin_counts:
        shipping_address = pin_counts.most_common(1)[0][0]

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

    all_lines = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            all_lines.extend(text.split("\n"))

    items = []

    i = 0
    while i < len(all_lines):
        line = all_lines[i].strip()

        # NEW FORMAT (2026 IGST): "1 2000054628 8214 BronsonProfessional EA 120 100.00 33.90 - - - 18.00 - 40.00 4,800.00"
        # Pattern: SR EAN HSN Desc EA Qty MRP Basic - - - IGST% - Landed Total
        m_new = re.match(
            r'^(\d+)\s+'                        # 1: SR No
            r'(\d{10,13})\s+'                   # 2: EAN (10-13 digits)
            r'(\d{4})\s+'                       # 3: HSN part 1
            r'(\S+)\s+'                         # 4: Description (concatenated, no spaces)
            r'EA\s+'                            # UOM
            r'(\d+)\s+'                         # 5: Qty
            r'([\d,]+\.?\d*)\s+'               # 6: MRP
            r'([\d,]+\.?\d*)\s+'               # 7: Basic Price
            r'-\s+-\s+-\s+'                     # Skip 3 dashes
            r'([\d.]+)\s+'                      # 8: IGST%
            r'-\s+'                             # Skip 1 dash
            r'([\d,]+\.?\d*)\s+'               # 9: Landed Price
            r'([\d,]+\.?\d*)$',                # 10: Total Value
            line
        )

        # OLD FORMAT (CGST+SGST): Handle rows with AND without DP column more carefully
        # The challenge: Some rows have DP, some don't, and we need to identify CGST/SGST correctly
        # CGST and SGST are always single-digit decimals like 9.00, 12.00, 18.00
        # DP can be larger values like 20.00, 25.00, etc.
        
        m_old = None
        
        # Strategy: Match the pattern more flexibly and check which numbers are GST rates
        m_old_flexible = re.match(
            r'^(\d+)\s+'                        # 1: SR No
            r'(\d{10,13})\s+'                   # 2: EAN
            r'(\d{4,8})\s+'                     # 3: HSN part 1
            r'(.+?)\s+'                         # 4: Description part 1
            r'EA\s+'                            # UOM
            r'(\d+)\s+'                         # 5: Qty
            r'([\d,]+\.?\d*)\s+'               # 6: MRP
            r'([\d,]+\.?\d*)\s+'               # 7: Basic Price
            r'([\d,]+\.?\d*)\s+'               # 8: Next number (could be DP or CGST)
            r'([\d.]+)\s+'                      # 9: Next number (could be CGST or SGST)
            r'([\d.]+)\s+'                      # 10: Next number (could be SGST or something else)
            r'.*?'                              # Skip other columns
            r'([\d,]+\.?\d*)\s+'               # 11: Landed Price
            r'([\d,]+\.?\d*)$',                # 12: Total Value
            line
        )
        
        if m_old_flexible:
            # Determine if there's a DP column by checking the values
            val1 = float(m_old_flexible.group(8).replace(",", ""))  # Could be DP or CGST
            val2 = float(m_old_flexible.group(9))                    # Could be CGST or SGST
            val3 = float(m_old_flexible.group(10))                   # Could be SGST or other
            
            # CGST and SGST are typically 9.00, 12.00, 14.00, 18.00 (single/low double digits)
            # DP is typically 20.00, 25.00, 30.00 or higher
            # If val1 is > 18, it's likely DP, so CGST=val2, SGST=val3
            # If val1 is <= 18, it's likely CGST, so CGST=val1, SGST=val2
            
            if val1 > 18.0:
                # Has DP column: val1=DP, val2=CGST, val3=SGST
                cgst_val = val2
                sgst_val = val3
            else:
                # No DP column: val1=CGST, val2=SGST
                cgst_val = val1
                sgst_val = val2
            
            m_old = m_old_flexible

        matched = False
        
        if m_new:
            # NEW FORMAT (IGST)
            sr_no = m_new.group(1)
            ean = m_new.group(2)
            hsn_part1 = m_new.group(3)
            desc_part1 = m_new.group(4).strip()
            qty = m_new.group(5)
            mrp = m_new.group(6).replace(",", "")
            basic = m_new.group(7).replace(",", "")
            igst = float(m_new.group(8))
            landed = m_new.group(9).replace(",", "")
            total = m_new.group(10).replace(",", "")
            gst_pct = igst
            matched = True
            
        elif m_old:
            # OLD FORMAT (CGST+SGST)
            sr_no = m_old.group(1)
            ean = m_old.group(2)
            hsn_part1 = m_old.group(3)
            desc_part1 = m_old.group(4).strip()
            qty = m_old.group(5)
            mrp = m_old.group(6).replace(",", "")
            basic = m_old.group(7).replace(",", "")
            # Use the detected cgst_val and sgst_val from the smart detection above
            cgst = cgst_val
            sgst = sgst_val
            landed = m_old.group(11).replace(",", "")
            total = m_old.group(12).replace(",", "")
            gst_pct = cgst + sgst
            matched = True

        if matched:
            # Line 2 has: article_no + hsn_part2 + desc_part2 + 1.00 + ...
            desc_part2 = ""
            hsn_part2 = ""
            if i + 1 < len(all_lines):
                next_line = all_lines[i + 1].strip()
                m2 = re.match(r'^(\d+)\s+(\d+)\s+(.+?)\s+1\.00', next_line)
                if m2:
                    hsn_part2 = m2.group(2)
                    desc_part2 = m2.group(3).strip()

            # Full HSN code
            hsn = hsn_part1 + hsn_part2

            # Full description
            product_name = (desc_part1 + " " + desc_part2).strip()

            items.append({
                "Sr #": int(sr_no),
                "EAN": ean,
                "Product Name": product_name,
                "HSN Code": hsn,
                "Quantity": int(qty),
                "MRP": float(mrp),
                "Base Rate": float(basic),
                "GST %": gst_pct,
                "Total": float(total),
            })

        i += 1

    return pd.DataFrame(items)


# ------------------ SUMMARY EXTRACTION ------------------

def extract_summary(pdf_path):

    grand_total = 0.0

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                m = re.match(r'^Total\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)$', line)
                if m:
                    grand_total = float(m.group(2).replace(",", ""))
                    break

    return grand_total


# ================== PUBLIC FUNCTION ==================

def convert_pdf_to_excel(pdf_path, output_excel_path):

    header_data = extract_po_header(pdf_path)
    products = extract_line_items(pdf_path)

    if products.empty:
        raise Exception("No line items found in DMart PO")

    grand_total = extract_summary(pdf_path)
    total_base = round((products["Base Rate"] * products["Quantity"]).sum(), 2)
    total_tax = round(grand_total - total_base, 2)

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
