import pdfplumber
import pandas as pd
import re


def convert_pdf_to_excel(pdf_path: str, output_path: str) -> None:
    """
    Parse a FOY Purchase Order PDF and write a standardised Excel file
    matching the app.py expected format (header rows + table + summary).
    """
    header_info, items, summary = _extract_all(pdf_path)

    if not items:
        raise ValueError("No line items could be extracted from the FOY PO PDF.")

    # Build item DataFrame
    df_items = pd.DataFrame(items)
    df_items.insert(2, "EAN", "")  # ← ADD THIS BACK

    # Header metadata rows (Party Name, PO No, PO Date, etc.)
    header_rows = list(header_info.items())  # list of (key, val) tuples

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Write header metadata (no column headers, raw rows)
        df_header = pd.DataFrame(header_rows)
        df_header.to_excel(writer, index=False, header=False,
                           sheet_name="FOY PO", startrow=0)

        # Blank separator rows
        blank_start = len(header_rows)
        ws = writer.sheets["FOY PO"]
        for i in range(3):
            ws.cell(row=blank_start + i + 1, column=1, value="")

        # Write item table below header block + blanks
        table_start = blank_start + 3
        df_items.to_excel(writer, index=False, sheet_name="FOY PO",
                          startrow=table_start)

        # Write summary rows below items
        summary_start = table_start + len(df_items) + 2  # +1 header +1 blank
        summary_rows = [
            ("Total Base Value", summary.get("taxable_total", "")),
            ("Total Tax",        summary.get("tax_total", "")),
            ("Grand Total",      summary.get("grand_total", "")),
        ]
        for i, (label, val) in enumerate(summary_rows):
            ws.cell(row=summary_start + i + 1, column=1, value=label)
            ws.cell(row=summary_start + i + 1, column=2, value=val)

    print(f"✅ FOY parser: {len(df_items)} items written to {output_path}")


def _extract_all(pdf_path: str):
    """Extract header info, line items and summary from FOY PO PDF."""

    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                pages_text.append(text)

    full_text = "\n".join(pages_text)

    # ── Header metadata ────────────────────────────────────────────────
    header_info = {}
    header_info["Party Name"] = "FOY E-RETAIL PVT LTD"

    po_no = re.search(r'PO No\s*:\s*(\S+)', full_text)
    if po_no:
        header_info["PO No"] = po_no.group(1).strip()

    po_date = re.search(r'PO Date\s*:\s*([A-Za-z]+ \d+,\s*\d{4})', full_text)
    if po_date:
        header_info["PO Date"] = po_date.group(1).strip()

    po_release = re.search(r'PO Release Date\s*:\s*([A-Za-z]+ \d+,\s*\d{4})', full_text)
    if po_release:
        header_info["PO Expiry Date"] = po_release.group(1).strip()

    # Shipping address — warehouse address block
    ship_match = re.search(
        r'FOY E-RETAIL PVT LTD \(WAREHOUSE\)(.*?)(?:GSTIN|$)',
        full_text, re.DOTALL
    )
    if ship_match:
        addr = re.sub(r'\s+', ' ', ship_match.group(1)).strip()
        header_info["Shipping Address"] = addr

    # GSTIN of FOY (second occurrence = FOY's GSTIN)
    gstin_matches = re.findall(r'GSTIN\s*:(\S+)', full_text)
    if len(gstin_matches) >= 2:
        header_info["GST No"] = gstin_matches[1].strip()
    elif gstin_matches:
        header_info["GST No"] = gstin_matches[0].strip()

    # ── Remove page-break carry-over orphan total ──────────────────────
    full_text = re.sub(
        r'(Brand:[^\n]+\n)([\d.]+\n)(\d+\s+FOY)',
        r'\1\3',
        full_text,
    )

    # ── Line items ─────────────────────────────────────────────────────
    pattern = re.compile(
        r'(\d+)\s+(FOY\w+)\s*:(.*?)'    # sno, item code, product name start
        r'(?:Colour:[^\n]*\n)?'
        r'(?:Size:[^\n]*\n)?'
        r'(?:Brand:[^\n]*\n)?'
        r'(\d{8})\s+'                    # HSN Code
        r'(\d+)\s+'                      # Qty
        r'([\d.]+)\s+'                   # Base Rate
        r'([\d.]+)\s+'                   # Discount
        r'([\d.]+)\s+'                   # Taxable Value
        r'([\d.]+)\s+'                   # CGST Rate
        r'([\d.]+)\s+'                   # CGST Amt
        r'([\d.]+)\s+'                   # SGST Rate
        r'([\d.]+)\s+'                   # SGST Amt
        r'([\d.]+)\s+'                   # IGST Rate
        r'([\d.]+)\s+'                   # IGST Amt
        r'([\d.]+)',                      # Total
        re.DOTALL,
    )

    items = []
    for m in pattern.finditer(full_text):
        name_raw = m.group(3)
        name = re.sub(r'\s+', ' ', name_raw).strip()
        name = re.sub(r'\s*Colour:.*', '', name, flags=re.DOTALL).strip()
        name = re.sub(r'\s*Size:.*',   '', name, flags=re.DOTALL).strip()
        name = re.sub(r'\s*Brand:.*',  '', name, flags=re.DOTALL).strip()
        # Fix mid-word line breaks (e.g. "Bronson Pr ofessional" → "Bronson Professional")
        name = re.sub(r'(?<=[a-zA-Z]) (?=[a-z])', '', name)

        cgst_rate = float(m.group(9))
        sgst_rate = float(m.group(11))
        igst_rate = float(m.group(13))

        # GST %: CGST+SGST if no IGST, else IGST
        if igst_rate > 0:
            gst_pct = igst_rate
        else:
            gst_pct = cgst_rate + sgst_rate

        items.append({
            'Sr No':         int(m.group(1)),
            'Item Code':     m.group(2),
            'Product Name':  name,
            'HSN Code':      m.group(4),
            'Qty':           int(m.group(5)),
            'Base Rate':     float(m.group(6))
            'GST %':         gst_pct,
            'Discount':      float(m.group(7)),
            'Taxable Value': float(m.group(8)),
            'CGST Rate':     cgst_rate,
            'CGST Amt':      float(m.group(10)),
            'SGST Rate':     sgst_rate,
            'SGST Amt':      float(m.group(12)),
            'IGST Rate':     igst_rate,
            'IGST Amt':      float(m.group(14)),
            'Total':         float(m.group(15)),
        })

    # ── Summary ────────────────────────────────────────────────────────
    summary = {}
    taxable_match = re.search(r'Total Amount \(INR\)\s*([\d.]+)', full_text)
    tax_match     = re.search(r'Total Tax \(INR\)\s*([\d.]+)', full_text)
    grand_match   = re.search(r'Grand Total \(INR\)\s*([\d.]+)', full_text)

    if taxable_match:
        summary["taxable_total"] = float(taxable_match.group(1))
    if tax_match:
        summary["tax_total"] = float(tax_match.group(1))
    if grand_match:
        summary["grand_total"] = float(grand_match.group(1))

    return header_info, items, summary
