import pdfplumber
import pandas as pd
import re


def convert_pdf_to_excel(pdf_path: str, output_path: str) -> None:
    """
    Parse a FOY Purchase Order PDF and write a standardised Excel file.

    Columns produced (matching the app.py expected schema):
        Sr No | Item Code | EAN | Product Name | HSN Code | Qty |
        Base Rate | Discount | Taxable Value |
        CGST Rate | CGST Amt | SGST Rate | SGST Amt |
        IGST Rate | IGST Amt | Total
    """
    items = _extract_line_items(pdf_path)

    if not items:
        raise ValueError("No line items could be extracted from the FOY PO PDF.")

    df = pd.DataFrame(items)

    # EAN is not present in the PO — it is looked up from master via Item Code in app.py
    df.insert(2, "EAN", "")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="FOY PO")

    print(f"✅ FOY parser: {len(df)} items written to {output_path}")


def _extract_line_items(pdf_path: str) -> list[dict]:
    """Extract all line items from the FOY PO PDF."""
    r = PdfReader(pdf_path)
    pages_text = [page.extract_text() or "" for page in r.pages]
    full_text = "\n".join(pages_text)

    # Remove page-break carry-over: a standalone total number that appears at
    # the top of the next page (e.g. "14699.99\n9 FOY003538 :")
    full_text = re.sub(
        r'(Brand:[^\n]+\n)([\d.]+\n)(\d+\s+FOY)',
        r'\1\3',
        full_text,
    )

    # Each item follows the pattern:
    #   <sno>  FOY<code> : <product name (multi-line)>
    #   [Colour: ...]  [Size: ...]  [Brand: ...]
    #   <HSN:8d>  <Qty>  <BaseCost>  <Discount>  <TaxableValue>
    #   <CGSTRate>  <CGSTAmt>  <SGSTRate>  <SGSTAmt>
    #   <IGSTRate>  <IGSTAmt>  <Total>
    pattern = re.compile(
        r'(\d+)\s+(FOY\w+)\s*:(.*?)'    # sno, item code, product name start
        r'(?:Colour:[^\n]*\n)?'          # optional Colour line
        r'(?:Size:[^\n]*\n)?'            # optional Size line
        r'(?:Brand:[^\n]*\n)?'           # optional Brand line
        r'(\d{8})\s+'                    # HSN Code (8 digits)
        r'(\d+)\s+'                      # Qty
        r'([\d.]+)\s+'                   # Base Cost / Rate
        r'([\d.]+)\s+'                   # Discount
        r'([\d.]+)\s+'                   # Taxable Value
        r'([\d.]+)\s+'                   # CGST Rate
        r'([\d.]+)\s+'                   # CGST Amt
        r'([\d.]+)\s+'                   # SGST Rate
        r'([\d.]+)\s+'                   # SGST Amt
        r'([\d.]+)\s+'                   # IGST Rate
        r'([\d.]+)\s+'                   # IGST Amt
        r'([\d.]+)',                      # Total (INR)
        re.DOTALL,
    )

    items = []
    for m in pattern.finditer(full_text):
        name_raw = m.group(3)
        # Clean up the product name: collapse whitespace, drop trailing metadata
        name = re.sub(r'\s+', ' ', name_raw).strip()
        name = re.sub(r'\s*Colour:.*', '', name, flags=re.DOTALL).strip()
        name = re.sub(r'\s*Size:.*',   '', name, flags=re.DOTALL).strip()
        name = re.sub(r'\s*Brand:.*',  '', name, flags=re.DOTALL).strip()

        items.append({
            'Sr No':         int(m.group(1)),
            'Item Code':     m.group(2),
            'Product Name':  name,
            'HSN Code':      m.group(4),
            'Qty':           int(m.group(5)),
            'Base Rate':     float(m.group(6)),
            'Discount':      float(m.group(7)),
            'Taxable Value': float(m.group(8)),
            'CGST Rate':     float(m.group(9)),
            'CGST Amt':      float(m.group(10)),
            'SGST Rate':     float(m.group(11)),
            'SGST Amt':      float(m.group(12)),
            'IGST Rate':     float(m.group(13)),
            'IGST Amt':      float(m.group(14)),
            'Total':         float(m.group(15)),
        })

    return items
