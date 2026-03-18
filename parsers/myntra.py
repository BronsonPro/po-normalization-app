import pdfplumber
import pandas as pd
import re

def convert_pdf_to_excel(pdf_path, output_excel_path):
    """Diagnostic version - shows what pdfplumber extracts"""
    
    print("\n========== MYNTRA PDF DIAGNOSTIC ==========")
    
    with pdfplumber.open(pdf_path) as pdf:
        print(f"Total pages: {len(pdf.pages)}")
        
        for page_num, page in enumerate(pdf.pages):
            print(f"\n--- Page {page_num + 1} ---")
            
            # Try text extraction
            text = page.extract_text()
            if text:
                print(f"Text extraction: SUCCESS ({len(text)} chars)")
                print("First 500 chars:")
                print(text[:500])
                print(f"\nBNPL count in text: {text.count('BNPL')}")
            else:
                print("Text extraction: FAILED (empty)")
            
            # Try table extraction
            tables = page.extract_tables()
            if tables:
                print(f"\nTable extraction: SUCCESS ({len(tables)} tables)")
                for table_num, table in enumerate(tables):
                    print(f"\nTable {table_num + 1}: {len(table)} rows")
                    if table:
                        print(f"First row: {table[0][:5] if len(table[0]) > 5 else table[0]}")
                        if len(table) > 1:
                            print(f"Second row: {table[1][:5] if len(table[1]) > 5 else table[1]}")
                        
                        # Check for BNPL
                        bnpl_rows = [r for r in table if any('BNPL' in str(cell) for cell in r)]
                        print(f"Rows with BNPL: {len(bnpl_rows)}")
                        if bnpl_rows:
                            print(f"First BNPL row: {bnpl_rows[0][:10]}")
            else:
                print("\nTable extraction: FAILED (no tables)")
    
    print("\n========== END DIAGNOSTIC ==========\n")
    
    raise Exception("DIAGNOSTIC MODE - Check the output above to understand PDF structure")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 2:
        convert_pdf_to_excel(sys.argv[1], sys.argv[2])
