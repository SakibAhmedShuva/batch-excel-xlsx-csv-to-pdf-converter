import os
import pandas as pd
from weasyprint import HTML, CSS
from weasyprint.fonts import FontConfiguration # For better font handling if needed
import glob

# --- Configuration ---
INPUT_XLSX_DIR = "input_xlsx_files"  # Folder containing your .xlsx files
OUTPUT_PDF_DIR = "output_pdf_files"  # Folder where PDFs will be saved
RECURSIVE_SEARCH = False             # Set to True to search in subdirectories of INPUT_XLSX_DIR
CREATE_SUBFOLDERS_IN_OUTPUT = True   # If RECURSIVE_SEARCH is True, mirror subfolder structure in output

# Basic CSS for table styling (optional, but recommended for readability)
TABLE_CSS = """
@page {
    size: A4 landscape; /* Or A4 portrait, letter, etc. */
    margin: 0.75in;
}
body {
    font-family: 'Arial', sans-serif; /* Consider installing a common font */
    font-size: 9pt;
}
table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 15px; /* Space between tables if multiple sheets */
}
th, td {
    border: 1px solid #cccccc;
    padding: 4px 6px;
    text-align: left;
    word-wrap: break-word; /* Prevent long text from overflowing */
}
th {
    background-color: #f2f2f2;
    font-weight: bold;
}
caption {
    caption-side: top;
    font-size: 1.2em;
    font-weight: bold;
    margin-bottom: 10px;
    text-align: left;
}
"""

def convert_xlsx_to_pdf(xlsx_path, pdf_path):
    """
    Converts a single XLSX file to a PDF.
    Each sheet in the XLSX will be a separate table in the PDF.
    If only one sheet, the PDF will be named directly.
    If multiple sheets, the sheet name will be appended to the PDF filename,
    UNLESS all sheets are combined into one PDF (see 'combine_sheets' logic).
    """
    try:
        xls = pd.ExcelFile(xlsx_path)
        sheet_names = xls.sheet_names

        if not sheet_names:
            print(f"WARNING: No sheets found in {xlsx_path}. Skipping.")
            return

        # --- Option 1: Combine all sheets into a single PDF ---
        # (Uncomment this block and comment out Option 2 if you want one PDF per XLSX file)
        # all_html_content = ""
        # for sheet_name in sheet_names:
        #     df = xls.parse(sheet_name)
        #     if df.empty:
        #         print(f"  Sheet '{sheet_name}' in {xlsx_path} is empty. Skipping this sheet.")
        #         continue
        #     # Add a caption for the sheet name
        #     table_html = f"<caption>Sheet: {sheet_name}</caption>" + df.to_html(index=False, escape=True)
        #     all_html_content += table_html
        #
        # if not all_html_content:
        #     print(f"WARNING: All sheets in {xlsx_path} were empty or skipped. No PDF generated.")
        #     return
        #
        # full_html = f"<html><head><meta charset='UTF-8'></head><body>{all_html_content}</body></html>"
        # HTML(string=full_html).write_pdf(pdf_path, stylesheets=[CSS(string=TABLE_CSS)])
        # print(f"  Successfully converted {xlsx_path} (all sheets) to {pdf_path}")
        # --- End of Option 1 ---


        # --- Option 2: Create one PDF per sheet (if multiple sheets) ---
        # (Comment this block and uncomment Option 1 if you want one PDF per XLSX file)
        if len(sheet_names) == 1:
            df = xls.parse(sheet_names[0])
            if df.empty:
                print(f"  Sheet '{sheet_names[0]}' in {xlsx_path} is empty. Skipping.")
                return
            html_content = df.to_html(index=False, escape=True) # escape=True for special chars
            full_html = f"<html><head><meta charset='UTF-8'></head><body>{html_content}</body></html>"
            HTML(string=full_html).write_pdf(pdf_path, stylesheets=[CSS(string=TABLE_CSS)])
            print(f"  Successfully converted {xlsx_path} (sheet: {sheet_names[0]}) to {pdf_path}")
        else:
            base_pdf_name, ext = os.path.splitext(pdf_path)
            for sheet_name in sheet_names:
                df = xls.parse(sheet_name)
                if df.empty:
                    print(f"  Sheet '{sheet_name}' in {xlsx_path} is empty. Skipping this sheet.")
                    continue
                # Sanitize sheet name for use in filename
                safe_sheet_name = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in sheet_name).rstrip()
                sheet_pdf_path = f"{base_pdf_name}_{safe_sheet_name}{ext}"
                html_content = df.to_html(index=False, escape=True)
                full_html = f"<html><head><meta charset='UTF-8'></head><body>{html_content}</body></html>"
                HTML(string=full_html).write_pdf(sheet_pdf_path, stylesheets=[CSS(string=TABLE_CSS)])
                print(f"  Successfully converted {xlsx_path} (sheet: {sheet_name}) to {sheet_pdf_path}")
        # --- End of Option 2 ---

    except FileNotFoundError:
        print(f"ERROR: XLSX file not found: {xlsx_path}")
    except ImportError as e:
        print(f"ERROR: Missing a required library. Please ensure pandas, openpyxl, and weasyprint are installed. {e}")
    except Exception as e:
        print(f"ERROR: Failed to convert {xlsx_path} to PDF. Reason: {e}")
        import traceback
        traceback.print_exc()


def main():
    """
    Main function to find XLSX files and initiate conversion.
    """
    if not os.path.isdir(INPUT_XLSX_DIR):
        print(f"ERROR: Input directory '{INPUT_XLSX_DIR}' not found. Please create it and add XLSX files.")
        return

    os.makedirs(OUTPUT_PDF_DIR, exist_ok=True)
    print(f"Looking for .xlsx files in: {os.path.abspath(INPUT_XLSX_DIR)}")
    print(f"PDFs will be saved in: {os.path.abspath(OUTPUT_PDF_DIR)}")

    xlsx_files_to_process = []
    if RECURSIVE_SEARCH:
        for root, _, files in os.walk(INPUT_XLSX_DIR):
            for file in files:
                if file.lower().endswith(".xlsx"):
                    xlsx_files_to_process.append(os.path.join(root, file))
    else:
        for file in glob.glob(os.path.join(INPUT_XLSX_DIR, "*.xlsx")):
            xlsx_files_to_process.append(file)

    if not xlsx_files_to_process:
        print("No .xlsx files found to convert.")
        return

    print(f"\nFound {len(xlsx_files_to_process)} .xlsx file(s) to process.\n")

    for xlsx_file_path in xlsx_files_to_process:
        print(f"Processing: {xlsx_file_path}...")
        base_name = os.path.splitext(os.path.basename(xlsx_file_path))[0]
        pdf_filename = f"{base_name}.pdf"

        current_output_dir = OUTPUT_PDF_DIR
        if RECURSIVE_SEARCH and CREATE_SUBFOLDERS_IN_OUTPUT:
            # Get the relative path from the input base directory
            relative_path = os.path.relpath(os.path.dirname(xlsx_file_path), INPUT_XLSX_DIR)
            if relative_path and relative_path != '.':
                current_output_dir = os.path.join(OUTPUT_PDF_DIR, relative_path)
                os.makedirs(current_output_dir, exist_ok=True)

        pdf_file_path = os.path.join(current_output_dir, pdf_filename)
        convert_xlsx_to_pdf(xlsx_file_path, pdf_file_path)
        print("-" * 30)

    print("\nBatch conversion finished.")

if __name__ == "__main__":
    # Example usage:
    # 1. Create a folder named "input_xlsx_files" in the same directory as this script.
    # 2. Place your .xlsx files into "input_xlsx_files".
    # 3. Run this script: python your_script_name.py
    # 4. A folder "output_pdf_files" will be created with the converted PDFs.

    # For advanced font configuration with weasyprint (if default fonts are problematic):
    # font_config = FontConfiguration()
    # You might need to point to specific font files if they are not system-wide.
    # css_with_fonts = CSS(string='@font-face { font-family: MyCustomFont; src: url(file:///path/to/your/font.ttf); } body { font-family: "MyCustomFont"; }' + TABLE_CSS, font_config=font_config)
    # Then pass `stylesheets=[css_with_fonts]` to `write_pdf`.
    # For most cases, the default TABLE_CSS with common system fonts like Arial should work.

    main()