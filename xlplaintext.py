#!/usr/bin/env python3
"""
Take an XLSX workbook and convert it to plain text as follows:
  1. Copy it to a temporary file, so that the user can keep the original
     file open if required.
  2. Using ``pywin32``, open the copy then use the COM *SaveAs* method
     combined with the ``xlTextPrinter`` format to generate the text
     file output.
  3. Clean up the temporary file.

If the XLSX workbook was laid out using a fixed-width font, when
converted to a text file columns will be aligned as closely as possible
to the original layout.
"""
import argparse
import os
import pywintypes
import shutil
import tempfile
import win32com.client as win32

# Written by Eric J. Whitney, April 2026.

# ======================================================================

# noinspection PyUnresolvedReferences
def convert_xlsx_to_text(xlsx_file, txt_file):
    """
    Convert XLSX file to text file using Excel's xlTextPrinter format.
    """
    # Create temporary copy of XLSX.
    fd, temp_xlsx = tempfile.mkstemp(suffix='.xlsx', prefix='temp_')
    os.close(fd)  # Close the file descriptor; we just need the name.
    shutil.copy2(xlsx_file, temp_xlsx)

    # Ensure txt_file is a full path.
    txt_file = os.path.abspath(txt_file)

    # Start Excel COM automation.
    excel, was_running = start_excel()
    workbook: win32.CDispatch | None = None
    try:
        # Open the temporary file and save as text using xlTextPrinter
        # format (36).
        workbook = excel.Workbooks.Open(temp_xlsx)
        xlTextPrinter = 36
        workbook.SaveAs(txt_file, FileFormat=xlTextPrinter)

    except pywintypes.com_error as e:
        print(f"Error during conversion to plain text: {e}")
        raise

    finally:
        if workbook:
            workbook.Close(SaveChanges=False)

        if excel and not was_running:
            excel.Quit()

        if os.path.exists(temp_xlsx):
            os.remove(temp_xlsx)


# ----------------------------------------------------------------------

def start_excel() -> tuple[win32.CDispatch, bool]:
    # First try to get existing Excel instance.
    try:
        excel = win32.GetActiveObject("Excel.Application")
        return excel, True

    except pywintypes.com_error:
        pass

    # Perhaps Excel was not already running.  Launch it.
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel, False

    except pywintypes.com_error as e:
        # Something else is wrong.
        print(f"Error launching Excel: {e}")
        raise


# ----------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="""
        Convert an XLSX workbook to plain text using Excel's
        `xlTextPrinter` save format.   If the workbook was laid out
        using a fixed-width font, then conversion should give a text
        file with columns aligned simlarly to the original file.
        """
    )
    parser.add_argument('xlsx_file', help='Input XLSX file path')
    parser.add_argument('txt_file', nargs='?',
                        help="Output text file path (optional)")
    parser.add_argument('-f', '--force', action='store_true',
                        help="Force overwrite of the output file if it "
                             "already exists")

    args = parser.parse_args()

    # Generate txt_file if not provided.
    if args.txt_file is None:
        txt_file = os.path.splitext(args.xlsx_file)[0] + '.txt'
    else:
        txt_file = args.txt_file

    # Check if output file already exists.
    if os.path.exists(txt_file) and not args.force:
        raise RuntimeError(f"Output file already exists: {txt_file}. "
                           f"Use --force to replace it.")

    convert_xlsx_to_text(args.xlsx_file, txt_file)
    print(f"Successfully converted: {args.xlsx_file} -> {txt_file}")

# ======================================================================

if __name__ == "__main__":
    main()
