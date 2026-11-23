"""
Builds the downloadable Excel template shipped with the Confirm Nepal app
(served at GET /template). Run this once to (re)generate static/confirm_nepal_template.xlsx.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

wb = Workbook()

# ---------------------------------------------------------------- Data sheet
ws = wb.active
ws.title = "Data"

headers = ["Month", "Date", "Bill No", "Seller Name", "Address", "VAT No.", "Amount"]
header_fill = PatternFill("solid", fgColor="1F2937")
header_font = Font(name="Arial", bold=True, color="FFFFFF")
example_fill = PatternFill("solid", fgColor="FFF7CC")

for col, h in enumerate(headers, start=1):
    c = ws.cell(row=1, column=col, value=h)
    c.font = header_font
    c.fill = header_fill
    c.alignment = Alignment(horizontal="center")

widths = [10, 12, 14, 32, 26, 16, 14]
for col, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(col)].width = w

example_row = [3, "2082-03-15", "4888", "Suryamukhi Trade Link Pvt. Ltd", "Dharan-1, Sunsari", "619727228", 57750]
for col, val in enumerate(example_row, start=1):
    c = ws.cell(row=2, column=col, value=val)
    c.font = Font(name="Arial")
    c.fill = example_fill

ws.freeze_panes = "A2"

# ------------------------------------------------------------- Legend sheet
legend = wb.create_sheet("Read Me")
legend.column_dimensions["A"].width = 20
legend.column_dimensions["B"].width = 90

rows = [
    ("Confirm Nepal — Template Guide", ""),
    ("", ""),
    ("Column", "What to put here"),
    ("Month", "Bikram Sambat month number (1-12) the bill falls under. Used only for your own reference."),
    ("Date", "Bill date, any format you like (e.g. 2082-03-15). Not used in calculations."),
    ("Bill No", "Your internal bill/invoice number for this transaction."),
    ("Seller Name", "Required. Exact legal or trade name of the seller/vendor."),
    ("Address", "Optional. Seller's address, printed on the letter if provided."),
    ("VAT No.", "Required. Seller's VAT registration number."),
    ("Amount", "Required. Taxable (pre-VAT) amount for this bill, as a number (no currency symbol or commas)."),
    ("", ""),
    ("Notes", ""),
    ("", "Row 2 (highlighted) is an example — delete it before uploading, or leave it and it will just show up as a tiny extra bill."),
    ("", "One row per bill. The app sums every bill per seller (matched by Seller Name + VAT No.) and generates one confirmation letter per seller whose total crosses your chosen threshold."),
    ("", "Do not rename, remove, or reorder the headers in row 1 of the Data sheet."),
]
for r_idx, (a, b) in enumerate(rows, start=1):
    ca = legend.cell(row=r_idx, column=1, value=a)
    cb = legend.cell(row=r_idx, column=2, value=b)
    ca.font = Font(name="Arial", bold=(r_idx in (1, 3, 12)))
    cb.font = Font(name="Arial", bold=(r_idx == 3))
    cb.alignment = Alignment(wrap_text=True, vertical="top")

wb.save("static/confirm_nepal_template.xlsx")
print("saved static/confirm_nepal_template.xlsx")
