"""
Confirm Nepal — confirmation letter generation.

Stateless by design: everything here operates on in-memory bytes / temp files
for a single request. Nothing is persisted.
"""
from __future__ import annotations

import io
import os
import re
import zipfile
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from PIL import Image

REQUIRED_COLUMNS = ["Seller Name", "VAT No.", "Amount"]
OPTIONAL_COLUMNS = ["Address", "Bill No", "Month", "Date"]
ALL_COLUMNS = REQUIRED_COLUMNS + OPTIONAL_COLUMNS


class TemplateError(Exception):
    """Raised when the uploaded workbook doesn't match the expected template."""


def sanitize_string(text) -> str:
    return re.sub(r"[^\x20-\x7E]+", "", str(text)).strip()


@dataclass
class LetterBranding:
    """Everything the letter-writer supplies about their own business."""
    business_name: str
    business_address: str = ""
    business_email: str = ""
    fiscal_year_label: str = ""      # e.g. "2081/82"
    period_end_label: str = ""       # e.g. "end of Ashad 2082"
    letter_date_label: str = ""      # e.g. "2082/06/11"
    vat_rate_percent: float = 13.0
    threshold: float = 100000.0
    letterhead_bytes: Optional[bytes] = None
    signature_bytes: Optional[bytes] = None
    stamp_bytes: Optional[bytes] = None


@dataclass
class ParsedWorkbook:
    bills: pd.DataFrame
    warnings: list[str] = field(default_factory=list)


def parse_workbook(file_bytes: bytes) -> ParsedWorkbook:
    """
    Reads the Confirm Nepal template. Expects a 'Data' sheet (or the first
    sheet, for leniency) with fixed headers in row 1 -- see build_template.py.
    """
    try:
        sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
    except Exception as e:
        raise TemplateError(f"Couldn't read this as an Excel file: {e}")

    sheet_name = "Data" if "Data" in sheets else next(iter(sheets))
    df = sheets[sheet_name].copy()
    df.columns = df.columns.astype(str).str.strip()

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise TemplateError(
            "This doesn't look like the Confirm Nepal template -- missing column(s): "
            + ", ".join(missing)
            + ". Download the template from the app and fill that in rather than "
              "renaming your own columns."
        )

    warnings = []
    df["Seller Name"] = df["Seller Name"].apply(
        lambda v: sanitize_string(v) if pd.notna(v) else ""
    )
    df["VAT No."] = df["VAT No."].apply(
        lambda v: sanitize_string(v) if pd.notna(v) else ""
    )
    if "Address" in df.columns:
        df["Address"] = df["Address"].apply(
            lambda v: sanitize_string(v) if pd.notna(v) else ""
        )
    else:
        df["Address"] = ""

    if "Bill No" not in df.columns:
        df["Bill No"] = ""

    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")

    total_rows = len(df)
    bad_amount = df["Amount"].isna() & (df["Seller Name"] != "")
    if bad_amount.any():
        warnings.append(
            f"{int(bad_amount.sum())} row(s) had a non-numeric or blank Amount and were skipped."
        )

    df = df[(df["Seller Name"] != "") & (df["VAT No."] != "") & df["Amount"].notna()].copy()

    blank_rows = total_rows - len(df) - int(bad_amount.sum())
    if blank_rows > 0:
        warnings.append(f"{blank_rows} blank row(s) were ignored.")

    if df.empty:
        raise TemplateError(
            "No usable rows found. Make sure Seller Name, VAT No., and Amount are filled in."
        )

    return ParsedWorkbook(bills=df, warnings=warnings)


def summarize_sellers(bills: pd.DataFrame, threshold: float) -> pd.DataFrame:
    grouped = (
        bills.groupby(["Seller Name", "VAT No.", "Address"], dropna=False)["Amount"]
        .sum()
        .reset_index()
        .rename(columns={"Amount": "Total Taxable"})
    )
    result = grouped[grouped["Total Taxable"] > threshold].sort_values(
        "Total Taxable", ascending=False
    )
    result.insert(0, "S.No", range(1, len(result) + 1))
    return result.reset_index(drop=True)


def _write_image_from_bytes(pdf: FPDF, img_bytes: bytes, x: float, y: float, w: float) -> float:
    """Writes an image from raw bytes and returns its rendered height in mm."""
    img = Image.open(io.BytesIO(img_bytes))
    iw, ih = img.size
    h = (w / iw) * ih
    tmp_path = f"/tmp/_cn_img_{id(img_bytes)}.png"
    img.convert("RGB").save(tmp_path)
    pdf.image(tmp_path, x=x, y=y, w=w, h=h)
    os.remove(tmp_path)
    return h


def generate_letter_pdf(
    seller_name: str,
    vat_no: str,
    address: str,
    total_taxable: float,
    bills: pd.DataFrame,
    branding: LetterBranding,
) -> bytes:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(left=10, top=10, right=10)
    pdf.add_page()

    if branding.letterhead_bytes:
        h = _write_image_from_bytes(pdf, branding.letterhead_bytes, x=10, y=10, w=190)
        pdf.set_y(10 + min(h, 40) + 3)
    else:
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 10, branding.business_name, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        pdf.set_y(pdf.get_y() + 2)

    pdf.set_font("Helvetica", size=9)
    if branding.letter_date_label:
        pdf.cell(0, 8, f"Date: {branding.letter_date_label}", align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "To,", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font("Helvetica", size=10)
    recipient = seller_name
    if address:
        recipient += f"\nAddress: {address}"
    recipient += f"\nVAT NO: {vat_no}"
    pdf.multi_cell(w=pdf.epw, h=5, text=recipient)

    pdf.ln(5)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 8, "Subject: Confirmation of Account Statement.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(3)
    pdf.set_font("Helvetica", size=9)
    fy = branding.fiscal_year_label or "the relevant fiscal year"
    period = branding.period_end_label or "the period end"
    pdf.multi_cell(
        w=pdf.epw, h=4,
        text=(
            f"Dear Sir/Madam,\n\nAs per our account records for the fiscal year {fy} "
            f"at {period}, total Purchase and the balance in your account are as follows:"
        ),
    )

    vat_rate = branding.vat_rate_percent / 100.0
    vat_amount = total_taxable * vat_rate
    total_amount = total_taxable + vat_amount

    pdf.ln(5)
    pdf.set_font("Helvetica", "B", 9)
    headers = ["Particulars", "Taxable", f"VAT ({branding.vat_rate_percent:g}%)", "Total"]
    col_widths = [70, 40, 40, 40]
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 8, h, border=1, align="C")
    pdf.ln()
    pdf.set_font("Helvetica", size=9)
    data_rows = [
        ("Net Purchase Amount", total_taxable, vat_amount, total_amount),
        ("Total", total_taxable, vat_amount, total_amount),
    ]
    for label, tax, vat, tot in data_rows:
        pdf.cell(col_widths[0], 8, label, border=1)
        pdf.cell(col_widths[1], 8, f"{tax:,.2f}", border=1, align="C")
        pdf.cell(col_widths[2], 8, f"{vat:,.2f}", border=1, align="C")
        pdf.cell(col_widths[3], 8, f"{tot:,.2f}", border=1, align="C")
        pdf.ln()

    pdf.ln(5)
    pdf.set_font("Helvetica", size=9)
    pdf.multi_cell(
        w=pdf.epw, h=4,
        text=(
            "Please check the above details and return this confirmation duly signed "
            "and stamped within 15 days, or else we shall treat the above balance as "
            "confirmed and accepted."
        ),
    )
    pdf.ln(3)
    pdf.multi_cell(w=pdf.epw, h=4, text="Thank You,\n\nWith Regards,")

    current_y = pdf.get_y()
    sig_h = stamp_h = 0.0
    if branding.signature_bytes:
        sig_h = _write_image_from_bytes(pdf, branding.signature_bytes, x=10, y=current_y, w=30)
    if branding.stamp_bytes:
        stamp_h = _write_image_from_bytes(pdf, branding.stamp_bytes, x=55, y=current_y, w=40)
    pdf.set_y(current_y + max(sig_h, stamp_h) + 2)

    footer_lines = [branding.business_name]
    if branding.business_address:
        footer_lines.append(branding.business_address)
    if branding.business_email:
        footer_lines.append(branding.business_email)
    pdf.multi_cell(w=pdf.epw, h=4, text="\n".join(footer_lines))
    pdf.ln(8)
    pdf.cell(0, 4, "Confirmed By,", align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 4, "Stamp", align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    if bills is not None and not bills.empty and "Bill No" in bills.columns:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 10, f"Bill-wise Details (Total Taxable: {total_taxable:,.2f})",
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        pdf.ln(5)

        pdf.set_font("Helvetica", "B", 9)
        col_widths2 = [50, 45, 45, 45]
        for i, h in enumerate(["Bill No", "Taxable", f"VAT ({branding.vat_rate_percent:g}%)", "Total"]):
            pdf.cell(col_widths2[i], 8, h, border=1, align="C")
        pdf.ln()
        pdf.set_font("Helvetica", size=9)
        for _, bill in bills.sort_values("Bill No").iterrows():
            tax = float(bill["Amount"])
            vat = tax * vat_rate
            tot = tax + vat
            pdf.cell(col_widths2[0], 8, str(bill["Bill No"]), border=1, align="C")
            pdf.cell(col_widths2[1], 8, f"{tax:,.2f}", border=1, align="C")
            pdf.cell(col_widths2[2], 8, f"{vat:,.2f}", border=1, align="C")
            pdf.cell(col_widths2[3], 8, f"{tot:,.2f}", border=1, align="C")
            pdf.ln()

    return bytes(pdf.output())


def generate_all(
    bills: pd.DataFrame, summary: pd.DataFrame, branding: LetterBranding
) -> tuple[bytes, str, int]:
    """
    Returns (file_bytes, filename, count). A single seller yields a lone PDF;
    multiple sellers are bundled into a ZIP.
    """
    outputs: list[tuple[str, bytes]] = []
    for _, row in summary.iterrows():
        seller_bills = bills[
            (bills["Seller Name"] == row["Seller Name"]) & (bills["VAT No."] == row["VAT No."])
        ]
        pdf_bytes = generate_letter_pdf(
            seller_name=row["Seller Name"],
            vat_no=row["VAT No."],
            address=row["Address"],
            total_taxable=row["Total Taxable"],
            bills=seller_bills,
            branding=branding,
        )
        fname = re.sub(r"[^A-Za-z0-9_-]+", "_", row["Seller Name"]).strip("_") + ".pdf"
        outputs.append((fname, pdf_bytes))

    if len(outputs) == 1:
        fname, data = outputs[0]
        return data, fname, 1

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, data in outputs:
            zf.writestr(fname, data)
    return buf.getvalue(), "confirm_nepal_letters.zip", len(outputs)
