"""
Confirm Nepal (Desktop) — generate VAT confirmation letters from a bill list,
entirely offline. No server, no account, no internet required at runtime.
"""
from __future__ import annotations

import os
import sys
import threading
import traceback
import zipfile
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from generator import LetterBranding, TemplateError, generate_letter_pdf, parse_workbook, summarize_sellers

# ------------------------------------------------------------------ paths
def resource_path(relative: str) -> str:
    """Works both when run as a script and when frozen by PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative)
    return os.path.join(os.path.abspath(os.path.dirname(__file__)), relative)


TEMPLATE_PATH = resource_path(os.path.join("static", "confirm_nepal_template.xlsx"))

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

CRIMSON = "#a8283b"
CRIMSON_DARK = "#7d1e2b"
INDIGO = "#122c56"
PAPER = "#fefaf6"
CARD = "#ffffff"
LINE = "#e7e0d8"
MUTED = "#78716c"
INK = "#1c1917"


class LabeledEntry(ctk.CTkFrame):
    def __init__(self, master, label, placeholder="", optional=False, **kwargs):
        super().__init__(master, fg_color="transparent")
        text = label + ("  (optional)" if optional else "")
        ctk.CTkLabel(self, text=text, font=("Arial", 12, "bold"), text_color=INK,
                     anchor="w").pack(fill="x", pady=(0, 4))
        self.entry = ctk.CTkEntry(self, placeholder_text=placeholder, fg_color="#fdfbf8",
                                   border_color=LINE, text_color=INK, **kwargs)
        self.entry.pack(fill="x")

    def get(self) -> str:
        return self.entry.get().strip()

    def set(self, value: str):
        self.entry.delete(0, "end")
        self.entry.insert(0, value)


class FilePicker(ctk.CTkFrame):
    def __init__(self, master, label, filetypes, optional=True):
        super().__init__(master, fg_color="transparent")
        text = label + ("  (optional)" if optional else "")
        ctk.CTkLabel(self, text=text, font=("Arial", 12, "bold"), text_color=INK,
                     anchor="w").pack(fill="x", pady=(0, 4))
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x")
        self.path_var = ctk.StringVar(value="")
        self.path_label = ctk.CTkLabel(row, textvariable=self.path_var, text_color=MUTED,
                                        anchor="w", font=("Arial", 11))
        self.path_label.pack(side="left", fill="x", expand=True)
        self.filetypes = filetypes
        ctk.CTkButton(row, text="Browse…", width=90, fg_color=INDIGO, hover_color="#0c1f3d",
                       command=self._browse).pack(side="right")

    def _browse(self):
        path = filedialog.askopenfilename(filetypes=self.filetypes)
        if path:
            self.path_var.set(os.path.basename(path))
            self._full_path = path

    def get_path(self) -> str | None:
        return getattr(self, "_full_path", None)

    def get_bytes(self) -> bytes | None:
        p = self.get_path()
        if not p:
            return None
        with open(p, "rb") as f:
            return f.read()


class ConfirmNepalApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Confirm Nepal — VAT Confirmation Letters")
        self.geometry("880x900")
        self.configure(fg_color=PAPER)

        self._build_header()
        self._build_scroll_area()
        self._build_form()
        self._build_footer()

    # --------------------------------------------------------------- layout
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=CARD, height=64, corner_radius=0,
                               border_width=0)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        bar = ctk.CTkFrame(header, fg_color=CRIMSON, height=3, corner_radius=0)
        bar.pack(fill="x", side="bottom")
        title = ctk.CTkLabel(header, text="Confirm Nepal", font=("Arial", 18, "bold"),
                              text_color=INDIGO)
        title.pack(side="left", padx=24)

    def _build_scroll_area(self):
        self.scroll = ctk.CTkScrollableFrame(self, fg_color=PAPER)
        self.scroll.pack(fill="both", expand=True, padx=0, pady=0)

    def _card(self, parent, title=None, subtitle=None):
        card = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=12, border_width=1,
                             border_color=LINE)
        card.pack(fill="x", padx=40, pady=14)
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=24, pady=20)
        if title:
            ctk.CTkLabel(inner, text=title, font=("Arial", 15, "bold"),
                         text_color=INK, anchor="w").pack(fill="x")
        if subtitle:
            ctk.CTkLabel(inner, text=subtitle, font=("Arial", 11), text_color=MUTED,
                         anchor="w", justify="left", wraplength=760).pack(fill="x", pady=(2, 14))
        return inner

    def _section_label(self, parent, text):
        sep = ctk.CTkFrame(parent, fg_color=LINE, height=1)
        sep.pack(fill="x", pady=(16, 8))
        ctk.CTkLabel(parent, text=text.upper(), font=("Arial", 10, "bold"),
                     text_color=CRIMSON, anchor="w").pack(fill="x", pady=(0, 8))

    def _row(self, parent):
        r = ctk.CTkFrame(parent, fg_color="transparent")
        r.pack(fill="x", pady=6)
        return r

    def _build_form(self):
        # --- Step 1: template
        c1 = self._card(
            self.scroll,
            "1. Get the bill list template",
            "Fill this in with your bills, one row per bill. Don't rename the column headers.",
        )
        ctk.CTkButton(c1, text="⬇ Save Excel template…", fg_color="transparent",
                      border_width=1, border_color=CRIMSON, text_color=CRIMSON,
                      hover_color="#fbeceb", command=self._save_template).pack(anchor="w")

        # --- Step 2: business details
        c2 = self._card(self.scroll, "2. Your business details",
                         "These appear on the letter itself.")

        row = self._row(c2)
        self.business_name = LabeledEntry(row, "Business name",
                                           "Shear Engineering and Builders Pvt. Ltd.")
        self.business_name.pack(fill="x")

        row = self._row(c2)
        left = ctk.CTkFrame(row, fg_color="transparent"); left.pack(side="left", fill="x", expand=True, padx=(0, 8))
        right = ctk.CTkFrame(row, fg_color="transparent"); right.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.business_address = LabeledEntry(left, "Business address", "Harinagar-2, Sunsari", optional=True)
        self.business_address.pack(fill="x")
        self.business_email = LabeledEntry(right, "Business email", "you@company.com", optional=True)
        self.business_email.pack(fill="x")

        self._section_label(c2, "Letter wording")
        row = self._row(c2)
        a = ctk.CTkFrame(row, fg_color="transparent"); a.pack(side="left", fill="x", expand=True, padx=(0, 6))
        b = ctk.CTkFrame(row, fg_color="transparent"); b.pack(side="left", fill="x", expand=True, padx=6)
        cc = ctk.CTkFrame(row, fg_color="transparent"); cc.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.fiscal_year = LabeledEntry(a, "Fiscal year (as printed)", "2081/82", optional=True)
        self.fiscal_year.pack(fill="x")
        self.period_end = LabeledEntry(b, "Period end (as printed)", "the end of Ashad 2082", optional=True)
        self.period_end.pack(fill="x")
        self.letter_date = LabeledEntry(cc, "Letter date (as printed)", "2082/06/11", optional=True)
        self.letter_date.pack(fill="x")

        row = self._row(c2)
        a = ctk.CTkFrame(row, fg_color="transparent"); a.pack(side="left", fill="x", expand=True, padx=(0, 6))
        b = ctk.CTkFrame(row, fg_color="transparent"); b.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.vat_rate = LabeledEntry(a, "VAT rate (%)", "13")
        self.vat_rate.set("13")
        self.vat_rate.pack(fill="x")
        self.threshold = LabeledEntry(b, "Threshold (generate above this)", "100000")
        self.threshold.set("100000")
        self.threshold.pack(fill="x")

        self._section_label(c2, "Branding")
        row = self._row(c2)
        a = ctk.CTkFrame(row, fg_color="transparent"); a.pack(side="left", fill="x", expand=True, padx=(0, 6))
        b = ctk.CTkFrame(row, fg_color="transparent"); b.pack(side="left", fill="x", expand=True, padx=6)
        cc = ctk.CTkFrame(row, fg_color="transparent"); cc.pack(side="left", fill="x", expand=True, padx=(6, 0))
        img_types = [("Image", "*.png *.jpg *.jpeg")]
        self.letterhead_picker = FilePicker(a, "Letterhead image", img_types)
        self.letterhead_picker.pack(fill="x")
        self.signature_picker = FilePicker(b, "Signature image", img_types)
        self.signature_picker.pack(fill="x")
        self.stamp_picker = FilePicker(cc, "Stamp image", img_types)
        self.stamp_picker.pack(fill="x")

        # --- Step 3: bill list + output
        c3 = self._card(self.scroll, "3. Upload your filled bill list")
        self.bills_picker = FilePicker(c3, "Bill list (.xlsx)", [("Excel", "*.xlsx")], optional=False)
        self.bills_picker.pack(fill="x")

        c4 = self._card(self.scroll, "4. Choose where to save the letters")
        row = ctk.CTkFrame(c4, fg_color="transparent")
        row.pack(fill="x")
        self.output_var = ctk.StringVar(value=str(Path.home() / "Confirm Nepal Letters"))
        self.output_label = ctk.CTkLabel(row, textvariable=self.output_var, text_color=MUTED,
                                          anchor="w", font=("Arial", 11))
        self.output_label.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(row, text="Choose folder…", width=110, fg_color=INDIGO,
                      hover_color="#0c1f3d", command=self._choose_output).pack(side="right")

        # --- Generate button + status
        gen_frame = ctk.CTkFrame(self.scroll, fg_color="transparent")
        gen_frame.pack(fill="x", padx=40, pady=(6, 4))
        self.generate_btn = ctk.CTkButton(gen_frame, text="Generate confirmation letters",
                                           fg_color=CRIMSON, hover_color=CRIMSON_DARK,
                                           font=("Arial", 13, "bold"), height=42,
                                           command=self._on_generate)
        self.generate_btn.pack(fill="x")
        self.status_var = ctk.StringVar(value="")
        self.status_label = ctk.CTkLabel(gen_frame, textvariable=self.status_var,
                                          text_color=MUTED, anchor="w", justify="left",
                                          wraplength=760, font=("Arial", 11))
        self.status_label.pack(fill="x", pady=(8, 0))

        privacy = ctk.CTkFrame(self.scroll, fg_color="#f4efe8", corner_radius=10)
        privacy.pack(fill="x", padx=40, pady=(6, 20))
        ctk.CTkLabel(
            privacy,
            text="Runs entirely on this computer. Your bill list, business details, and letters "
                 "never leave this machine — nothing is uploaded anywhere.",
            text_color=MUTED, font=("Arial", 11), wraplength=760, justify="left",
        ).pack(padx=16, pady=12, anchor="w")

    def _build_footer(self):
        footer = ctk.CTkLabel(self, text="Confirm Nepal — offline VAT confirmation letters",
                               text_color=MUTED, font=("Arial", 10))
        footer.pack(pady=(0, 10))

    # --------------------------------------------------------------- actions
    def _save_template(self):
        dest = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="confirm_nepal_template.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not dest:
            return
        with open(TEMPLATE_PATH, "rb") as src, open(dest, "wb") as out:
            out.write(src.read())
        messagebox.showinfo("Saved", f"Template saved to:\n{dest}")

    def _choose_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_var.set(folder)

    def _set_status(self, text, error=False):
        self.status_var.set(text)
        self.status_label.configure(text_color=CRIMSON_DARK if error else "#166534")

    def _on_generate(self):
        bills_path = self.bills_picker.get_path()
        if not bills_path:
            self._set_status("Please choose a filled-in bill list (.xlsx) first.", error=True)
            return
        if not self.business_name.get():
            self._set_status("Please enter your business name.", error=True)
            return

        try:
            vat_rate = float(self.vat_rate.get() or 13)
            threshold = float(self.threshold.get() or 100000)
        except ValueError:
            self._set_status("VAT rate and threshold must be numbers.", error=True)
            return

        self.generate_btn.configure(state="disabled", text="Generating…")
        self._set_status("Working…")
        thread = threading.Thread(target=self._generate_worker, args=(bills_path, vat_rate, threshold), daemon=True)
        thread.start()

    def _generate_worker(self, bills_path, vat_rate, threshold):
        try:
            with open(bills_path, "rb") as f:
                bills_bytes = f.read()

            parsed = parse_workbook(bills_bytes)
            summary = summarize_sellers(parsed.bills, threshold=threshold)
            if summary.empty:
                self.after(0, self._set_status,
                           f"No sellers crossed the threshold of {threshold:,.2f}. "
                           f"Check your Amount column or lower the threshold.", True)
                return

            branding = LetterBranding(
                business_name=self.business_name.get(),
                business_address=self.business_address.get(),
                business_email=self.business_email.get(),
                fiscal_year_label=self.fiscal_year.get(),
                period_end_label=self.period_end.get(),
                letter_date_label=self.letter_date.get(),
                vat_rate_percent=vat_rate,
                threshold=threshold,
                letterhead_bytes=self.letterhead_picker.get_bytes(),
                signature_bytes=self.signature_picker.get_bytes(),
                stamp_bytes=self.stamp_picker.get_bytes(),
            )

            out_dir = Path(self.output_var.get())
            out_dir.mkdir(parents=True, exist_ok=True)

            count = 0
            for _, row in summary.iterrows():
                seller_bills = parsed.bills[
                    (parsed.bills["Seller Name"] == row["Seller Name"])
                    & (parsed.bills["VAT No."] == row["VAT No."])
                ]
                pdf_bytes = generate_letter_pdf(
                    seller_name=row["Seller Name"],
                    vat_no=row["VAT No."],
                    address=row["Address"],
                    total_taxable=row["Total Taxable"],
                    bills=seller_bills,
                    branding=branding,
                )
                import re
                fname = re.sub(r"[^A-Za-z0-9_-]+", "_", row["Seller Name"]).strip("_") + ".pdf"
                (out_dir / fname).write_bytes(pdf_bytes)
                count += 1

            warn = (" " + " ".join(parsed.warnings)) if parsed.warnings else ""
            self.after(0, self._set_status,
                       f"✅ Generated {count} letter(s) in {out_dir}.{warn}", False)
            self.after(0, lambda: messagebox.showinfo(
                "Done", f"Generated {count} confirmation letter(s) in:\n{out_dir}"))
        except TemplateError as e:
            self.after(0, self._set_status, str(e), True)
        except Exception:
            err = traceback.format_exc()
            self.after(0, self._set_status, "Unexpected error — see details.", True)
            self.after(0, lambda: messagebox.showerror("Error", err))
        finally:
            self.after(0, lambda: self.generate_btn.configure(
                state="normal", text="Generate confirmation letters"))


if __name__ == "__main__":
    app = ConfirmNepalApp()
    app.mainloop()
