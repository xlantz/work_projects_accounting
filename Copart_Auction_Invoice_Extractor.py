#!/usr/bin/env python3
"""
Copart Auction Proceed Invoice Extractor

Processes multiple Copart PDF invoices in a folder and writes a formatted Excel file.

Field Mapping:
  Loan #        → Claim# (fallback: VIN)
  Gross         → PROCEEDS FROM SALE
  Net           → CHECK PAYMENT FROM COPART
  Auct          → Gross - Net (Excel formula)
  Sell          → Sell Fee line (rare; defaults to 0)
  Recon         → TOTAL COPART SERVICE CHARGES
  Credit Memo   → Copart Lot#
  Auction House → Location from Lot# header line (e.g. "88 GA - TIFTON")
  Date          → Today (MM/DD/YYYY)
"""

import os
import re
from datetime import datetime

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Config ──────────────────────────────────────────────────────────────────

HEADERS = [
    "Loan #", "Gross", "Net", "Auct",
    "Sell", "Recon", "Credit Memo",
    "Auction House", "Date", "Note for Account"
]

COL_WIDTHS = [12, 12, 12, 12, 12, 12, 16, 25, 14, 20]

SELL_KEYWORDS = ["sell fee", "seller fee", "seller success fee"]


# ── Helpers ──────────────────────────────────────────────────────────────────

def parse_amount(raw: str) -> float:
    """Parse dollar amount strings like '$2,410.00' or '2,410.00CR' into floats."""
    raw = raw.strip().replace(",", "").replace("$", "").upper()
    credit = raw.endswith("CR")
    raw = raw.rstrip("CR").strip()
    try:
        return float(raw)
    except ValueError:
        return 0.0


def extract_text(pdf_path: str) -> str:
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    return text


# ── Parser ───────────────────────────────────────────────────────────────────

def parse_invoice(text: str) -> dict:
    data = {
        "loan_number": "",
        "vin": "",
        "lot_number": "",
        "auction_house": "",
        "gross": 0.0,
        "net": 0.0,
        "sell_fees": 0.0,
        "recon_fees": 0.0,
    }

    # Claim# → Loan #
    m = re.search(r"Claim#\s+(\d+)", text)
    if m:
        data["loan_number"] = m.group(1)

    # VIN
    m = re.search(r"Vehicle ID#\s+([A-HJ-NPR-Z0-9]{17})", text, re.IGNORECASE)
    if m:
        data["vin"] = m.group(1)

    # Fallback Loan # to VIN
    if not data["loan_number"]:
        data["loan_number"] = data["vin"]

    # Copart Lot# → Credit Memo
    m = re.search(r"Copart Lot#\s+(\S+)", text, re.IGNORECASE)
    if m:
        data["lot_number"] = m.group(1).strip()

    data["auction_house"] = "Copart"

    # Gross → PROCEEDS FROM SALE (appears as "2600.00CR")
    m = re.search(r"PROCEEDS FROM SALE[.\s]+([0-9,]+\.\d{2})CR?", text, re.IGNORECASE)
    if m:
        data["gross"] = parse_amount(m.group(1))

    # Net → CHECK PAYMENT FROM COPART
    m = re.search(r"(?:COPART CHECK#\s*\S+\s+\S+\s+([\d,]+\.\d{2})|CHECK PAYMENT FROM COPART[.\s]+([\d,]+\.\d{2}))", text, re.IGNORECASE)
    if m:
        val = m.group(1) or m.group(2)
        data["net"] = parse_amount(val)

    # Recon → TOTAL COPART SERVICE CHARGES
    m = re.search(r"TOTAL COPART SERVICE CHARGES[.\s]+([\d,]+\.\d{2})", text, re.IGNORECASE)
    if m:
        data["recon_fees"] = parse_amount(m.group(1))

    # Sell fees (rare — scan line by line)
    for line in text.split("\n"):
        line_lower = line.lower()
        if any(k in line_lower for k in SELL_KEYWORDS):
            amt_m = re.search(r"\$?([\d,]+\.\d{2})", line)
            if amt_m:
                data["sell_fees"] += parse_amount(amt_m.group(1))

    return data


# ── Excel styling ─────────────────────────────────────────────────────────────

def _thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def style_header(cell):
    cell.font = Font(bold=True, name="Arial")
    cell.fill = PatternFill("solid", start_color="D9E1F2")
    cell.alignment = Alignment(horizontal="center")
    cell.border = _thin_border()


def style_cell(cell, currency=False, is_date=False):
    cell.font = Font(name="Arial")
    cell.border = _thin_border()
    cell.alignment = Alignment(horizontal="center")
    if currency:
        cell.number_format = '$#,##0.00'
    if is_date:
        cell.number_format = 'MM/DD/YYYY'


# ── Excel writer ──────────────────────────────────────────────────────────────

def write_excel(records: list, folder: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = datetime.today().strftime("%m.%d.%y")

    # Headers
    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        style_header(cell)
        ws.column_dimensions[get_column_letter(col)].width = COL_WIDTHS[col - 1]

    # Data rows
    for i, rec in enumerate(records, start=2):
        ws.cell(i, 1, rec["loan_number"])
        ws.cell(i, 2, rec["gross"])
        ws.cell(i, 3, rec["net"])
        ws.cell(i, 4, f"=B{i}-C{i}")
        ws.cell(i, 5, rec["sell_fees"])
        ws.cell(i, 6, rec["recon_fees"])
        ws.cell(i, 7, rec["lot_number"])
        ws.cell(i, 8, rec["auction_house"])
        ws.cell(i, 9, datetime.today())
        ws.cell(i, 10, "")

        for col in range(1, 11):
            style_cell(
                ws.cell(i, col),
                currency=(col in [2, 3, 4, 5, 6]),
                is_date=(col == 9),
            )

    tag = datetime.today().strftime("%m.%d.%y")
    output = os.path.join(folder, f"CopartProceeds_{tag}.xlsx")
    wb.save(output)
    print(f"\n✅ Excel saved: {output}")
    return output


# ── PDF renamer ───────────────────────────────────────────────────────────────

def rename_pdfs(folder: str, records: list, files: list):
    print("\nRenaming PDFs...")
    for file, rec in zip(files, records):
        identifier = rec["loan_number"] or rec["vin"] or rec["lot_number"]
        if not identifier:
            print(f"⚠️  Skipping {file} (no identifier found)")
            continue
        old = os.path.join(folder, file)
        new = os.path.join(folder, f"{identifier}Copart.pdf")
        if not os.path.exists(new):
            os.rename(old, new)
            print(f"✓  {file}  →  {identifier}Copart.pdf")
        else:
            print(f"⚠️  {new} already exists, skipping rename")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("\n=== Copart Invoice Extractor ===")

    folder = input("Enter folder path containing Copart PDFs: ").strip().strip('"')

    if not os.path.isdir(folder):
        print("❌ Invalid folder path.")
        return

    pdfs = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))

    if not pdfs:
        print("❌ No PDF files found in that folder.")
        return

    records = []
    for pdf in pdfs:
        path = os.path.join(folder, pdf)
        print(f"\nProcessing: {pdf}")
        try:
            text = extract_text(path)
            rec = parse_invoice(text)
            print(
                f"  Claim#: {rec['loan_number']}  |  Lot#: {rec['lot_number']}  |  "
                f"Gross: ${rec['gross']:,.2f}  |  Net: ${rec['net']:,.2f}  |  "
                f"Recon: ${rec['recon_fees']:,.2f}"
            )
            records.append(rec)
        except Exception as e:
            print(f"  ERROR processing {pdf}: {e}")

    if not records:
        print("No records extracted.")
        return

    write_excel(records, folder)
    rename_pdfs(folder, records, pdfs)

    print("\n🎉 Done!")


if __name__ == "__main__":
    main()