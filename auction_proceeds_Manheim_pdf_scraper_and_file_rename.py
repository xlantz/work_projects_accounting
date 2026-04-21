#!/usr/bin/env python3
"""
Auction Proceed Invoice Extractor (Manheim PDFs)

Features:
- Processes multiple PDF invoices in a folder
- Extracts required fields based on your mapping
- Writes formatted Excel file
- Renames PDFs to {Loan# or VIN}Auct.pdf

Mapping:
Loan # → Lease Account No (fallback VIN)
Gross → Vehicle sale amount (absolute value)
Net → Total Payments
Auct → Gross - Net (Excel formula)
Sell → Sell Fee + Seller Success Fee + Simulcast Seller Premium Fee
Recon → All other fees
Credit Memo → Credit Memo #
Auction House → Transaction Location (cleaned)
Date → TODAY (MM/DD/YYYY)
"""

import os
import re
import sys
from datetime import datetime
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# -----------------------------
# CONFIG
# -----------------------------
SELL_FEE_KEYWORDS = [
    "sell fee",
    "seller success fee",
    "simulcast seller premium fee",
]

HEADERS = [
    "Loan #", "Gross", "Net", "Auct",
    "Sell", "Recon", "Credit Memo",
    "Auction House", "Date", "Note for Account"
]

COL_WIDTHS = [12, 12, 12, 12, 12, 12, 16, 25, 14, 20]


# -----------------------------
# HELPERS
# -----------------------------
def parse_amount(raw: str) -> float:
    raw = raw.strip().replace(",", "").replace("$", "")
    negative = raw.startswith("(") or raw.startswith("-")
    raw = raw.strip("()").lstrip("-")
    try:
        val = float(raw)
    except:
        return 0.0
    return -val if negative else val


def extract_text_from_pdf(pdf_path: str) -> str:
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    return text


def is_sell_fee(desc: str) -> bool:
    desc = desc.lower()
    return any(k in desc for k in SELL_FEE_KEYWORDS)


# -----------------------------
# CORE PARSER
# -----------------------------
def parse_invoice(text: str) -> dict:
    data = {
        "loan_number": "",
        "vin": "",
        "credit_memo": "",
        "auction_house": "",
        "gross": 0.0,
        "net": 0.0,
        "sell_fees": 0.0,
        "recon_fees": 0.0,
    }

    # --- Credit Memo ---
    m = re.search(r"Credit Memo #\s*(\d+)", text, re.IGNORECASE)
    if m:
        data["credit_memo"] = m.group(1)

    # --- Loan # ---
    m = re.search(r"LEASE ACCOUNT NO\s*(\d+)", text, re.IGNORECASE)
    if m:
        data["loan_number"] = m.group(1)

    # --- VIN ---
    m = re.search(r"VIN\s*([A-HJ-NPR-Z0-9]{17})", text)
    if m:
        data["vin"] = m.group(1)

    # --- Fallback Loan # ---
    if not data["loan_number"]:
        data["loan_number"] = data["vin"]

    # --- Auction House ---
    m = re.search(r"TRANSACTION LOCATION\s+(.+)", text, re.IGNORECASE)
    if m:
        loc = m.group(1).strip()
        loc = re.sub(r"^\d+\s+", "", loc)
        data["auction_house"] = loc

    # --- Gross (vehicle sale = negative value) ---
    gross_match = re.search(r"\(\$[\d,]+\.\d{2}\)", text)
    if gross_match:
        data["gross"] = abs(parse_amount(gross_match.group(0)))

    # --- Net (Total Payments) ---
    m = re.search(r"TOTAL PAYMENTS\s+\$?([\d,]+\.\d{2})", text, re.IGNORECASE)
    if m:
        data["net"] = float(m.group(1).replace(",", ""))

    # --- Extract Fees ---
    lines = text.split("\n")

    for line in lines:
        amt_match = re.search(r"\$[\d,]+\.\d{2}", line)
        if not amt_match:
            continue

        amount = parse_amount(amt_match.group(0))
        desc = line.lower()

        if amount < 0:
            continue  # skip gross line

        if is_sell_fee(desc):
            data["sell_fees"] += amount
        else:
            # avoid totals
            if any(x in desc for x in ["total", "payments", "amount due"]):
                continue
            data["recon_fees"] += amount

    return data


# -----------------------------
# EXCEL STYLING
# -----------------------------
def style_header(cell):
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", start_color="D9E1F2")
    cell.alignment = Alignment(horizontal="center")
    thin = Side(style="thin")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_cell(cell, currency=False, date=False):
    thin = Side(style="thin")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    cell.alignment = Alignment(horizontal="center")

    if currency:
        cell.number_format = '$#,##0.00'
    if date:
        cell.number_format = 'MM/DD/YYYY'


# -----------------------------
# WRITE EXCEL
# -----------------------------
def write_excel(records, folder):
    wb = openpyxl.Workbook()
    ws = wb.active

    tab_name = datetime.today().strftime("%m.%d.%y")
    ws.title = tab_name

    # headers
    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        style_header(cell)
        ws.column_dimensions[get_column_letter(col)].width = COL_WIDTHS[col - 1]

    for i, rec in enumerate(records, start=2):
        ws.cell(i, 1, rec["loan_number"])
        ws.cell(i, 2, rec["gross"])
        ws.cell(i, 3, rec["net"])
        ws.cell(i, 4, f"=B{i}-C{i}")
        ws.cell(i, 5, rec["sell_fees"])
        ws.cell(i, 6, rec["recon_fees"])
        ws.cell(i, 7, rec["credit_memo"])
        ws.cell(i, 8, rec["auction_house"])
        ws.cell(i, 9, datetime.today())
        ws.cell(i, 10, "")

        for col in range(1, 11):
            cell = ws.cell(i, col)
            style_cell(cell, currency=(col in [2,3,4,5,6]), date=(col==9))

    output = os.path.join(folder, f"AuctionProceeds_{tab_name}.xlsx")
    wb.save(output)
    print(f"\n✅ Excel saved: {output}")


# -----------------------------
# RENAME FILES
# -----------------------------
def rename_pdfs(folder, records, files):
    print("\nRenaming PDFs...")
    for file, rec in zip(files, records):
        identifier = rec["loan_number"] or rec["vin"]

        if not identifier:
            print(f"⚠️ Skipping {file} (no ID)")
            continue

        old = os.path.join(folder, file)
        new = os.path.join(folder, f"{identifier}Auct.pdf")

        if not os.path.exists(new):
            os.rename(old, new)
            print(f"✓ {file} → {identifier}Auct.pdf")


# -----------------------------
# MAIN
# -----------------------------
def main():
    print("\n=== Auction Invoice Extractor ===")

    folder = input("Enter folder path: ").strip().strip('"')

    if not os.path.isdir(folder):
        print("Invalid folder.")
        return

    pdfs = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]

    if not pdfs:
        print("No PDFs found.")
        return

    records = []

    for pdf in pdfs:
        path = os.path.join(folder, pdf)
        print(f"\nProcessing: {pdf}")

        try:
            text = extract_text_from_pdf(path)
            rec = parse_invoice(text)

            print(f"Loan: {rec['loan_number']} | Gross: {rec['gross']} | Net: {rec['net']}")

            records.append(rec)

        except Exception as e:
            print(f"ERROR: {e}")

    write_excel(records, folder)
    rename_pdfs(folder, records, pdfs)

    print("\n🎉 Done!")


if __name__ == "__main__":
    main()