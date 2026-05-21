"""
RDN Invoice Scraper
====================
Opens a folder-picker dialog, reads all RDN invoice files (.pdf) in the
selected folder, extracts structured data, and saves rdn_invoices.xlsx
back into that same folder.

Handles two real-world RDN file formats:
  - ZIP-based RDN  : contains 1.txt, 1.jpeg, manifest.json inside the .pdf
  - Standard PDF   : extracted with pdfplumber; two-column layout is handled

Requirements:
    pip install openpyxl pdfplumber

Run:
    python rdn_scraper.py
"""

import sys
import os
import re
import zipfile
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing Dependency",
        "Please install openpyxl:\n\n    pip install openpyxl")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Text extraction — ZIP first, then pdfplumber, then pypdf
# ---------------------------------------------------------------------------

def extract_pages(filepath: str) -> tuple:
    """
    Returns ([page_text, ...], format_hint, error_or_None).
    format_hint is 'zip' or 'pdf'.
    """
    # Strategy 1 — ZIP/RDN bundle
    if zipfile.is_zipfile(filepath):
        try:
            with zipfile.ZipFile(filepath, "r") as zf:
                txt_files = sorted(f for f in zf.namelist() if f.endswith(".txt"))
                if txt_files:
                    pages = [zf.read(n).decode("utf-8", errors="replace")
                             for n in txt_files]
                    return pages, "zip", None
        except Exception as e:
            return [], "zip", f"ZIP read failed: {e}"

    # Strategy 2 — pdfplumber
    try:
        import pdfplumber
        with pdfplumber.open(filepath) as pdf:
            pages = [p.extract_text() or "" for p in pdf.pages]
        if any(p.strip() for p in pages):
            return pages, "pdf", None
    except ImportError:
        pass
    except Exception:
        pass

    # Strategy 3 — pypdf fallback
    try:
        from pypdf import PdfReader
        reader = PdfReader(filepath)
        pages = [p.extract_text() or "" for p in reader.pages]
        if any(p.strip() for p in pages):
            return pages, "pdf", None
    except Exception as e:
        return [], "pdf", f"pypdf failed: {e}"

    return [], "pdf", (
        "Could not read file. Install pdfplumber:  pip install pdfplumber"
    )


# ---------------------------------------------------------------------------
# Line-item parser  (shared by both formats)
# ---------------------------------------------------------------------------

# Matches a data row: date [optional inline service] qty $rate tax tax $subtotal
# Tax rate and tax amount may be "n/a"
_ITEM_RE = re.compile(
    r"(\d{1,2}/\d{1,2}/\d{2,4})\s*"   # date  (no space required after — PDF quirk)
    r"([^$\d\n]*?)\s*"                  # optional inline service name
    r"(\d+)\s+"                          # qty
    r"\$([\d,\.]+)\s+"                   # rate
    r"([\d\.]+%|n/a)\s+"                # tax rate  (or n/a)
    r"(\$?[\d,\.]+|n/a)\s+"             # tax amount (or n/a)
    r"\$([\d,\.]+)",                     # subtotal
    re.I
)


def _parse_items(lines: list, header_idx: int) -> list:
    """
    Extract line items starting after header_idx.
    Handles:
      - service names on the line before the data row
      - service names on the line after  the data row
      - date + service merged with no space (PDF column artefact)
      - n/a tax values
    """
    section = []
    for l in lines[header_idx + 1:]:
        if re.search(
            r"Account Information|Sales Tax|Tax ID|Total Due|file://|"
            r"appreciate your business",
            l, re.I
        ):
            break
        section.append(l.strip())

    items = []
    i = 0
    while i < len(section):
        line = section[i]
        m = _ITEM_RE.match(line)
        if m:
            service = m.group(2).strip()

            # Service name may be on the PREVIOUS line (PDF split)
            if not service and i > 0:
                prev = section[i - 1]
                if prev and not _ITEM_RE.match(prev):
                    service = prev

            # Service name may continue on the NEXT line (PDF split)
            advance = 1
            if i + 1 < len(section):
                nxt = section[i + 1]
                if (nxt
                        and not _ITEM_RE.match(nxt)
                        and not re.search(
                            r"Account Information|Sales Tax|Tax ID|Total Due|file://",
                            nxt, re.I)):
                    service = (service + " " + nxt).strip()
                    advance = 2       # skip the continuation line

            tax_rate   = m.group(5).strip()
            tax_amount = m.group(6).strip().lstrip("$") if m.group(6).lower() != "n/a" else ""

            items.append({
                "service_date": m.group(1),
                "service_type": service,
                "qty":          m.group(3),
                "rate":         m.group(4),
                "tax_rate":     tax_rate if tax_rate.lower() != "n/a" else "",
                "tax_amount":   tax_amount,
                "subtotal":     m.group(7),
            })
            i += advance
        else:
            i += 1

    return items


# ---------------------------------------------------------------------------
# Main invoice parser
# ---------------------------------------------------------------------------

def _find(pattern, text, group=1, default=""):
    m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
    return m.group(group).strip() if m else default


def parse_rdn_text(raw: str, fmt: str = "zip") -> dict:
    """
    Parse one page of RDN invoice text.
    fmt: 'zip' (clean line-per-field) or 'pdf' (two-column merge artefacts).
    """
    lines = [l.rstrip() for l in raw.replace("\r\n", "\n").split("\n")]
    text  = "\n".join(lines)
    data  = {}

    # ── Universal regex fields (work in both formats) ──────────────────────
    data["invoice_number"] = _find(r"Invoice[ \t]+Number[ \t]*:?[ \t]*(\S+)", text)
    data["invoice_date"]   = _find(r"Invoice Date[:\s]+([^\r\n]+)", text)
    data["comp_ref"]       = _find(r"Comp Ref\s*#[:\s]+(\S+)", text)
    data["account_number"] = _find(r"Account\s*#[:\s]+(\S+)", text)
    data["collateral"]     = _find(r"Collateral[:\s]+([^\r\n]+)", text)
    data["vin"]            = _find(r"V\.?I\.?N\.?[:\s]+(\S+)", text)
    data["tax_id"]         = _find(r"Tax ID\s*#[:\s]+([\S]+)", text)
    data["sales_tax"]      = _find(r"Sales Tax[:\s]+\$?([\d,\.]+)", text)

    # Total Due — two layouts:
    #   ZIP: "Total Due Upon\nReceipt: $135.31"   (amount after Receipt:)
    #   PDF: "Total Due Upon\nTax ID... $850.00\nReceipt:" (amount before Receipt:)
    m_total = re.search(
        r"Total Due[\s\S]{0,60}?\$([\d,\.]+)",
        text, re.I
    )
    data["total_due"] = m_total.group(1).strip() if m_total else ""

    # ── Sender company ─────────────────────────────────────────────────────
    if fmt == "zip":
        # ZIP: company name is the line immediately after "Comp Ref #:"
        comp_ref_idx = next(
            (i for i, l in enumerate(lines) if re.search(r"Comp Ref", l, re.I)), None
        )
        if comp_ref_idx is not None and comp_ref_idx + 1 < len(lines):
            data["sender_company"] = lines[comp_ref_idx + 1].strip()
            addr_parts = []
            for l in lines[comp_ref_idx + 2:]:
                if re.match(r"Phone:", l, re.I):
                    break
                if l.strip():
                    addr_parts.append(l.strip())
            data["sender_address"] = ", ".join(addr_parts)
        else:
            data["sender_company"] = ""
            data["sender_address"] = ""
    else:
        # PDF: company name sits between the timestamp and "Invoice Number"
        inv_idx = next(
            (i for i, l in enumerate(lines)
             if re.match(r"Invoice Number", l, re.I)), None
        )
        data["sender_company"] = ""
        data["sender_address"] = ""
        if inv_idx and inv_idx > 0:
            for l in reversed(lines[:inv_idx]):
                s = l.strip()
                if s and not re.match(r"[\d/]+,\s*[\d:]+\s*[AP]M", s, re.I):
                    data["sender_company"] = s
                    break
            # Address: lines between company and "Invoice Number"
            comp_line = next(
                (i for i, l in enumerate(lines)
                 if l.strip() == data["sender_company"]), None
            )
            if comp_line is not None:
                addr_parts = []
                for l in lines[comp_line + 1: inv_idx]:
                    if l.strip():
                        addr_parts.append(l.strip())
                data["sender_address"] = ", ".join(addr_parts)

    # ── Clean sender_company of any leftover "Invoice Number" text ──────────
    # Guards against: "Copart Inc. Invoice Number 98765" ending up as company.
    if data.get("sender_company"):
        data["sender_company"] = re.sub(
            r"\s*Invoice\s+Number\s*:?\s*\S*", "",
            data["sender_company"], flags=re.I
        ).strip()
        # Strip a trailing standalone number only when the name ends in a
        # letter or period — preserves "1st Adjusters" and "Copart Inc."
        m_trail = re.match(r"^(.*?)\s+(\d\S*)\s*$",
                           data["sender_company"])
        if m_trail:
            data["sender_company"] = m_trail.group(1).strip()
            if not data.get("invoice_number"):
                data["invoice_number"] = m_trail.group(2).strip()

    # ── Sender phone / fax ─────────────────────────────────────────────────
    # Grab only the FIRST Phone: / Fax: (the sender's; bill-to usually has blank ones)
    data["sender_phone"] = _find(r"Phone:\s*(\S[^\r\n]+)", text)
    data["sender_fax"]   = _find(r"Fax:\s*(\S[^\r\n]+)", text)

    # ── Debtor ─────────────────────────────────────────────────────────────
    # Use [ \t]* so we don't consume the newline after "Debtor:" (PDF quirk)
    debtor_inline = _find(r"Debtor:[ \t]*(\S[^\r\n]*)", text)
    if debtor_inline:
        data["debtor"] = debtor_inline
    elif fmt == "pdf":
        # PDF two-column mixing puts the debtor name after "ATTN:" on the same line
        data["debtor"] = _find(r"ATTN:[ \t]*(\S[^\r\n]*)", text)
    else:
        data["debtor"] = ""

    # ── Bill-to / ATTN ─────────────────────────────────────────────────────
    if fmt == "zip":
        attn_idx = next(
            (i for i, l in enumerate(lines) if re.match(r"ATTN:", l, re.I)), None
        )
        if attn_idx is not None:
            attn_parts = []
            for l in lines[attn_idx + 1:]:
                if re.match(r"Account\s*#", l, re.I):
                    break
                if l.strip():
                    attn_parts.append(l.strip())
            data["bill_to"] = " | ".join(attn_parts)
        else:
            data["bill_to"] = ""
    else:
        # PDF: bill-to address lines follow "ATTN:" but have collateral merged in;
        # grab lines between ATTN and Account # that don't contain known field labels
        attn_idx = next(
            (i for i, l in enumerate(lines) if re.match(r"ATTN:", l, re.I)), None
        )
        acct_idx = next(
            (i for i, l in enumerate(lines)
             if re.match(r"Account\s*#", l, re.I)), None
        )
        bill_parts = []
        if attn_idx is not None and acct_idx is not None:
            for l in lines[attn_idx + 1: acct_idx]:
                clean = re.sub(
                    r"Collateral[:\s]+.*|V\.?I\.?N\.?[:\s]+\S+|Debtor[:\s]*",
                    "", l, flags=re.I
                ).strip()
                if clean and not re.match(r"(Phone|Fax):", clean, re.I):
                    bill_parts.append(clean)
        data["bill_to"] = " | ".join(p for p in bill_parts if p)

    # ── Account info entity (e.g. "Arivo Acceptance LLC") ─────────────────
    acct_info_idx = next(
        (i for i, l in enumerate(lines)
         if re.search(r"Account Information", l, re.I)), None
    )
    data["account_info_entity"] = (
        lines[acct_info_idx].replace("Account Information", "").strip()
        if acct_info_idx is not None else ""
    )

    # ── Line items ─────────────────────────────────────────────────────────
    header_idx = next(
        (i for i, l in enumerate(lines)
         if re.search(r"Date\s+Service\s+QTY\s+Rate", l, re.I)), None
    )
    data["line_items"] = _parse_items(lines, header_idx) if header_idx is not None else []

    return data


# ---------------------------------------------------------------------------
# File reader
# ---------------------------------------------------------------------------

def read_rdn_file(filepath: str) -> tuple:
    """Returns ([records], error_or_None)."""
    pages, fmt, err = extract_pages(filepath)
    if err:
        return [], err
    if not pages:
        return [], "No text content found."

    records = []
    fname = os.path.basename(filepath)
    for page_text in pages:
        if not page_text.strip():
            continue
        parsed = parse_rdn_text(page_text, fmt=fmt)
        parsed["source_file"] = fname
        records.append(parsed)

    if not records:
        return [], "Readable but no RDN invoice data found."
    return records, None


# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------

HEADER_COLOR  = "1F4E79"
ALT_ROW_COLOR = "D6E4F0"
BORDER_COLOR  = "BFBFBF"


def _thin_border():
    s = Side(style="thin", color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)


def _write_header_row(ws, headers, row):
    hf = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    hfl = PatternFill("solid", fgColor=HEADER_COLOR)
    ha = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = hf; c.fill = hfl; c.alignment = ha; c.border = _thin_border()
    ws.row_dimensions[row].height = 30


def _autofit(ws, headers, min_w=12, max_w=45):
    for col, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(col)].width = max(
            min_w, min(max_w, len(h) + 4))


def _num(val, cast=float):
    try:
        return cast(str(val).replace(",", "").replace("$", "").strip())
    except (ValueError, AttributeError):
        return val


def write_excel(all_records: list, output_path: str):
    wb = Workbook()

    # ── Sheet 1: Invoice Summary ───────────────────────────────────────────
    ws = wb.active
    ws.title = "Invoice Summary"
    sum_hdrs = [
        "Source File", "Invoice Number", "Invoice Date",
        "Comp Ref #", "Sender Company", "Sender Address", "Sender Phone",
        "Bill To", "Account #", "Debtor", "Collateral", "VIN",
        "Account Info Entity", "Tax ID", "Sales Tax ($)", "Total Due ($)",
    ]
    _write_header_row(ws, sum_hdrs, 1)

    for row, rec in enumerate(all_records, 2):
        fill = PatternFill("solid", fgColor=ALT_ROW_COLOR) if row % 2 == 0 else None
        vals = [
            rec.get("source_file",""),   rec.get("invoice_number",""),
            rec.get("invoice_date",""),
            rec.get("comp_ref",""),       rec.get("sender_company",""),
            rec.get("sender_address",""), rec.get("sender_phone",""),
            rec.get("bill_to",""),        rec.get("account_number",""),
            rec.get("debtor",""),         rec.get("collateral",""),
            rec.get("vin",""),            rec.get("account_info_entity",""),
            rec.get("tax_id",""),
            _num(rec.get("sales_tax","")),
            _num(rec.get("total_due","")),
        ]
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.border = _thin_border()
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(wrap_text=False, vertical="center")
            if fill: c.fill = fill
            if col >= 16: c.number_format = "#,##0.00"

    _autofit(ws, sum_hdrs)
    ws.freeze_panes = "A2"

    # ── Sheet 2: Line Items ────────────────────────────────────────────────
    ws2 = wb.create_sheet("Line Items")
    item_hdrs = [
        "Source File", "Invoice Number", "Invoice Date",
        "Sender Company", "Account #", "Debtor", "Collateral", "VIN",
        "Service Date", "Service Type", "QTY", "Rate ($)",
        "Tax Rate", "Tax Amount ($)", "Subtotal ($)",
        "Invoice Total Due ($)",
    ]
    _write_header_row(ws2, item_hdrs, 1)

    row2 = 2
    for rec in all_records:
        for item in rec.get("line_items", []):
            fill = PatternFill("solid", fgColor=ALT_ROW_COLOR) if row2 % 2 == 0 else None
            vals = [
                rec.get("source_file",""),    rec.get("invoice_number",""),
                rec.get("invoice_date",""),
                rec.get("sender_company",""),  rec.get("account_number",""),
                rec.get("debtor",""),
                rec.get("collateral",""),      rec.get("vin",""),
                item.get("service_date",""),   item.get("service_type",""),
                _num(item.get("qty",""), int), _num(item.get("rate","")),
                item.get("tax_rate",""),       _num(item.get("tax_amount","")),
                _num(item.get("subtotal","")), _num(rec.get("total_due","")),
            ]
            for col, val in enumerate(vals, 1):
                c = ws2.cell(row=row2, column=col, value=val)
                c.border = _thin_border()
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(wrap_text=False, vertical="center")
                if fill: c.fill = fill
                if col in (13, 15, 16, 17): c.number_format = "#,##0.00"
            row2 += 1

    _autofit(ws2, item_hdrs)
    ws2.freeze_panes = "A2"

    wb.save(output_path)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RDN Invoice Scraper")
        self.resizable(False, False)
        self.configure(bg="#f0f4f8")
        self._build_ui()
        self._center()

    def _build_ui(self):
        BG = "#f0f4f8"; DARK = "#1F4E79"; BLUE = "#2E75B6"; GREEN = "#1F7A1F"

        hdr = tk.Frame(self, bg=DARK)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  RDN Invoice Scraper", bg=DARK, fg="white",
                 font=("Arial", 14, "bold"), pady=12).pack(side="left")

        body = tk.Frame(self, bg=BG, padx=16, pady=16)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="Invoice Folder:", bg=BG,
                 font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0,4))

        fr = tk.Frame(body, bg=BG)
        fr.grid(row=1, column=0, sticky="ew", pady=(0,12))
        self.folder_var = tk.StringVar()
        tk.Entry(fr, textvariable=self.folder_var, width=46,
                 font=("Arial", 10), relief="solid", bd=1).pack(side="left", padx=(0,6))
        tk.Button(fr, text="Browse…", bg=BLUE, fg="white", relief="flat",
                  font=("Arial", 10), padx=10, pady=4, cursor="hand2",
                  command=self._browse).pack(side="left")

        tk.Button(body, text="▶   Extract & Export to Excel",
                  bg=GREEN, fg="white", relief="flat",
                  font=("Arial", 11, "bold"), padx=14, pady=8,
                  cursor="hand2", activebackground="#155015",
                  command=self._run).grid(row=2, column=0, sticky="ew", pady=(0,12))

        self.progress = ttk.Progressbar(body, mode="determinate", length=420)
        self.progress.grid(row=3, column=0, sticky="ew", pady=(0,6))

        self.status_var = tk.StringVar(value="Select a folder to get started.")
        tk.Label(body, textvariable=self.status_var, bg=BG,
                 font=("Arial", 9), fg="#444", wraplength=420,
                 justify="left").grid(row=4, column=0, sticky="w", pady=(0,8))

        tk.Label(body, text="Log:", bg=BG,
                 font=("Arial", 9, "bold")).grid(row=5, column=0, sticky="w", pady=(4,2))

        lf = tk.Frame(body, bg=BG)
        lf.grid(row=6, column=0, sticky="nsew")
        self.log_box = tk.Text(lf, height=12, width=62, font=("Courier", 9),
                               state="disabled", relief="solid", bd=1, bg="#ffffff")
        sb = tk.Scrollbar(lf, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=sb.set)
        self.log_box.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _browse(self):
        folder = filedialog.askdirectory(title="Select folder containing RDN invoices")
        if folder:
            self.folder_var.set(folder)
            self.status_var.set("Folder selected. Click Extract to begin.")

    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _run(self):
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showwarning("No Folder Selected", "Please select a folder first.")
            return
        if not os.path.isdir(folder):
            messagebox.showerror("Invalid Folder", f"Folder not found:\n{folder}")
            return

        # Deduplicate on Windows (case-insensitive FS)
        seen, files = set(), []
        for pat in ("*.pdf", "*.PDF"):
            for f in glob.glob(os.path.join(folder, pat)):
                n = os.path.normcase(f)
                if n not in seen:
                    seen.add(n); files.append(f)
        files.sort()

        if not files:
            messagebox.showinfo("No Files Found",
                                "No .pdf files found in the selected folder.")
            return

        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.progress["value"] = 0
        self.progress["maximum"] = len(files)

        self._log(f"Folder : {folder}")
        self._log(f"Files  : {len(files)} found\n")

        all_records, skipped = [], 0
        file_record_map = {}   # filepath -> first record (for renaming)

        for i, filepath in enumerate(files, 1):
            fname = os.path.basename(filepath)
            self.status_var.set(f"Processing {i}/{len(files)}: {fname}")
            records, err = read_rdn_file(filepath)
            if records:
                items = sum(len(r["line_items"]) for r in records)
                self._log(f"  ✓  {fname}  ({len(records)} page(s), {items} line item(s))")
                all_records.extend(records)
                # Use the first record that has line items — the same source
                # as the Line Items tab — so rename values always match the sheet.
                source_rec = next(
                    (r for r in records if r.get("line_items")),
                    records[0]
                )
                file_record_map[filepath] = source_rec
            else:
                self._log(f"  ✗  {fname}  — skipped")
                self._log(f"       Reason: {err}")
                skipped += 1
            self.progress["value"] = i

        if not all_records:
            self.status_var.set("No valid RDN records found.")
            messagebox.showwarning("Nothing Extracted",
                "No valid RDN invoice data could be extracted.\n"
                "Check the Log for details.")
            return

        output_path = os.path.join(folder, "rdn_invoices.xlsx")
        try:
            write_excel(all_records, output_path)
        except PermissionError:
            messagebox.showerror("File In Use",
                "Could not save rdn_invoices.xlsx — close it in Excel and try again.")
            self.status_var.set("Export failed — close the file in Excel and retry.")
            return

        # ── Rename each invoice file ────────────────────────────────────────
        self._log("\nRenaming files...")
        renamed, rename_errors = 0, 0
        for filepath, rec in file_record_map.items():
            # Pull account # and invoice # from the record that fed the
            # Line Items tab — the user confirmed those values are correct.
            line_item_rec = next(
                (r for r in all_records
                 if r.get("source_file") == os.path.basename(filepath)
                 and r.get("line_items")),
                rec   # fallback to the stored record
            )
            company = line_item_rec.get("sender_company", "").strip()
            account = line_item_rec.get("account_number", "").strip()
            invoice = line_item_rec.get("invoice_number", "").strip()
            # Sanitise: remove characters illegal in Windows filenames
            def _safe(s):
                return re.sub(r'[<>:"/\\|?*]', '', s).strip() or "Unknown"
            new_name = (f"{_safe(company)} - Loan # {_safe(account)}"
                        f" - Inv # {_safe(invoice)}.pdf")
            new_path = os.path.join(folder, new_name)
            old_name = os.path.basename(filepath)
            if os.path.normcase(filepath) == os.path.normcase(new_path):
                self._log(f"  —  {old_name}  (already correctly named)")
                continue
            try:
                # If target already exists, add a suffix to avoid overwriting
                if os.path.exists(new_path):
                    base, ext = os.path.splitext(new_name)
                    new_path = os.path.join(folder, f"{base} (dup){ext}")
                os.rename(filepath, new_path)
                self._log(f"  ✓  {old_name}")
                self._log(f"     → {new_name}")
                renamed += 1
            except Exception as e:
                self._log(f"  ✗  Could not rename {old_name}: {e}")
                rename_errors += 1

        total_items = sum(len(r["line_items"]) for r in all_records)
        summary = (f"Done!  {len(all_records)} invoice(s), {total_items} line item(s)"
                   + (f", {skipped} skipped" if skipped else "")
                   + (f", {renamed} renamed" if renamed else ""))
        self.status_var.set(summary)
        self._log(f"\n✓ Saved: {output_path}")
        self._log(f"  Invoices   : {len(all_records)}")
        self._log(f"  Line items : {total_items}")
        self._log(f"  Renamed    : {renamed}" + (f"  ({rename_errors} errors)" if rename_errors else ""))
        if skipped:
            self._log(f"  Skipped    : {skipped}")

        if messagebox.askyesno("Export Complete",
                f"Excel file saved and files renamed!\n\n{output_path}\n\nOpen the folder now?"):
            import subprocess, platform
            if platform.system() == "Windows":
                os.startfile(folder)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])


if __name__ == "__main__":
    App().mainloop()