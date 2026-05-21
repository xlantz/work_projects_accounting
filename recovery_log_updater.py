"""
recovery_log_updater.py

Scans all date-named tabs in the bank deposit workbook and copies rows into
the Recovery log tab:
  1. Any row whose Recovery - 11901 column (col H) has a non-zero value.
  2. Every data row inside any "Auction Proceeds" section.

Layout written to the Recovery tab:
  Row 2  – summation row (preserved from original)
  Row 3  – column headers (written fresh each run)
  Row 4+ – data rows

Changes in this version:
  • Column headers are written to row 3 on every run.
  • "see image above" batch-label fix: the code now scans backward all
    the way to the "Auction Proceeds" section-header row across cols P/Q/R,
    so it finds the real label even when it lives in a merged cell several
    rows above the column-header row.
  • Clean Aptos Narrow 12pt formatting, no bold, no underline, on both
    header and data rows.
  • ZIP surgery (no openpyxl save) keeps VBA, shared strings, drawings,
    and all other content intact.

Usage:
    python recovery_log_updater.py
"""

import os
import re
import zipfile
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook


# ─────────────────────────────────────────────────────────────────────────────
# Column headers (exactly as requested)
# ─────────────────────────────────────────────────────────────────────────────

HEADERS = [
    'Date',
    'Check',
    'Loan # ',
    'Payment Source',
    'Amount',
    'Principal - 11821',
    'Recovery - 11901',
    'Interest - 94040',
    'Fees - 43191',
    'Auction Fees - 83802',
    'Unam Dealer Reserve - 11831',
    'Unam Acquisition Fees - 11832',
    'Batch Label',
]


# ─────────────────────────────────────────────────────────────────────────────
# Filename helpers
# ─────────────────────────────────────────────────────────────────────────────

def build_output_path(input_path: str) -> str:
    p = Path(input_path)
    base = p.parent / f"{p.stem}_updated{p.suffix}"
    n = 2
    while base.exists():
        base = p.parent / f"{p.stem}_updated_{n}{p.suffix}"
        n += 1
    return str(base)


# ─────────────────────────────────────────────────────────────────────────────
# Excel data extraction  (openpyxl, data_only=True)
# ─────────────────────────────────────────────────────────────────────────────

DATE_SHEET_RE = re.compile(r'^(\d{2})\.(\d{2})\.(\d{2})')


def parse_sheet_date(name):
    m = DATE_SHEET_RE.match(name)
    if not m:
        return None
    mo, day, yr = int(m.group(1)), int(m.group(2)), int(m.group(3)) + 2000
    return datetime(yr, mo, day)


def rv(ws, row_idx, n=18):
    """Return a list of cell values (cols A–R) for the given 1-based row."""
    return [ws.cell(row=row_idx, column=c).value for c in range(1, n + 1)]


def is_section_hdr(v):
    b, d = v[1], v[3]
    if isinstance(b, str) and b.strip() in ('DD-COL', 'LBX-1', 'LBX-2'):
        return True
    if isinstance(d, str) and v[5] is None and v[6] is None and v[7] is None:
        return True
    return False


def is_col_hdr(v):
    return (isinstance(v[1], str) and v[1].strip() == 'Notes') or \
           (isinstance(v[2], str) and v[2].strip() == 'Notes')


def is_totals(v):
    return isinstance(v[3], str) and v[3].strip().startswith('Totals')


def is_data(v):
    if is_section_hdr(v) or is_col_hdr(v) or is_totals(v):
        return False
    return (isinstance(v[5], (int, float)) and v[5] != 0) or \
           (v[3] is not None and str(v[3]).strip() not in ('', '0', 'Loan # '))


def _valid_label(val) -> bool:
    """Return True for a real batch-label string (not None, not 'see image above')."""
    if val is None:
        return False
    s = str(val).strip()
    return bool(s) and 'see image above' not in s.lower()


def _scan_for_label(ws, start_row: int, stop_row: int) -> str | None:
    """
    Scan rows [start_row .. stop_row] (inclusive, going backwards from
    start_row) across columns P, Q, R for the first valid batch label.
    Stops at stop_row so we never cross section boundaries.
    """
    for r in range(start_row, stop_row - 1, -1):
        v = rv(ws, r)
        for col_idx in (15, 16, 17):          # cols P, Q, R (0-based)
            if col_idx < len(v) and _valid_label(v[col_idx]):
                return v[col_idx]
    return None


def extract_from_sheet(ws, date):
    batch_label = ws.cell(row=3, column=17).value   # col Q row 3 of date tab
    rows, seen = [], set()

    def add(d):
        key = (d['loan'], d['src'], d['amount'])
        if key not in seen:
            seen.add(key)
            rows.append(d)

    # ── Upper sections (rows 8–190) ──────────────────────────────────────────
    in_auct = False
    for r in range(8, 191):
        v = rv(ws, r)
        if is_section_hdr(v):
            in_auct = 'AUCT' in str(v[3] or '')
            continue
        if is_col_hdr(v) or is_totals(v) or not is_data(v):
            continue
        rec = v[7]
        if (isinstance(rec, (int, float)) and rec != 0) or in_auct:
            add(dict(date=date, check=v[2], loan=v[3], src=v[4],
                     amount=v[5], principal=v[6], recovery=v[7],
                     interest=v[8], fees=v[9], auction_fees=v[10],
                     unam_dr=v[11], unam_af=v[12], label=batch_label))

    # ── Bottom "Auction Proceeds" sections (rows 190+) ───────────────────────
    #
    # Batch-label search strategy:
    #   When we hit the column-header row, scan backward all the way back to
    #   the "Auction Proceeds" section-header row across cols P, Q, R.
    #   This handles merged cells where the real label lives several rows
    #   above, which openpyxl returns as None for the lower merged rows.
    #   Any cell containing "see image above" is skipped automatically by
    #   _valid_label().
    in_ap       = False
    ap_label    = None
    ap_start_r  = None     # row where current "Auction Proceeds" section began

    for r in range(190, ws.max_row + 1):
        v   = rv(ws, r)
        d4  = str(v[3] or '').strip()

        # Entering an Auction Proceeds section
        if d4 == 'Auction Proceeds' and v[5] is None:
            in_ap      = True
            ap_label   = None
            ap_start_r = r
            continue

        # Entering a different section — close the current one
        if d4 in ('Insurance Proceeds', 'Operating Deposit') and v[5] is None:
            in_ap = False
            continue

        if not in_ap:
            continue

        # Column-header row: find the batch label
        if is_col_hdr(v):
            # Scan from this row back to the section-header row (inclusive)
            ap_label = _scan_for_label(ws, r, ap_start_r)
            continue

        if is_totals(v) or not is_data(v):
            continue

        add(dict(date=date, check=None, loan=v[3], src=v[4],
                 amount=v[5], principal=v[6], recovery=v[7],
                 interest=v[8], fees=v[9], auction_fees=v[10],
                 unam_dr=None, unam_af=None,
                 label=ap_label or batch_label))

    return rows


# ─────────────────────────────────────────────────────────────────────────────
# styles.xml surgery: add clean size-12 styles
# ─────────────────────────────────────────────────────────────────────────────

def _update_styles_xml(xml_bytes: bytes):
    """
    Append four new cell-format (xf) entries:
      [+0] date cell    – mm-dd-yy, Aptos Narrow 12, thin border, left
      [+1] text cell    – general,  Aptos Narrow 12, thin border, left
      [+2] numeric cell – accounting, Aptos Narrow 12, thin border, left
      [+3] header cell  – general,  Aptos Narrow 12, thin border, center

    Returns (updated_bytes, date_s, text_s, num_s, hdr_s).
    """
    xml = xml_bytes.decode('utf-8')

    # ── Add Aptos Narrow 12pt font ───────────────────────────────────────────
    fc_m    = re.search(r'<fonts count="(\d+)"', xml)
    old_fc  = int(fc_m.group(1))
    new_fnt = ('<font>'
               '<sz val="12"/><color theme="1"/>'
               '<name val="Aptos Narrow"/>'
               '<family val="2"/><scheme val="minor"/>'
               '</font>')
    xml = xml.replace('</fonts>', new_fnt + '</fonts>', 1)
    xml = xml.replace(f'<fonts count="{old_fc}"',
                      f'<fonts count="{old_fc + 1}"', 1)
    fi = old_fc                    # 0-based index of the new font

    # ── Ensure mm-dd-yy numFmt ───────────────────────────────────────────────
    DATE_FMT = 166
    if 'formatCode="mm-dd-yy"' not in xml:
        nfc_m = re.search(r'<numFmts count="(\d+)"', xml)
        if nfc_m:
            old_nfc = int(nfc_m.group(1))
            xml = xml.replace(
                '</numFmts>',
                f'<numFmt numFmtId="{DATE_FMT}" formatCode="mm-dd-yy"/></numFmts>', 1)
            xml = xml.replace(f'<numFmts count="{old_nfc}"',
                              f'<numFmts count="{old_nfc + 1}"', 1)
        else:
            xml = xml.replace(
                '<cellStyleXfs',
                f'<numFmts count="1">'
                f'<numFmt numFmtId="{DATE_FMT}" formatCode="mm-dd-yy"/>'
                f'</numFmts><cellStyleXfs', 1)

    # ── Append four new xf entries ───────────────────────────────────────────
    xfc_m   = re.search(r'<cellXfs count="(\d+)"', xml)
    old_xfc = int(xfc_m.group(1))
    bi, fill = 1, 0              # borderId 1 = thin-all-sides; fillId 0 = none

    xf_date = (f'<xf numFmtId="{DATE_FMT}" fontId="{fi}" fillId="{fill}" '
               f'borderId="{bi}" xfId="0" '
               f'applyFont="1" applyBorder="1" applyNumberFormat="1"/>')

    xf_text = (f'<xf numFmtId="0" fontId="{fi}" fillId="{fill}" '
               f'borderId="{bi}" xfId="0" '
               f'applyFont="1" applyBorder="1" applyAlignment="1">'
               f'<alignment horizontal="left"/></xf>')

    xf_num  = (f'<xf numFmtId="43" fontId="{fi}" fillId="{fill}" '
               f'borderId="{bi}" xfId="0" '
               f'applyFont="1" applyBorder="1" applyNumberFormat="1" '
               f'applyAlignment="1">'
               f'<alignment horizontal="left"/></xf>')

    xf_hdr  = (f'<xf numFmtId="0" fontId="{fi}" fillId="{fill}" '
               f'borderId="{bi}" xfId="0" '
               f'applyFont="1" applyBorder="1" applyAlignment="1">'
               f'<alignment horizontal="center"/></xf>')

    xml = xml.replace('</cellXfs>',
                      xf_date + xf_text + xf_num + xf_hdr + '</cellXfs>', 1)
    xml = xml.replace(f'<cellXfs count="{old_xfc}"',
                      f'<cellXfs count="{old_xfc + 4}"', 1)

    date_s = old_xfc
    text_s = old_xfc + 1
    num_s  = old_xfc + 2
    hdr_s  = old_xfc + 3

    return xml.encode('utf-8'), date_s, text_s, num_s, hdr_s


# ─────────────────────────────────────────────────────────────────────────────
# Recovery sheet XML generation
# ─────────────────────────────────────────────────────────────────────────────

_EXCEL_EPOCH = datetime(1899, 12, 30)
_COLS        = list('ABCDEFGHIJKLM')


def _esc(text: str) -> str:
    return (str(text)
            .replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;'))


def _num_cell(col, row, val, s) -> str:
    if val is not None and val != 0:
        return f'<c r="{col}{row}" s="{s}"><v>{val}</v></c>'
    return f'<c r="{col}{row}" s="{s}"/>'


def _str_cell(col, row, val, s) -> str:
    if val is not None and str(val).strip():
        return (f'<c r="{col}{row}" s="{s}" t="inlineStr">'
                f'<is><t>{_esc(val)}</t></is></c>')
    return f'<c r="{col}{row}" s="{s}"/>'


def _header_row_xml(hdr_s: int) -> str:
    cells = [_str_cell(col, 3, hdr, hdr_s)
             for col, hdr in zip(_COLS, HEADERS)]
    return f'<row r="3" spans="1:13">{"".join(cells)}</row>'


def _data_row_xml(row_num: int, d: dict,
                  date_s: int, text_s: int, num_s: int) -> str:
    serial = (d['date'] - _EXCEL_EPOCH).days if d['date'] else None
    cells = [
        _num_cell('A', row_num, serial,            date_s),
        _num_cell('B', row_num, d['check'],         text_s),
        _num_cell('C', row_num, d['loan'],          text_s),
        _str_cell('D', row_num, d['src'],           text_s),
        _num_cell('E', row_num, d['amount'],        num_s),
        _num_cell('F', row_num, d['principal'],     num_s),
        _num_cell('G', row_num, d['recovery'],      num_s),
        _num_cell('H', row_num, d['interest'],      num_s),
        _num_cell('I', row_num, d['fees'],          num_s),
        _num_cell('J', row_num, d['auction_fees'],  num_s),
        _num_cell('K', row_num, d['unam_dr'],       num_s),
        _num_cell('L', row_num, d['unam_af'],       num_s),
        _str_cell('M', row_num, d['label'],         text_s),
    ]
    return f'<row r="{row_num}" spans="1:13">{"".join(cells)}</row>'


def _find_recovery_path(zf: zipfile.ZipFile) -> str:
    wb   = zf.read('xl/workbook.xml').decode('utf-8')
    m    = re.search(r'<sheet\s[^>]*name="Recovery"[^>]*r:id="([^"]+)"', wb)
    if not m:
        raise ValueError('Recovery sheet not found in workbook.xml')
    rid  = m.group(1)
    rels = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    m2   = re.search(
        rf'<Relationship\s[^>]*Id="{re.escape(rid)}"[^>]*Target="([^"]+)"',
        rels)
    if not m2:
        raise ValueError(f'Relationship {rid} not found')
    t = m2.group(1)
    if not t.startswith('/') and not t.startswith('xl/'):
        t = 'xl/' + t
    return t.lstrip('/')


def _update_recovery_xml(xml_bytes: bytes, all_rows: list,
                          date_s, text_s, num_s, hdr_s) -> bytes:
    """
    Keep row 2 (summation) intact.
    Replace row 3 with fresh column headers.
    Replace rows 4+ with new data rows.
    """
    xml = xml_bytes.decode('utf-8')

    # Build all new content (header row 3 + data rows 4+)
    new_content = _header_row_xml(hdr_s) + '\n'
    new_content += '\n'.join(
        _data_row_xml(4 + i, d, date_s, text_s, num_s)
        for i, d in enumerate(all_rows)
    )

    # Cut point: first occurrence of row 3 or higher
    first_old = re.search(r'<row r="[3-9]\d*"', xml)
    cut_start  = first_old.start() if first_old else xml.index('</sheetData>')
    sd_end     = xml.index('</sheetData>')

    new_xml = xml[:cut_start] + new_content + '\n' + xml[sd_end:]

    # Update <dimension> to reflect actual extent
    last_row = 3 + len(all_rows)
    new_xml = re.sub(r'<dimension ref="[^"]*"',
                     f'<dimension ref="A2:M{last_row}"',
                     new_xml)
    return new_xml.encode('utf-8')


# ─────────────────────────────────────────────────────────────────────────────
# Main ZIP surgery
# ─────────────────────────────────────────────────────────────────────────────

def update_xlsm(input_path: str, output_path: str, all_rows: list):
    """
    Copy the original .xlsm, replacing only xl/styles.xml and the Recovery
    worksheet XML.  All other ZIP entries are copied byte-for-byte.
    """
    STYLES = 'xl/styles.xml'

    with zipfile.ZipFile(input_path, 'r') as zin:
        rec_path = _find_recovery_path(zin)

        styles_b                              = zin.read(STYLES)
        new_styles, date_s, text_s, num_s, hdr_s = _update_styles_xml(styles_b)

        with zipfile.ZipFile(output_path, 'w',
                             compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == STYLES:
                    data = new_styles
                elif item.filename == rec_path:
                    data = _update_recovery_xml(
                        data, all_rows, date_s, text_s, num_s, hdr_s)
                zout.writestr(item, data)


# ─────────────────────────────────────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title('Recovery Log Updater')
        self.resizable(False, False)
        self.attributes('-topmost', True)

        self._inp = tk.StringVar()
        self._out = tk.StringVar()
        pad = dict(padx=10, pady=6)

        tk.Label(self, text='Source file:').grid(row=0, column=0, sticky='e', **pad)
        tk.Entry(self, textvariable=self._inp, width=54,
                 state='readonly').grid(row=0, column=1, **pad)
        tk.Button(self, text='Browse…', width=9,
                  command=self._browse).grid(row=0, column=2, **pad)

        tk.Label(self, text='Output file:').grid(row=1, column=0, sticky='e', **pad)
        tk.Entry(self, textvariable=self._out, width=54,
                 state='readonly').grid(row=1, column=1, **pad)

        self._log = tk.Text(self, width=66, height=12, state='disabled',
                            font=('Courier', 9), wrap='word')
        self._log.grid(row=2, column=0, columnspan=3, padx=10, pady=(0, 4))

        self._bar = ttk.Progressbar(self, length=500, mode='indeterminate')
        self._bar.grid(row=3, column=0, columnspan=3, padx=10, pady=(0, 4))

        bf = tk.Frame(self)
        bf.grid(row=4, column=0, columnspan=3, pady=(0, 10))
        self._run = tk.Button(bf, text='Run', width=12, state='disabled',
                              command=self._go)
        self._run.pack(side='left', padx=6)
        tk.Button(bf, text='Close', width=12,
                  command=self.destroy).pack(side='left', padx=6)

        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f'+{(sw-w)//2}+{(sh-h)//2}')
        self.after(200, lambda: self.attributes('-topmost', False))

    def _log_msg(self, msg):
        self._log.configure(state='normal')
        self._log.insert('end', msg + '\n')
        self._log.see('end')
        self._log.configure(state='disabled')
        self.update()

    def _browse(self):
        path = filedialog.askopenfilename(
            parent=self,
            title='Select the bank deposit workbook',
            filetypes=[
                ('Excel macro-enabled workbook', '*.xlsm'),
                ('Excel workbook',               '*.xlsx'),
                ('All files',                    '*.*'),
            ],
        )
        if path:
            self._inp.set(path)
            self._out.set(build_output_path(path))
            self._run.configure(state='normal')

    def _go(self):
        inp = self._inp.get()
        out = self._out.get()
        if not inp:
            messagebox.showwarning('No file', 'Please browse for a file first.')
            return

        self._run.configure(state='disabled')
        self._bar.start(12)

        try:
            self._log_msg(f'Source:  {inp}')
            self._log_msg(f'Output:  {out}\n')
            self._log_msg('Reading workbook data…')

            wb_read     = load_workbook(inp, data_only=True)
            date_sheets = [s for s in wb_read.sheetnames
                           if DATE_SHEET_RE.match(s)]
            self._log_msg(f'Found {len(date_sheets)} date tab(s):\n  '
                          + ', '.join(date_sheets) + '\n')

            all_rows = []
            for sn in date_sheets:
                rows = extract_from_sheet(wb_read[sn], parse_sheet_date(sn))
                self._log_msg(f'  {sn}: {len(rows)} row(s) extracted')
                all_rows.extend(rows)

            self._log_msg(
                f'\nWriting {len(all_rows)} data row(s) + header row…')
            update_xlsm(inp, out, all_rows)

            self._bar.stop()
            self._bar.configure(mode='determinate', value=100)
            self._log_msg(f'\nDone!  Saved to:\n  {out}')
            messagebox.showinfo('Complete',
                                f'Recovery log updated.\n\nSaved to:\n{out}')

        except Exception as exc:
            self._bar.stop()
            self._log_msg(f'\nERROR: {exc}')
            if os.path.exists(out):
                try:
                    os.remove(out)
                except Exception:
                    pass
            messagebox.showerror('Error', str(exc))

        finally:
            self._run.configure(state='normal')


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app = App()
    app.mainloop()