"""
Invoice Reconciliation Script
==============================
Matches clinic invoices to business invoices by shared reference number.
Flags cash payment invoices (containing "Administrative and Support Fees").
Moves matched invoices to /reconciled, unmatched to /unmatched.
Outputs a reconciliation report (Excel).

Folder structure expected:
  input/
    business/   ← e.g. #35891.pdf
    clinic/     ← e.g. (35891) 3232.pdf
  output/
    reconciled/
    unmatched/
    reconciliation_report.xlsx
"""

import os
import re
import shutil
import glob
from pathlib import Path
from decimal import Decimal, InvalidOperation

import pdfplumber
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── Configuration ────────────────────────────────────────────────────────────

CASH_PAYMENT_TRIGGER = "Administrative and Support Fees"
CURRENCY = "SGD"

INPUT_BUSINESS_DIR = "input/business"
INPUT_CLINIC_DIR   = "input/clinic"
OUTPUT_RECONCILED  = "output/reconciled"
OUTPUT_UNMATCHED   = "output/unmatched"
OUTPUT_REPORT      = "output/reconciliation_report.xlsx"


# ─── Helpers ──────────────────────────────────────────────────────────────────

def extract_ref_from_business(filename: str) -> str | None:
    """Extract reference number from business invoice filename.
    e.g. '#35891.pdf' → '35891'
    """
    name = Path(filename).stem  # strip .pdf
    match = re.search(r'#?(\d+)', name)
    return match.group(1) if match else None


def extract_ref_from_clinic(filename: str) -> str | None:
    """Extract reference number from clinic invoice filename.
    e.g. '(35891) 3232.pdf' → '35891'
    """
    name = Path(filename).stem
    match = re.search(r'\((\d+)\)', name)
    return match.group(1) if match else None


def extract_text_from_pdf(filepath: str) -> str:
    """Extract all text from a PDF using pdfplumber."""
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"  ⚠ Could not read {filepath}: {e}")
    return text


def extract_sgd_total(text: str) -> Decimal | None:
    """
    Attempt to extract the invoice total amount from PDF text.
    Looks for patterns like:
      - SGD 1,234.56
      - Total: 1,234.56
      - Total SGD 1,234.56
      - Amount Due: 1,234.56
      - Grand Total 1,234.56
    Returns the LAST matched amount (usually the final total on the invoice).
    """
    patterns = [
        r'SGD\s*\$?\s*([\d,]+\.\d{2})',
        r'\$\s*([\d,]+\.\d{2})',
        r'(?:Grand\s+)?Total(?:\s+Due)?(?:\s+SGD)?[\s:]*\$?\s*([\d,]+\.\d{2})',
        r'Amount\s+Due[\s:]*\$?\s*([\d,]+\.\d{2})',
        r'Total\s+Amount[\s:]*\$?\s*([\d,]+\.\d{2})',
        r'Invoice\s+Total[\s:]*\$?\s*([\d,]+\.\d{2})',
    ]

    found = []
    for pattern in patterns:
        for m in re.finditer(pattern, text, re.IGNORECASE):
            raw = m.group(1).replace(',', '')
            try:
                found.append(Decimal(raw))
            except InvalidOperation:
                pass

    # Return the last amount found (most likely the final total)
    return found[-1] if found else None


def is_cash_payment_invoice(text: str) -> bool:
    return CASH_PAYMENT_TRIGGER.lower() in text.lower()


def ensure_dirs():
    for d in [INPUT_BUSINESS_DIR, INPUT_CLINIC_DIR, OUTPUT_RECONCILED, OUTPUT_UNMATCHED]:
        os.makedirs(d, exist_ok=True)


# ─── Main Reconciliation Logic ─────────────────────────────────────────────────

def run_reconciliation():
    ensure_dirs()

    print("\n" + "="*60)
    print("  INVOICE RECONCILIATION")
    print("="*60)

    # ── Step 1: Load & classify business invoices ──────────────────
    print("\n[1/4] Scanning business invoices...")

    business_invoices = {}   # ref → {filename, filepath, total, is_cash}
    cash_invoices     = []   # list of info dicts

    for filepath in glob.glob(os.path.join(INPUT_BUSINESS_DIR, "*.pdf")):
        filename = os.path.basename(filepath)
        ref = extract_ref_from_business(filename)

        if not ref:
            print(f"  ⚠ Could not extract reference from: {filename} — skipping")
            continue

        text  = extract_text_from_pdf(filepath)
        total = extract_sgd_total(text)
        is_cash = is_cash_payment_invoice(text)

        info = {
            "ref":      ref,
            "filename": filename,
            "filepath": filepath,
            "total":    total,
            "is_cash":  is_cash,
            "text":     text,
        }

        if is_cash:
            print(f"  💵 CASH PAYMENT  → {filename}  (ref: {ref})")
            cash_invoices.append(info)
        else:
            print(f"  📄 Business      → {filename}  (ref: {ref}, total: {CURRENCY} {total})")
            business_invoices[ref] = info

    # ── Step 2: Load clinic invoices ───────────────────────────────
    print("\n[2/4] Scanning clinic invoices...")

    clinic_groups = {}   # ref → list of {filename, filepath, total}

    for filepath in glob.glob(os.path.join(INPUT_CLINIC_DIR, "*.pdf")):
        filename = os.path.basename(filepath)
        ref = extract_ref_from_clinic(filename)

        if not ref:
            print(f"  ⚠ Could not extract reference from: {filename} — skipping")
            continue

        text  = extract_text_from_pdf(filepath)
        total = extract_sgd_total(text)

        info = {
            "ref":      ref,
            "filename": filename,
            "filepath": filepath,
            "total":    total,
        }

        clinic_groups.setdefault(ref, []).append(info)
        print(f"  🏥 Clinic        → {filename}  (ref: {ref}, total: {CURRENCY} {total})")

    # ── Step 3: Match & tally ─────────────────────────────────────
    print("\n[3/4] Matching & tallying...")

    report_rows = []    # for Excel report
    all_refs = set(business_invoices.keys()) | set(clinic_groups.keys())

    for ref in sorted(all_refs):
        biz = business_invoices.get(ref)
        clinics = clinic_groups.get(ref, [])

        biz_total    = biz["total"] if biz else None
        clinic_total = sum(c["total"] for c in clinics if c["total"] is not None) if clinics else Decimal("0")
        clinic_files = [c["filename"] for c in clinics]

        # Determine match status
        if not biz:
            status = "⚠ NO BUSINESS INVOICE"
        elif not clinics:
            status = "⚠ NO CLINIC INVOICES"
        elif biz_total is None:
            status = "⚠ BUSINESS TOTAL UNREADABLE"
        elif clinic_total == biz_total:
            status = "✅ MATCHED"
        else:
            status = "❌ MISMATCH"

        print(f"  Ref {ref}: Biz={biz_total} | Clinic sum={clinic_total} | {status}")

        # Move clinic files
        for clinic in clinics:
            if status == "✅ MATCHED":
                dest_dir = OUTPUT_RECONCILED
            else:
                dest_dir = OUTPUT_UNMATCHED

            dest_path = os.path.join(dest_dir, clinic["filename"])
            shutil.copy2(clinic["filepath"], dest_path)

        report_rows.append({
            "ref":            ref,
            "business_file":  biz["filename"] if biz else "—",
            "business_total": float(biz_total) if biz_total else None,
            "clinic_files":   "\n".join(clinic_files) if clinic_files else "—",
            "clinic_sum":     float(clinic_total),
            "difference":     float(clinic_total - biz_total) if (biz_total and clinic_total) else None,
            "status":         status,
        })

    # Add cash payment invoices to report
    for c in cash_invoices:
        report_rows.append({
            "ref":            c["ref"],
            "business_file":  c["filename"],
            "business_total": float(c["total"]) if c["total"] else None,
            "clinic_files":   "—",
            "clinic_sum":     0.0,
            "difference":     None,
            "status":         "💵 CASH PAYMENT",
        })

    # ── Step 4: Write Excel report ────────────────────────────────
    print("\n[4/4] Writing reconciliation report...")
    write_excel_report(report_rows)

    # ── Summary ───────────────────────────────────────────────────
    matched   = sum(1 for r in report_rows if "MATCHED" in r["status"] and "CASH" not in r["status"])
    mismatched= sum(1 for r in report_rows if "MISMATCH" in r["status"])
    no_match  = sum(1 for r in report_rows if "NO " in r["status"] or "UNREADABLE" in r["status"])
    cash      = len(cash_invoices)

    print("\n" + "="*60)
    print(f"  ✅ Matched:        {matched}")
    print(f"  ❌ Mismatched:     {mismatched}")
    print(f"  ⚠  Unmatched/err: {no_match}")
    print(f"  💵 Cash Payment:  {cash}")
    print(f"\n  Report saved to:  {OUTPUT_REPORT}")
    print(f"  Reconciled PDFs:  {OUTPUT_RECONCILED}/")
    print(f"  Unmatched PDFs:   {OUTPUT_UNMATCHED}/")
    print("="*60 + "\n")


# ─── Excel Report ─────────────────────────────────────────────────────────────

def write_excel_report(rows: list[dict]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reconciliation"

    # Colour palette
    GREEN  = PatternFill("solid", fgColor="C6EFCE")
    RED    = PatternFill("solid", fgColor="FFC7CE")
    ORANGE = PatternFill("solid", fgColor="FFEB9C")
    BLUE   = PatternFill("solid", fgColor="BDD7EE")
    GREY   = PatternFill("solid", fgColor="D9D9D9")

    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = [
        "Reference #",
        "Business Invoice",
        f"Business Total ({CURRENCY})",
        "Clinic Invoice(s)",
        f"Clinic Sum ({CURRENCY})",
        f"Difference ({CURRENCY})",
        "Status",
    ]

    # Header row
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font      = Font(bold=True, color="FFFFFF")
        cell.fill      = PatternFill("solid", fgColor="2F4F7F")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = border
    ws.row_dimensions[1].height = 30

    # Data rows
    for i, row in enumerate(rows, start=2):
        status = row["status"]

        if   "MATCHED"   in status and "CASH" not in status: fill = GREEN
        elif "MISMATCH"  in status:                           fill = RED
        elif "CASH"      in status:                           fill = BLUE
        elif "NO "       in status or "UNREADABLE" in status: fill = ORANGE
        else:                                                  fill = GREY

        values = [
            row["ref"],
            row["business_file"],
            row["business_total"],
            row["clinic_files"],
            row["clinic_sum"],
            row["difference"],
            status,
        ]

        for col, val in enumerate(values, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.fill      = fill
            cell.border    = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

            # Format currency columns
            if col in (3, 5, 6) and val is not None:
                cell.number_format = f'"{CURRENCY}" #,##0.00'

        ws.row_dimensions[i].height = max(
            15 * row["clinic_files"].count("\n") + 20, 20
        )

    # Column widths
    widths = [14, 28, 22, 40, 20, 20, 26]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # Freeze header
    ws.freeze_panes = "A2"

    # Add a summary sheet
    ws2 = wb.create_sheet("Summary")
    matched   = sum(1 for r in rows if "MATCHED"    in r["status"] and "CASH" not in r["status"])
    mismatched= sum(1 for r in rows if "MISMATCH"   in r["status"])
    no_match  = sum(1 for r in rows if "NO "        in r["status"] or "UNREADABLE" in r["status"])
    cash      = sum(1 for r in rows if "CASH"       in r["status"])

    summary_data = [
        ("Invoice Reconciliation Summary", None),
        (None, None),
        ("Category",       "Count"),
        ("✅ Matched",      matched),
        ("❌ Mismatched",   mismatched),
        ("⚠ Unmatched",    no_match),
        ("💵 Cash Payment", cash),
        (None, None),
        ("Total processed", len(rows)),
    ]

    for r_idx, (label, value) in enumerate(summary_data, 1):
        ws2.cell(row=r_idx, column=1, value=label)
        if value is not None:
            ws2.cell(row=r_idx, column=2, value=value)

    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 12
    ws2["A1"].font = Font(bold=True, size=13)

    wb.save(OUTPUT_REPORT)


# ─── Entry Point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    run_reconciliation()
