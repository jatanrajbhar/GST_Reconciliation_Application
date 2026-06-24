# GST Reconciliation Tool

> Copyright (c) 2024-2025 Jatan Rajbhar. All rights reserved. — [LICENSE](LICENSE.md) | [NOTICE](NOTICE.md)


A Windows desktop application for Indian businesses to reconcile GSTR-2B supplier data against their ITC (Input Tax Credit) purchase register — built with Python and CustomTkinter.

**Developed by [GSC in time](mailto:info@gscintime.com)**

---

## What It Does

Every month, GST-registered businesses must compare invoices uploaded by their suppliers on the GST portal (GSTR-2B) against their own purchase records (ITC register). Mismatches mean ITC claims may be disallowed by the government. This tool automates that reconciliation, flags differences, and produces a colour-coded Excel report.

---

## Features

| Feature | Description |
| --- | --- |
| **GSTR-2B vs ITC Reconciliation** | Matches B2B, B2BA, CDNR, CDNRA, IMPG, IMPGSEZ against your ITC register |
| **Smart Invoice Matching** | Exact match → amount-based fuzzy match → string-similarity fallback |
| **Suffix / Prefix Guard** | Prevents false cross-vendor matches while allowing valid cross-fiscal-year matches |
| **Status Classification** | Matched / Higher in 2A / Lower in 2A / Not found in 2A per invoice |
| **2B Results Sheet** | Every 2B row gets a status: Matched, Unmatched, or Not Found in ITC |
| **Debug Matching** | Review borderline pairs one-by-one; accept or skip with remarks |
| **YTD Database** | SQLite-backed year-to-date store; browse by period, mark matched/unmatched |
| **Match with Past 2A DB** | Reconcile current ITC against previously unmatched 2B rows from database |
| **Excel Export** | Colour-coded ITC Results + 2B Results in a single workbook |
| **Download Template** | Pre-formatted Excel template with all required sheets |
| **Upload Excel (All Sheets)** | Load all sheets from a single workbook in one click |
| **GST Portal Status** | Optional: automated portal check via Selenium |

---

## Installation

### Prerequisites

- Python 3.9 or higher
- pip

### Install dependencies

```bash
pip install -r requirements_desktop.txt
```

### Run

```bash
python gst_reconciliation_app.py
```

> **Windows executable**: A pre-built `.exe` installer is available for clients — contact [info@gscintime.com](mailto:info@gscintime.com).

---

## Quick Start

### Step 1 — Download the template

Click **Download Template** (top-right of dashboard) to get the Excel workbook with all required sheets pre-formatted.

### Step 2 — Fill in your data

| Sheet | Content |
| --- | --- |
| `2B-B2B` | Invoice-wise B2B supplies from GSTR-2B |
| `2B-B2BA` | Amendments to B2B |
| `2B-CDNR` | Credit / Debit Notes Received |
| `2B-CDNRA` | Amendments to CDNR |
| `2B-IMPG` | Import of Goods |
| `2B-IMPGSEZ` | Import of Goods from SEZ |
| `Register-ITC` | Your company's ITC purchase register *(mandatory)* |

### Step 3 — Upload

- **All at once**: click **Upload Excel (All Sheets)** and select your filled workbook.
- **Sheet by sheet**: use the Browse button on each row (CSV or xlsx).

### Step 4 — Reconcile

Click **Process and Reconcile**. Progress is shown in the log panel.

### Step 5 — Review and download results

- **View Results** — interactive table with colour-coded status rows.
- **Download ITC Results** — Excel with ITC Results + 2B Results sheets.
- **Download 2A Results** — standalone 2B results sheet.

### Step 6 — Debug borderline matches *(optional)*

Click **Debug Matching** to review pairs where GSTIN and amounts agree but invoice numbers only partially match. Accept or skip each pair, then click **Save to YTD**.

### Step 7 — YTD Database *(optional)*

- **Save to YTD Database** — stores ITC and 2B rows with status in a local SQLite database.
- **YTD Database** (sidebar) — browse saved data by database → year → month.
- **Match with Past 2A DB** — match current ITC register against unmatched 2B rows from previous months.

---

## Reconciliation Status Values

| Status | Meaning |
| --- | --- |
| **Matched** | Invoice found in 2B; tax amounts agree |
| **Higher in 2A** | 2B tax amount is higher than ITC register |
| **Lower in 2A** | 2B tax amount is lower than ITC register |
| **Not found in 2A** | No matching 2B invoice found for this ITC row |
| **Unmatched** | Key matched but amounts differ beyond tolerance |
| **Matched but invoice number is not accurate** | Manually matched via Debug Matching |

---

## Output Columns — ITC Results Sheet

| Column | Description |
| --- | --- |
| Status | Reconciliation result (see above) |
| CGST / SGST / IGST as per 2A | Tax amounts from the matched 2B row |
| Type | Source table: B2B, CDNR, IMPG, B2BA … or *ytd database* |
| Booking Month as per GSTR-2A | Tax period from the 2B file |
| Booking Month as per ITC | Tax period from the ITC register |
| 2A Invoice No | Invoice number of the matched 2B row |
| 2A GSTIN | GSTIN of the matched 2B supplier |
| Remarks | Auto-generated notes (many-to-one, YTD match, debug action) |

---

## File Column Requirements

Column names are matched **case-insensitively** and flexibly — the tool finds columns by keyword, not exact name.

**B2B / B2BA / CDNR / CDNRA**
- Invoice No / Note No — GSTN / GSTIN — CGST — SGST — IGST — Taxable Value

**IMPG / IMPGSEZ**
- BOE No / Invoice No — GSTN / GSTIN *(optional)* — IGST — Taxable Value

**ITC Register**
- Vendor's GSTN / Vendor GSTIN
- Vendor Inv. No / External Doc No
- CGST — SGST — IGST — Taxable Value

---

## Tech Stack

| Component | Library |
| --- | --- |
| UI Framework | [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) |
| Data Processing | Pandas, NumPy |
| Excel I/O | OpenPyXL |
| Database | SQLite (`sqlite3` stdlib) |
| Portal Automation | Selenium *(optional)* |
| Build | PyInstaller + Inno Setup |

---

## Building the Executable

```bash
# Install build dependencies
pip install -r requirements_desktop.txt

# Build the .exe
python build.py
```

The spec files `GST_Reconciliation_Tool.spec` / `gst_reconciliation.spec` are used by PyInstaller. The Inno Setup script `installer.iss` packages the output into a Windows installer.

See [BUILD_README.md](BUILD_README.md) for detailed build instructions.

---

## Support & Licensing

This is a **commercial product**. License keys are required for production use.

**GSC in time**
- Email: [info@gscintime.com](mailto:info@gscintime.com)
- Phone: +91-22-4612 5600

