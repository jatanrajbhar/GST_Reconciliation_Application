# GST Reconciliation Tool

A portable desktop application for automated GST (Goods and Services Tax) reconciliation and data processing. This tool helps businesses efficiently match, validate, and reconcile GST transactions across multiple document types.

## Features

- **Multi-Format Support**: Process ITC (Input Tax Credit), B2B, B2BA, CDNR, CDNRA, IMPG, and IMPGSEZ transactions
- **Duplicate Merging**: Automatically merge duplicate vendor invoices and consolidate tax amounts
- **Intelligent Matching**: Fuzzy matching logic for invoices and GSTIN numbers to handle minor discrepancies
- **Amendment Processing**: Match and update original transactions with amendment documents
- **Data Validation**: Comprehensive error checking and data quality validation
- **Excel & CSV Support**: Import data from both Excel (multiple sheets) and CSV files
- **Template Generation**: Download pre-formatted templates for data entry
- **Detailed Logging**: Real-time processing logs with step-by-step reconciliation details
- **Precision Rounding**: ROUND_HALF_UP rounding for accurate rupee calculations
- **User-Friendly GUI**: Modern, intuitive graphical interface built with CustomTkinter

## System Requirements

- **Python**: 3.8 or higher
- **OS**: Windows, macOS, or Linux
- **Memory**: Minimum 2GB RAM
- **Storage**: 100MB for application and dependencies

## Installation

### Prerequisites

- Python 3.8+ installed on your system

### Setup Instructions

1. **Clone or Extract the Repository**
   ```bash
   cd "GST Reconciliation tool"
   ```

2. **Create a Virtual Environment** (Recommended)
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

   Or manually install:
   ```bash
   pip install pandas numpy openpyxl pillow customtkinter
   ```

## Usage

### Running the Application

```bash
python gst_reconciliation_app.py
```

The application will open with an intuitive graphical interface for managing your GST reconciliation workflow.

### Basic Workflow

1. **Load Data Files**
   - Click "Browse" to select your input Excel or CSV files
   - The tool automatically detects and processes columns containing GSTIN, Invoice No, and tax amounts (CGST, SGST, IGST)

2. **Select Transaction Types**
   - Choose which GST document types to process:
     - **ITC**: Input Tax Credit invoices
     - **B2B**: Business-to-Business supplies
     - **B2BA**: B2B amendments
     - **CDNR**: Credit/Debit Notes - Registered
     - **CDNRA**: Credit/Debit Notes amendments
     - **IMPG**: Import of Goods
     - **IMPGSEZ**: Import of Goods - SEZ

3. **Configure Options**
   - Set filtering criteria if needed
   - Enable/disable specific reconciliation steps

4. **Execute Reconciliation**
   - Click the "Process" or "Reconcile" button
   - Monitor progress in the real-time logging window
   - Review results and generated reports

5. **Export Results**
   - Save reconciled data as Excel files
   - Export detailed reports for audit trails

### File Format Requirements

#### Excel Files

- Each sheet should represent a different transaction type or document class
- Required columns vary by document type but typically include:
  - Invoice Number / BOE No
  - GSTIN (Goods and Services Tax Identification Number)
  - CGST Amount (Central GST)
  - SGST Amount (State GST)
  - IGST Amount (Integrated GST)

#### CSV Files

- Standard comma-separated values format
- First row should contain column headers
- Same column naming conventions as Excel files

### Downloading Templates

Use the "Download Template" option to get pre-formatted Excel files with:
- Correct column structure
- Sample data
- Multiple sheets for different document types
- Ready-to-use layout for your data

## Supported GST Transaction Types

| Type | Full Name | Description |
|------|-----------|-------------|
| **ITC** | Input Tax Credit | Tax credit from inbound supplies |
| **B2B** | Business-to-Business | Supplies from one registered business to another |
| **B2BA** | B2B Amendment | Corrections to B2B invoices |
| **CDNR** | Credit/Debit Notes - Registered | Adjustments for registered suppliers |
| **CDNRA** | CDNR Amendment | Corrections to credit/debit notes |
| **IMPG** | Import of Goods | Foreign purchases with integrated GST |
| **IMPGSEZ** | Import of Goods - SEZ | Special Economic Zone imports |

## Advanced Features

### Fuzzy Matching

- Handles minor formatting variations in invoice numbers
- Matches GST numbers with leading/trailing whitespace differences
- Strips non-alphanumeric characters for better matching accuracy

### Duplicate Handling

- Automatically consolidates multiple entries for the same invoice
- Sums tax amounts (CGST, SGST, IGST) for duplicates
- Preserves first occurrence of other data fields

### Data Normalization

- Converts scientific notation in invoice numbers to standard format
- Standardizes GSTIN format (15-character alphanumeric)
- Handles various numeric formats and currency symbols

### Amendment Reconciliation

- Matches original invoices with amendment/correction documents
- Updates original records with amendment data
- Tracks matched and unmatched amendments

## Troubleshooting

### Common Issues

**Issue**: Application won't start
- **Solution**: Ensure Python 3.8+ is installed and all dependencies are installed via pip

**Issue**: Files not detected
- **Solution**: Verify file format (Excel .xlsx or CSV .csv), ensure headers are in the first row

**Issue**: Column not recognized
- **Solution**: Check that column names contain the expected keywords (GSTIN, CGST, SGST, IGST, Invoice, BOE)

**Issue**: Rounding discrepancies
- **Solution**: The tool uses ROUND_HALF_UP method for accuracy; minor rounding variations may occur with certain decimal values

## Data Privacy & Security

- No data is sent to external servers
- All processing occurs locally on your machine
- Files are not automatically saved; you control all output
- Compatible with standard Windows, macOS, and Linux file systems

## Performance Notes

- Handles datasets up to 100,000+ rows efficiently
- Processing speed depends on file size and system capabilities
- For very large files (>500,000 rows), consider splitting into batches

## Technical Stack

- **GUI Framework**: CustomTkinter (modern tkinter wrapper)
- **Data Processing**: Pandas, NumPy
- **File Handling**: openpyxl (Excel), CSV module
- **Image Processing**: Pillow
- **Precision Math**: Python Decimal module

## License & Attribution

This tool is provided as-is for GST reconciliation purposes. Ensure compliance with Indian GST regulations and your organization's data policies.

## Support & Feedback

For issues, feature requests, or improvements, please:
1. Review the error logs displayed in the application
2. Check that your input files follow the recommended format
3. Verify all required columns are present and properly named

## Version History

- **v1.0**: Initial release with core reconciliation features
  - Support for multiple GST document types
  - ITC duplicate merging
  - Amendment matching and updating
  - Excel and CSV support
  - User-friendly GUI interface

---

## Important Note

There might be some faulty values in CGST, SGST, IGST as per 2A column but the matching is 99% accurate.

**Last Updated**: February 2026

For more information on GST regulations and requirements, visit the [GST Council Official Website](https://gst.gov.in)
