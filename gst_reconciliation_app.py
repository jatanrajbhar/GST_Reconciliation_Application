"""
GST Reconciliation Tool - Desktop Application
A portable desktop application for GST reconciliation processing.
"""

import os
import sys
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk
import pandas as pd
import numpy as np
import math
from decimal import Decimal, ROUND_HALF_UP


def round_rupee(x):
    """Round a numeric value to nearest rupee using ROUND_HALF_UP behavior.
    Returns 0 for NaN or non-finite values."""
    try:
        if pd.isna(x) or not np.isfinite(x):
            return 0
        # Use Decimal for consistent HALF_UP rounding on .5
        d = Decimal(str(float(x)))
        return int(d.quantize(Decimal('1'), rounding=ROUND_HALF_UP))
    except Exception:
        return 0
from io import BytesIO
from PIL import Image
import re
from difflib import SequenceMatcher

# Set appearance mode and default color theme
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")  # Will override with custom magenta/pink colors

# Custom magenta/pink theme colors
THEME_PRIMARY = "#E91E63"  # Magenta Pink
THEME_HOVER = "#C2185B"    # Darker Pink
THEME_LIGHT = "#FCE4EC"    # Light Pink background
THEME_DARK = "#880E4F"     # Dark Pink/Maroon

# Increase Pandas Styler limit for large dataframes
pd.set_option("styler.render.max_elements", 5000000)


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def safe_numeric_conversion(value):
    """Safely convert value to numeric, handling various formats"""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        # Remove common formatting characters
        value = value.replace(',', '').replace('\u20b9', '').strip()
        if value in ['', '-', 'NA', 'N/A', 'null', 'NULL', 'None']:
            return 0.0
        try:
            return float(value)
        except:
            return 0.0
    return 0.0


# Helper: normalize invoice strings for consistent matching
def normalize_invoice(s):
    """Return a normalized invoice string: uppercase, alphanumeric only, strip leading zeros.
    Also expands scientific notation (e.g. '2.7183E+13') to full integer string for consistent
    matching between CSV and Excel loaded data."""
    if pd.isna(s):
        return ''
    s = str(s).strip()
    # Expand scientific notation to full integer string (e.g. '2.7183E+13' -> '27183000000000')
    if re.match(r'^-?[\d.]+[eE][+\-]?\d+$', s):
        try:
            s = str(int(float(s)))
        except (ValueError, OverflowError):
            pass
    s = s.upper()
    # Remove date-like parts, whitespace, and non-alphanumeric characters
    s = re.sub(r'[^A-Z0-9]', '', s)
    s = s.lstrip('0')
    return s


# Helper: normalize GSTIN to canonical 15-char alphanum uppercase without spaces
def normalize_gstin(s):
    if pd.isna(s):
        return ''
    s = str(s).upper().replace(' ', '')
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s


# Helper: fuzzy similarity between two strings
def similarity(a, b):
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def find_tax_amount_columns(df):
    """Find the correct CGST, SGST, IGST amount columns, preferring 'Amount' columns over '%' columns.
    Returns (cgst_col, sgst_col, igst_col) tuple with column names or None if not found."""
    cgst_col = None
    sgst_col = None
    igst_col = None

    # First pass: look specifically for 'Amount' columns
    for col in df.columns:
        col_lower = col.lower().strip()
        if 'cgst' in col_lower and 'amount' in col_lower and not cgst_col:
            cgst_col = col
        elif 'sgst' in col_lower and 'amount' in col_lower and not sgst_col:
            sgst_col = col
        elif 'igst' in col_lower and 'amount' in col_lower and not igst_col:
            igst_col = col

    # Second pass: look for exact 'cgst', 'sgst', 'igst' columns (without % or amount suffix)
    for col in df.columns:
        col_lower = col.lower().strip()
        if col_lower == 'cgst' and not cgst_col:
            cgst_col = col
        elif col_lower == 'sgst' and not sgst_col:
            sgst_col = col
        elif col_lower == 'igst' and not igst_col:
            igst_col = col

    # Third pass: fall back to any column containing cgst/sgst/igst but NOT containing '%'
    for col in df.columns:
        col_lower = col.lower().strip()
        if 'cgst' in col_lower and '%' not in col_lower and not cgst_col:
            cgst_col = col
        elif 'sgst' in col_lower and '%' not in col_lower and not sgst_col:
            sgst_col = col
        elif 'igst' in col_lower and '%' not in col_lower and not igst_col:
            igst_col = col

    return cgst_col, sgst_col, igst_col


def find_booking_month_column(df):
    """Find the booking month / period column in a DataFrame.
    Searches for: 'Booking Month', 'Month', 'Period' (in priority order).
    Returns column name or None."""
    if df is None or df.empty:
        return None
    for col in df.columns:
        cl = col.lower().strip()
        if cl == 'booking month':
            return col
    for col in df.columns:
        cl = col.lower().strip()
        if cl == 'month':
            return col
    for col in df.columns:
        cl = col.lower().strip()
        if cl == 'period':
            return col
    return None


def create_template_excel():
    """Create a template Excel file for download with sample data"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sample_gstin = '27AABCU9603R1ZM'

        itc_template = pd.DataFrame({
            'Invoice No': ['INV001', 'INV002'],
            'GSTIN': [sample_gstin, sample_gstin],
            'CGST': [1000.00, 2000.00],
            'SGST': [1000.00, 2000.00],
            'IGST': [0.00, 0.00]
        })
        itc_template.to_excel(writer, index=False, sheet_name='ITC')

        b2b_template = pd.DataFrame({
            'Invoice No': ['INV001', 'INV002'],
            'GSTIN': [sample_gstin, sample_gstin],
            'CGST': [1000.00, 2000.00],
            'SGST': [1000.00, 2000.00],
            'IGST': [0.00, 0.00]
        })
        b2b_template.to_excel(writer, index=False, sheet_name='B2B')

        b2ba_template = pd.DataFrame({
            'Invoice No': ['INV001'],
            'GSTIN': [sample_gstin],
            'CGST': [1100.00],
            'SGST': [1100.00],
            'IGST': [0.00]
        })
        b2ba_template.to_excel(writer, index=False, sheet_name='B2BA')

        cdnr_template = pd.DataFrame({
            'Invoice No': ['CDN001', 'CDN002'],
            'GSTIN': [sample_gstin, sample_gstin],
            'CGST': [500.00, 300.00],
            'SGST': [500.00, 300.00],
            'IGST': [0.00, 0.00]
        })
        cdnr_template.to_excel(writer, index=False, sheet_name='CDNR')

        cdnra_template = pd.DataFrame({
            'Invoice No': ['CDN001'],
            'GSTIN': [sample_gstin],
            'CGST': [550.00],
            'SGST': [550.00],
            'IGST': [0.00]
        })
        cdnra_template.to_excel(writer, index=False, sheet_name='CDNRA')

        impg_template = pd.DataFrame({
            'Invoice No': ['BOE001', 'BOE002'],
            'GSTIN': [sample_gstin, sample_gstin],
            'CGST': [0.00, 0.00],
            'SGST': [0.00, 0.00],
            'IGST': [5000.00, 3000.00]
        })
        impg_template.to_excel(writer, index=False, sheet_name='IMPG')

        impgsez_template = pd.DataFrame({
            'Invoice No': ['SEZ001', 'SEZ002'],
            'GSTIN': [sample_gstin, sample_gstin],
            'CGST': [0.00, 0.00],
            'SGST': [0.00, 0.00],
            'IGST': [2000.00, 1500.00]
        })
        impgsez_template.to_excel(writer, index=False, sheet_name='IMPGSEZ')

    output.seek(0)
    return output.getvalue()


def normalize_itc_columns(itc_df):
    """Normalize ITC column names - handle Invoice No and GSTIN aliases"""
    if itc_df is None or itc_df.empty:
        return itc_df

    itc_df = itc_df.copy()

    # Convert all object/datetime columns to string to avoid type issues
    for col in itc_df.columns:
        if itc_df[col].dtype == 'object' or 'datetime' in str(itc_df[col].dtype):
            itc_df[col] = itc_df[col].astype(str).replace('nan', '').replace('NaT', '')

    has_invoice_no = False
    has_vendor_inv = False
    invoice_no_col = None
    vendor_inv_col = None

    for col in itc_df.columns:
        col_lower = col.lower().strip()
        if col_lower == 'invoice no':
            has_invoice_no = True
            invoice_no_col = col
        elif 'vendor inv' in col_lower or 'external doc' in col_lower:
            has_vendor_inv = True
            vendor_inv_col = col

    if has_invoice_no:
        if has_vendor_inv:
            itc_df = itc_df.rename(columns={invoice_no_col: 'Vendor Inv. No/ External Doc No'})
        else:
            itc_df = itc_df.rename(columns={invoice_no_col: 'Vendor Inv. No/ External Doc No'})

    has_vendor_gstn = False
    gstin_col_to_rename = None

    for col in itc_df.columns:
        col_lower = col.lower().strip()
        if 'vendor' in col_lower and 'gstn' in col_lower:
            has_vendor_gstn = True
        elif col_lower == 'gstin' or ('supplier' in col_lower and 'gstin' in col_lower):
            gstin_col_to_rename = col

    if gstin_col_to_rename and not has_vendor_gstn:
        itc_df = itc_df.rename(columns={gstin_col_to_rename: "Vendor's GSTN"})

    return itc_df


def normalize_cdnr_columns(df):
    """Normalize CDNR/CDNRA column names - handle Invoice No as BOE No alias"""
    if df is None or df.empty:
        return df

    df = df.copy()

    # Convert all object/datetime columns to string to avoid type issues
    for col in df.columns:
        if df[col].dtype == 'object' or 'datetime' in str(df[col].dtype):
            df[col] = df[col].astype(str).replace('nan', '').replace('NaT', '')

    has_invoice_no = False
    has_boe_no = False
    invoice_no_col = None

    for col in df.columns:
        col_lower = col.lower().strip()
        if col_lower == 'invoice no':
            has_invoice_no = True
            invoice_no_col = col
        elif 'boe no' in col_lower:
            has_boe_no = True

    if has_invoice_no:
        if has_boe_no:
            df = df.rename(columns={invoice_no_col: 'BOE No'})
        else:
            df = df.rename(columns={invoice_no_col: 'BOE No'})

    return df


def merge_duplicate_vendor_invoices(itc_df, log_callback=None):
    """Step 1: Merge duplicate Vendor Inv. No in ITC"""
    if itc_df is None or itc_df.empty:
        if log_callback:
            log_callback("Warning: ITC table is empty or not loaded")
        return itc_df

    # Create a copy to avoid modifying original
    itc_df = itc_df.copy()

    vendor_inv_col = None
    vendor_gstn_col = None

    for col in itc_df.columns:
        col_lower = col.lower().strip()
        if 'vendor inv' in col_lower or 'external doc' in col_lower:
            vendor_inv_col = col
        elif 'vendor' in col_lower and 'gstn' in col_lower:
            vendor_gstn_col = col

    # Use helper to find correct amount columns (not percentage columns)
    cgst_col, sgst_col, igst_col = find_tax_amount_columns(itc_df)

    if not vendor_inv_col:
        if log_callback:
            log_callback("Error: Could not find 'Vendor Inv. No/ External Doc No' column in ITC")
        return itc_df

    original_count = len(itc_df)

    # Convert all non-numeric columns to string to avoid datetime/mixed type issues
    for col in itc_df.columns:
        if col not in [cgst_col, sgst_col, igst_col]:
            # Convert to string, handling NaN values
            itc_df[col] = itc_df[col].astype(str).replace('nan', '').replace('NaT', '')

    numeric_cols = []
    if cgst_col:
        itc_df[cgst_col] = itc_df[cgst_col].apply(safe_numeric_conversion)
        numeric_cols.append(cgst_col)
    if sgst_col:
        itc_df[sgst_col] = itc_df[sgst_col].apply(safe_numeric_conversion)
        numeric_cols.append(sgst_col)
    if igst_col:
        itc_df[igst_col] = itc_df[igst_col].apply(safe_numeric_conversion)
        numeric_cols.append(igst_col)

    agg_dict = {}
    for col in itc_df.columns:
        if col in numeric_cols:
            agg_dict[col] = 'sum'
        else:
            agg_dict[col] = 'first'

    merged_itc = itc_df.groupby(vendor_inv_col, as_index=False).agg(agg_dict)
    merged_count = len(merged_itc)

    if log_callback:
        log_callback(f"Step 1: Merged {original_count - merged_count} duplicate entries. Original: {original_count} -> After merge: {merged_count}")

    return merged_itc


def match_and_update(main_df, amendment_df, main_invoice_col, amend_invoice_col, table_name, log_callback=None):
    """Generic function to match and update tables (Steps 2 & 3).
    Returns (updated_main_df, remaining_amendment_df) where remaining has matched rows removed."""
    if main_df is None or main_df.empty:
        if log_callback:
            log_callback(f"{table_name} table is empty")
        return main_df, amendment_df

    if amendment_df is None or amendment_df.empty:
        if log_callback:
            log_callback(f"Amendment table is empty, skipping update for {table_name}")
        return main_df, amendment_df

    # Use global helper to find correct amount columns (not percentage columns)
    main_cgst, main_sgst, main_igst = find_tax_amount_columns(main_df)
    amend_cgst, amend_sgst, amend_igst = find_tax_amount_columns(amendment_df)

    if not main_invoice_col or not amend_invoice_col:
        if log_callback:
            log_callback(f"Could not find invoice columns for {table_name}")
        return main_df, amendment_df

    for df, cols in [(main_df, [main_cgst, main_sgst, main_igst]),
                     (amendment_df, [amend_cgst, amend_sgst, amend_igst])]:
        for col in cols:
            if col and col in df.columns:
                df[col] = df[col].apply(safe_numeric_conversion)

    matched_count = 0
    matched_indices = []
    main_df = main_df.copy()

    for idx, amend_row in amendment_df.iterrows():
        amend_invoice = amend_row[amend_invoice_col]
        matching_mask = main_df[main_invoice_col] == amend_invoice

        if matching_mask.any():
            matched_count += 1
            matched_indices.append(idx)
            if main_cgst and amend_cgst:
                main_df.loc[matching_mask, main_cgst] = amend_row[amend_cgst]
            if main_sgst and amend_sgst:
                main_df.loc[matching_mask, main_sgst] = amend_row[amend_sgst]
            if main_igst and amend_igst:
                main_df.loc[matching_mask, main_igst] = amend_row[amend_igst]

    # Remove matched rows from amendment table
    remaining_amendment = amendment_df.drop(matched_indices).reset_index(drop=True)

    # Round tax columns in the main table to nearest rupee (if decimals exist) and guard against NaN/inf
    for col in [main_cgst, main_sgst, main_igst]:
        if col and col in main_df.columns:
            # Ensure numeric and round to integer using HALF_UP, replacing NaN/inf with 0
            main_df[col] = main_df[col].apply(safe_numeric_conversion).apply(lambda x: round_rupee(x))

    # Reset index and return
    main_df = main_df.reset_index(drop=True)
    remaining_amendment = remaining_amendment.reset_index(drop=True)

    if log_callback:
        if matched_count > 0:
            log_callback(f"Matched and updated {matched_count} records in {table_name}, removed from amendment table")
        else:
            log_callback(f"No matches found for {table_name}, proceeding to next step")

    return main_df, remaining_amendment


def create_merged_table(tables_dict, log_callback=None):
    """Step 4: Create MERGED table"""
    merged_data = []
    table_order = ['B2B', 'B2BA', 'CDNR', 'CDNRA', 'IMPG', 'IMPGSEZ']

    for table_name in table_order:
        df = tables_dict.get(table_name)

        if df is None or df.empty:
            if log_callback:
                log_callback(f"{table_name}: No data")
            continue

        doc_col = None
        gstn_col = None
        tax_col = None
        date_col = None

        for col in df.columns:
            col_lower = col.lower().strip()
            if 'invoice no' in col_lower or 'note no' in col_lower or 'boe no' in col_lower:
                if not doc_col:
                    doc_col = col
            if 'gstn' in col_lower or 'gstin' in col_lower:
                if not gstn_col:
                    gstn_col = col
            if 'taxable' in col_lower and 'value' in col_lower and not tax_col:
                tax_col = col
            if col_lower == 'invoice date' and not date_col:
                date_col = col

        # Use helper to find correct amount columns (not percentage columns)
        cgst_col, sgst_col, igst_col = find_tax_amount_columns(df)
        bm_col = find_booking_month_column(df)

        if not doc_col:
            if log_callback:
                log_callback(f"{table_name}: Could not find document number column")
            continue

        for _, row in df.iterrows():
            record = {
                'Document_number': row[doc_col] if doc_col else None,
                'GSTN': row[gstn_col] if gstn_col else None,
                'CGST': safe_numeric_conversion(row[cgst_col]) if cgst_col else None,
                'SGST': safe_numeric_conversion(row[sgst_col]) if sgst_col else None,
                'IGST': safe_numeric_conversion(row[igst_col]) if igst_col else None,
                'TAX': safe_numeric_conversion(row[tax_col]) if tax_col else 0.0,
                'Invoice_Date': str(row[date_col]) if date_col else '',
                'Booking_Month': str(row[bm_col]).strip() if bm_col else '',
                'TYPE': table_name
            }
            merged_data.append(record)

        if log_callback:
            log_callback(f"Added {len(df)} records from {table_name}")

    merged_df = pd.DataFrame(merged_data)
    if log_callback:
        log_callback(f"Step 4: MERGED table created with {len(merged_df)} total records")

    return merged_df


def create_itc_register(itc_df, log_callback=None):
    """Step 5: Create as_per_itc_register table"""
    if itc_df is None or itc_df.empty:
        if log_callback:
            log_callback("Error: ITC table is empty")
        return pd.DataFrame()

    vendor_gstn_col = None
    vendor_inv_col = None
    tax_col = None

    for col in itc_df.columns:
        col_lower = col.lower().strip()
        if 'vendor' in col_lower and 'gstn' in col_lower:
            vendor_gstn_col = col
        elif 'vendor inv' in col_lower or 'external doc' in col_lower:
            vendor_inv_col = col
        elif 'taxable' in col_lower and 'value' in col_lower and not tax_col:
            tax_col = col

    # Use helper to find correct amount columns (not percentage columns)
    cgst_col, sgst_col, igst_col = find_tax_amount_columns(itc_df)

    if not vendor_gstn_col or not vendor_inv_col:
        if log_callback:
            log_callback("Error: Could not find required columns in ITC")
        return pd.DataFrame()

    itc_register = pd.DataFrame({
        'GSTINinvoice': itc_df[vendor_gstn_col].astype(str) + itc_df[vendor_inv_col].astype(str),
        'CGST': itc_df[cgst_col].apply(safe_numeric_conversion) if cgst_col else 0.0,
        'SGST': itc_df[sgst_col].apply(safe_numeric_conversion) if sgst_col else 0.0,
        'IGST': itc_df[igst_col].apply(safe_numeric_conversion) if igst_col else 0.0,
        'TAX': itc_df[tax_col].apply(safe_numeric_conversion) if tax_col else 0.0
    })

    if log_callback:
        log_callback(f"Step 5: Created as_per_itc_register with {len(itc_register)} records")

    return itc_register


def create_gstr_2a(merged_df, log_callback=None):
    """Step 6: Create as_per_gtsr_2a table"""
    if merged_df is None or merged_df.empty:
        if log_callback:
            log_callback("Error: MERGED table is empty")
        return pd.DataFrame()

    gstr_2a = pd.DataFrame({
        'GSTINinvoice': merged_df['GSTN'].astype(str) + merged_df['Document_number'].astype(str),
        'CGST': merged_df['CGST'].apply(safe_numeric_conversion),
        'SGST': merged_df['SGST'].apply(safe_numeric_conversion),
        'IGST': merged_df['IGST'].apply(safe_numeric_conversion),
        'TAX': merged_df['TAX'].apply(safe_numeric_conversion) if 'TAX' in merged_df.columns else 0.0
    })

    if log_callback:
        log_callback(f"Step 6: Created as_per_gtsr_2a with {len(gstr_2a)} records")

    return gstr_2a


def create_itc_register_2(itc_df, log_callback=None):
    """Step 9: Create as_per_itc_register(2) table (per-line) with normalized GSTIN and numeric TAX"""
    if itc_df is None or itc_df.empty:
        if log_callback:
            log_callback("Error: ITC table is empty for Step 9")
        return pd.DataFrame()

    vendor_gstn_col = None
    tax_col = None

    for col in itc_df.columns:
        col_lower = col.lower().strip()
        if 'vendor' in col_lower and 'gstn' in col_lower:
            vendor_gstn_col = col
        elif 'taxable' in col_lower and 'value' in col_lower and not tax_col:
            tax_col = col

    # Use helper to find correct amount columns (not percentage columns)
    cgst_col, sgst_col, igst_col = find_tax_amount_columns(itc_df)

    if not vendor_gstn_col or not tax_col:
        if log_callback:
            log_callback("Error: Could not find GSTIN or Taxable Value columns in ITC for Step 9")
        return pd.DataFrame()

    # Build per-line register with normalized GSTIN and numeric TAX (rounded to nearest rupee)
    reg = pd.DataFrame()
    reg['GSTIN'] = itc_df[vendor_gstn_col].apply(normalize_gstin)
    reg['TAX'] = itc_df[tax_col].apply(safe_numeric_conversion).apply(lambda x: round_rupee(x))
    reg['CGST'] = (itc_df[cgst_col].apply(safe_numeric_conversion).apply(lambda x: round_rupee(x))) if cgst_col else 0
    reg['SGST'] = (itc_df[sgst_col].apply(safe_numeric_conversion).apply(lambda x: round_rupee(x))) if sgst_col else 0
    reg['IGST'] = (itc_df[igst_col].apply(safe_numeric_conversion).apply(lambda x: round_rupee(x))) if igst_col else 0

    if log_callback:
        log_callback(f"Step 9: Created as_per_itc_register(2) (per-line) with {len(reg)} records")

    return reg


def create_gstr_2a_2(merged_df, log_callback=None):
    """Step 10: Create as_per_gtsr_2a(2) table with normalized GSTIN and numeric TAX"""
    if merged_df is None or merged_df.empty:
        if log_callback:
            log_callback("Error: MERGED table is empty for Step 10")
        return pd.DataFrame()

    df = merged_df.copy()
    df['GSTIN'] = df['GSTN'].astype(str).apply(normalize_gstin)
    df['TAX'] = (df['TAX'].apply(safe_numeric_conversion).apply(lambda x: round_rupee(x))) if 'TAX' in df.columns else 0
    df['CGST'] = df['CGST'].apply(safe_numeric_conversion).apply(lambda x: round_rupee(x))
    df['SGST'] = df['SGST'].apply(safe_numeric_conversion).apply(lambda x: round_rupee(x))
    df['IGST'] = df['IGST'].apply(safe_numeric_conversion).apply(lambda x: round_rupee(x))

    gstr_2a_2 = df[['GSTIN', 'TAX', 'CGST', 'SGST', 'IGST']].copy()

    if log_callback:
        log_callback(f"Step 10: Created as_per_gtsr_2a(2) with {len(gstr_2a_2)} records")

    return gstr_2a_2


def compare_tables_by_tax(itc_register_2, gstr_2a_2, log_callback=None):
    """Step 11: Compare tables by GSTIN and TAX (numeric), and map results back to per-line ITC rows."""
    if itc_register_2.empty or gstr_2a_2.empty:
        if log_callback:
            log_callback("Error: One or both comparison tables are empty for tax matching")
        return pd.DataFrame()

    # Both inputs are per-line DataFrames with 'GSTIN' and 'TAX' numeric columns
    # Aggregate to (GSTIN, TAX) level for stable comparison
    itc_agg = itc_register_2.groupby(['GSTIN', 'TAX'], as_index=False).agg({'CGST':'sum','SGST':'sum','IGST':'sum'})
    gstr_agg = gstr_2a_2.groupby(['GSTIN', 'TAX'], as_index=False).agg({'CGST':'sum','SGST':'sum','IGST':'sum'})

    # Round aggregated tax columns to nearest rupee (integers) to remove decimal points
    for c in ['CGST', 'SGST', 'IGST']:
        if c in itc_agg.columns:
            vals = itc_agg[c].fillna(0).astype(float).to_numpy()
            itc_agg[c] = np.floor(vals + 0.5).astype(int)
        if c in gstr_agg.columns:
            vals = gstr_agg[c].fillna(0).astype(float).to_numpy()
            gstr_agg[c] = np.floor(vals + 0.5).astype(int)

    # Merge on GSTIN and TAX
    comparison = pd.merge(itc_agg, gstr_agg, on=['GSTIN','TAX'], how='outer', suffixes=('_ITC','_GSTR2A'))

    for col in ['CGST_ITC', 'SGST_ITC', 'IGST_ITC', 'CGST_GSTR2A', 'SGST_GSTR2A', 'IGST_GSTR2A']:
        if col in comparison.columns:
            comparison[col] = comparison[col].fillna(0)
        else:
            comparison[col] = 0

    comparison['CGST_Difference'] = comparison['CGST_ITC'] - comparison['CGST_GSTR2A']
    comparison['SGST_Difference'] = comparison['SGST_ITC'] - comparison['SGST_GSTR2A']
    comparison['IGST_Difference'] = comparison['IGST_ITC'] - comparison['IGST_GSTR2A']

    tolerance = 10.0

    def get_status(row):
        cgst_match = abs(row['CGST_Difference']) <= tolerance
        sgst_match = abs(row['SGST_Difference']) <= tolerance
        igst_match = abs(row['IGST_Difference']) <= tolerance

        return 'MATCH' if (cgst_match and sgst_match and igst_match) else 'MISMATCH'

    comparison['Status'] = comparison.apply(get_status, axis=1)

    # Map statuses back to per-line ITC rows so user sees full original lines
    # Itc_register_2 contains per-line GSTIN and TAX
    itc_lines = itc_register_2.copy()
    itc_lines['Status'] = itc_lines.merge(
        comparison[['GSTIN','TAX','Status']],
        on=['GSTIN','TAX'], how='left')['Status']
    itc_lines['Status'] = itc_lines['Status'].fillna('MISMATCH')

    matched = len(comparison[comparison['Status'] == 'MATCH'])
    mismatched = len(comparison[comparison['Status'] == 'MISMATCH'])
    total = len(comparison)

    if log_callback:
        log_callback(f"Step 11: Tax Value Matching - Total invoice groups: {total}, Matched: {matched}, Mismatched: {mismatched}")
        log_callback(f"Step 11: Tax value mapped back to {len(itc_lines)} ITC line items")

    return itc_lines


def compare_tables(itc_register, gstr_2a, log_callback=None):
    """Step 7: Compare and show detailed matching results using normalized GSTIN+Invoice keys."""
    if itc_register is None or itc_register.empty or gstr_2a is None or gstr_2a.empty:
        if log_callback:
            log_callback("Error: One or both comparison tables are empty")
        return pd.DataFrame()

    # Work on copies
    itc = itc_register.copy()
    gstr = gstr_2a.copy()

    # Helper to split concatenated GSTINinvoice and normalize parts
    def _split_and_normalize(s):
        if pd.isna(s) or str(s).strip() == '':
            return ('', '')
        raw = str(s).upper()
        m = re.search(r'([0-9A-Z]{15})', raw)
        if m:
            gstin_raw = m.group(1)
            inv_raw = raw.replace(gstin_raw, '', 1)
        else:
            gstin_raw = raw[:15] if len(raw) >= 15 else raw
            inv_raw = raw[15:] if len(raw) > 15 else ''
        return (normalize_gstin(gstin_raw), normalize_invoice(inv_raw))

    def derive_norm_parts(df_input):
        """Extract and normalize GSTIN/Invoice from concatenated or separate columns."""
        if 'GSTINinvoice' in df_input.columns:
            parts = df_input['GSTINinvoice'].astype(str).apply(_split_and_normalize)
            df_input['_norm_gstin'] = parts.apply(lambda x: x[0])
            df_input['_norm_inv'] = parts.apply(lambda x: x[1])
        else:
            df_input['_norm_gstin'] = ''
            df_input['_norm_inv'] = ''
        return df_input

    itc = derive_norm_parts(itc)
    gstr = derive_norm_parts(gstr)

    # Build normalized composite key
    itc['_key'] = itc['_norm_gstin'].fillna('').astype(str) + '|' + itc['_norm_inv'].fillna('').astype(str)
    gstr['_key'] = gstr['_norm_gstin'].fillna('').astype(str) + '|' + gstr['_norm_inv'].fillna('').astype(str)

    # Aggregate taxes by normalized key
    itc_agg = itc.groupby('_key', as_index=False).agg({'CGST': 'sum', 'SGST': 'sum', 'IGST': 'sum'})
    gstr_agg = gstr.groupby('_key', as_index=False).agg({'CGST': 'sum', 'SGST': 'sum', 'IGST': 'sum'})

    # Ensure numeric and round to nearest rupee
    for df_agg in [itc_agg, gstr_agg]:
        for col in ['CGST', 'SGST', 'IGST']:
            df_agg[col] = df_agg[col].apply(safe_numeric_conversion).apply(lambda x: round_rupee(x))

    # Rename columns before merge
    itc_agg = itc_agg.rename(columns={'CGST': 'CGST_ITC', 'SGST': 'SGST_ITC', 'IGST': 'IGST_ITC'})
    gstr_agg = gstr_agg.rename(columns={'CGST': 'CGST_GSTR2A', 'SGST': 'SGST_GSTR2A', 'IGST': 'IGST_GSTR2A'})

    # Merge on normalized key
    comparison = pd.merge(itc_agg, gstr_agg, on='_key', how='outer')

    # Fill NaN with 0
    for col in ['CGST_ITC', 'SGST_ITC', 'IGST_ITC', 'CGST_GSTR2A', 'SGST_GSTR2A', 'IGST_GSTR2A']:
        if col not in comparison.columns:
            comparison[col] = 0
        comparison[col] = comparison[col].fillna(0)

    # Compute differences
    comparison['CGST_Difference'] = comparison['CGST_ITC'] - comparison['CGST_GSTR2A']
    comparison['SGST_Difference'] = comparison['SGST_ITC'] - comparison['SGST_GSTR2A']
    comparison['IGST_Difference'] = comparison['IGST_ITC'] - comparison['IGST_GSTR2A']

    # Build compact GSTINinvoice (no separator) for downstream compatibility
    def _compact(k):
        if pd.isna(k) or k == '':
            return ''
        gstin, inv = k.split('|', 1) if '|' in k else (k[:15], k[15:])
        return (gstin or '') + (inv or '')
    
    comparison['GSTINinvoice'] = comparison['_key'].apply(_compact)

    # Compare with tolerance
    tolerance = 10.0
    def _get_status(r):
        return 'MATCH' if (abs(r['CGST_Difference']) <= tolerance and 
                          abs(r['SGST_Difference']) <= tolerance and 
                          abs(r['IGST_Difference']) <= tolerance) else 'MISMATCH'
    
    comparison['Status'] = comparison.apply(_get_status, axis=1)

    # Clean up and reorder columns
    comparison = comparison.drop(columns=['_key'])
    
    cols_order = ['GSTINinvoice', 'CGST_ITC', 'SGST_ITC', 'IGST_ITC', 'CGST_GSTR2A', 'SGST_GSTR2A', 'IGST_GSTR2A', 'CGST_Difference', 'SGST_Difference', 'IGST_Difference', 'Status']
    existing_cols = [c for c in cols_order if c in comparison.columns]
    remaining_cols = [c for c in comparison.columns if c not in existing_cols]
    comparison = comparison[existing_cols + remaining_cols]

    total_records = len(comparison)
    matched = len(comparison[comparison['Status'] == 'MATCH'])
    mismatched = len(comparison[comparison['Status'] == 'MISMATCH'])

    if log_callback:
        log_callback(f"Step 7: Reconciliation complete - Total: {total_records}, Matched: {matched}, Mismatched: {mismatched}")

    return comparison


def create_itc_result(itc_df, itc_register, gstr_2a, comparison_df=None, merged_df=None, log_callback=None):
    """Create ITC table with improved matching logic and use `comparison_df` when available for consistency.
    If `comparison_df` is provided (aggregated matches from compare_tables), it will be used to assign invoice-level
    statuses (preferred) and then mapped back to original ITC rows so counts in ITC Results mirror reconciliation.

    Status values: Matched, Unmatched, Higher in 2A, Lower in 2A"""
    if itc_df is None or itc_df.empty:
        if log_callback:
            log_callback("Error: ITC table is empty")
        return pd.DataFrame()

    result = itc_df.copy()

    # Find GSTN and Invoice No columns in ITC
    vendor_gstn_col = None
    vendor_inv_col = None

    for col in itc_df.columns:
        col_lower = col.lower().strip()
        if 'vendor' in col_lower and 'gstn' in col_lower:
            vendor_gstn_col = col
        elif 'vendor inv' in col_lower or 'external doc' in col_lower:
            vendor_inv_col = col

    # Use helper to find correct amount columns (not percentage columns)
    itc_cgst_col, itc_sgst_col, itc_igst_col = find_tax_amount_columns(itc_df)

    if not vendor_gstn_col or not vendor_inv_col:
        if log_callback:
            log_callback("Error: Could not find GSTN or Invoice No columns in ITC")
        result['Status'] = 'Unmatched'
        return result

    # Build normalized keys on original ITC rows
    result['_norm_gstin'] = result[vendor_gstn_col].apply(normalize_gstin)
    result['_norm_inv'] = result[vendor_inv_col].apply(normalize_invoice)
    result['GSTINinvoice_norm'] = result['_norm_gstin'] + '|' + result['_norm_inv']

    # If comparison_df is available, use it to build a status lookup (prefer this for consistency)
    status_lookup = {}
    tolerance = 10.0


    # If any keys are still missing, fall back to earlier normalized matching algorithm
    # Aggregate original ITC to invoice level for comparison
    # Convert tax columns to numeric BEFORE aggregation (they are strings from dtype=str loading)
    agg_src = result[['GSTINinvoice_norm']].copy()
    if itc_cgst_col:
        agg_src['CGST'] = result[itc_cgst_col].apply(safe_numeric_conversion)
    if itc_sgst_col:
        agg_src['SGST'] = result[itc_sgst_col].apply(safe_numeric_conversion)
    if itc_igst_col:
        agg_src['IGST'] = result[itc_igst_col].apply(safe_numeric_conversion)

    agg_cols = {c: 'sum' for c in ['CGST', 'SGST', 'IGST'] if c in agg_src.columns}
    itc_agg = agg_src.groupby('GSTINinvoice_norm').agg(agg_cols).reset_index()

    # Build gstr lookup from provided gstr_2a (normalized)
    gstr_2a_local = gstr_2a.copy()
    if 'GSTINinvoice' in gstr_2a_local.columns:
        gstr_2a_local['_gstin_raw'] = gstr_2a_local['GSTINinvoice'].astype(str).apply(lambda x: x[:15] if len(x) >= 15 else x)
        gstr_2a_local['_inv_raw'] = gstr_2a_local['GSTINinvoice'].astype(str).apply(lambda x: x[15:] if len(x) > 15 else '')
    else:
        gstr_2a_local['_gstin_raw'] = ''
        gstr_2a_local['_inv_raw'] = ''
    if 'GSTN' in gstr_2a_local.columns and 'Document_number' in gstr_2a_local.columns:
        gstr_2a_local['_gstin_raw'] = gstr_2a_local['GSTN'].astype(str)
        gstr_2a_local['_inv_raw'] = gstr_2a_local['Document_number'].astype(str)
    gstr_2a_local['_norm_gstin'] = gstr_2a_local['_gstin_raw'].apply(normalize_gstin)
    gstr_2a_local['_norm_inv'] = gstr_2a_local['_inv_raw'].apply(normalize_invoice)
    gstr_2a_local['GSTINinvoice_norm'] = gstr_2a_local['_norm_gstin'] + '|' + gstr_2a_local['_norm_inv']
    gstr_agg = gstr_2a_local.groupby('GSTINinvoice_norm', as_index=False).agg({'CGST': 'sum', 'SGST': 'sum', 'IGST': 'sum'})
    for c in ['CGST', 'SGST', 'IGST']:
        if c in gstr_agg.columns:
            gstr_agg[c] = gstr_agg[c].apply(safe_numeric_conversion)

    gstr_lookup = {row['GSTINinvoice_norm']:{'CGST':row['CGST'],'SGST':row['SGST'],'IGST':row['IGST']} for _, row in gstr_agg.iterrows()}
    itc_lookup = {row['GSTINinvoice_norm']:{'CGST':row.get('CGST',0),'SGST':row.get('SGST',0),'IGST':row.get('IGST',0)} for _, row in itc_agg.iterrows()}

    # Build TYPE and Booking Month lookups from merged_df for 2A columns
    type_lookup = {}
    booking_month_2a_lookup = {}  # key → booking month string from 2A files
    if merged_df is not None and not merged_df.empty:
        for _, row in merged_df.iterrows():
            gstin = normalize_gstin(str(row.get('GSTN', '')))
            inv = normalize_invoice(str(row.get('Document_number', '')))
            key = gstin + '|' + inv
            if key not in type_lookup:
                type_lookup[key] = str(row.get('TYPE', ''))
            if key not in booking_month_2a_lookup:
                bm = str(row.get('Booking_Month', '')).strip()
                if bm:
                    booking_month_2a_lookup[key] = bm
    vals_2a_lookup = {}  # key → {'CGST', 'SGST', 'IGST', 'TYPE'}

    if comparison_df is not None and not comparison_df.empty:
        # Only use MATCH statuses from comparison_df
        # For MISMATCH, let the fuzzy matching below handle it (can find cross-fiscal-year matches by GSTIN + similar amounts)
        for _, row in comparison_df.iterrows():
            raw = str(row.get('GSTINinvoice', ''))
            # try to extract 15-char GSTIN if present
            m = re.search(r'[0-9A-Z]{15}', raw.upper())
            if m:
                gstin_raw = m.group(0)
                inv_raw = raw.upper().replace(gstin_raw, '', 1)
            else:
                gstin_raw = raw[:15]
                inv_raw = raw[15:]

            key = normalize_gstin(gstin_raw) + '|' + normalize_invoice(inv_raw)
            status_raw = str(row.get('Status', '')).upper()
            if status_raw == 'MATCH' or status_raw == 'MATCHED':
                status_lookup[key] = 'Matched'
                # Store 2A values for matched keys
                if key in gstr_lookup:
                    g = gstr_lookup[key]
                    vals_2a_lookup[key] = {'CGST': g['CGST'], 'SGST': g['SGST'], 'IGST': g['IGST'], 'TYPE': type_lookup.get(key, ''), 'BM': booking_month_2a_lookup.get(key, '')}
            # For MISMATCH cases, don't set status here - let fuzzy matching below handle them
            # This allows finding matches with different fiscal year suffixes (e.g., 21-22 vs 22-23)

    # per-GSTIN candidate sets
    gstr_by_gstin = {}
    for _, row in gstr_agg.iterrows():
        gstin = row['GSTINinvoice_norm'].split('|',1)[0]
        if gstin not in gstr_by_gstin:
            gstr_by_gstin[gstin] = []
        gstr_by_gstin[gstin].append({'key': row['GSTINinvoice_norm'], 'CGST': row['CGST'], 'SGST': row['SGST'], 'IGST': row['IGST']})

    # Employ fallback matching only for keys not set yet
    for key in set(list(itc_lookup.keys()) + list(gstr_lookup.keys())):
        if key in status_lookup:
            continue
        matched_key = None
        if key in gstr_lookup:
            itc_vals = itc_lookup.get(key, {'CGST':0,'SGST':0,'IGST':0})
            gstr_vals = gstr_lookup.get(key, {'CGST':0,'SGST':0,'IGST':0})
            matched_key = key
        else:
            gstin, inv = key.split('|',1)
            itc_vals = itc_lookup.get(key, {'CGST':0,'SGST':0,'IGST':0})
            candidates = gstr_by_gstin.get(gstin, [])
            gstr_vals = None
            best = None
            best_diff = None
            for cand in candidates:
                cg_diff = abs(itc_vals['CGST'] - cand['CGST'])
                sg_diff = abs(itc_vals['SGST'] - cand['SGST'])
                ig_diff = abs(itc_vals['IGST'] - cand['IGST'])
                total_diff = cg_diff + sg_diff + ig_diff
                if best is None or total_diff < best_diff:
                    best = cand
                    best_diff = total_diff
            if best is not None and best_diff <= (tolerance * 3):
                gstr_vals = {'CGST':best['CGST'],'SGST':best['SGST'],'IGST':best['IGST']}
                matched_key = best['key']
            if gstr_vals is None and candidates:
                best_sim = 0.0
                best_cand = None
                for cand in candidates:
                    cand_inv = cand['key'].split('|',1)[1]
                    sim = similarity(inv, cand_inv)
                    if sim > best_sim:
                        best_sim = sim
                        best_cand = cand
                if best_sim >= 0.85:
                    gstr_vals = {'CGST':best_cand['CGST'],'SGST':best_cand['SGST'],'IGST':best_cand['IGST']}
                    matched_key = best_cand['key']
            if gstr_vals is None:
                gstr_vals = {'CGST':0,'SGST':0,'IGST':0}
        cgst_diff = itc_vals['CGST'] - gstr_vals['CGST']
        sgst_diff = itc_vals['SGST'] - gstr_vals['SGST']
        igst_diff = itc_vals['IGST'] - gstr_vals['IGST']
        exact_match_in_2a = key in gstr_lookup
        if abs(cgst_diff) <= tolerance and abs(sgst_diff) <= tolerance and abs(igst_diff) <= tolerance:
            status = 'Matched'
        elif exact_match_in_2a:
            # Exact invoice found in 2A but amounts differ
            total_diff = cgst_diff + sgst_diff + igst_diff
            status = 'Higher in 2A' if total_diff < -tolerance else 'Lower in 2A'
        else:
            # Invoice not found in 2A at all (fuzzy match didn't produce exact match)
            status = 'Not found in 2A'
        status_lookup[key] = status
        # Store 2A values for keys that found a real match in 2A
        if matched_key:
            vals_2a_lookup[key] = {
                'CGST': gstr_vals['CGST'], 'SGST': gstr_vals['SGST'], 'IGST': gstr_vals['IGST'],
                'TYPE': type_lookup.get(matched_key, ''),
                'BM': booking_month_2a_lookup.get(matched_key, '')
            }

    # Ensure consistency between Status and 2A column values
    for key, status in status_lookup.items():
        if status == 'Not found in 2A':
            # Remove from vals_2a_lookup so 2A columns show the not-found label
            vals_2a_lookup.pop(key, None)
        elif status in ('Matched', 'Higher in 2A', 'Lower in 2A') and key not in vals_2a_lookup:
            # Found in comparison/matching but vals_2a_lookup not yet populated
            if key in gstr_lookup:
                g = gstr_lookup[key]
                vals_2a_lookup[key] = {'CGST': g['CGST'], 'SGST': g['SGST'], 'IGST': g['IGST'], 'TYPE': type_lookup.get(key, ''), 'BM': booking_month_2a_lookup.get(key, '')}
            else:
                # Try same-GSTIN fuzzy match from gstr_by_gstin
                gstin_part = key.split('|', 1)[0]
                cands = gstr_by_gstin.get(gstin_part, [])
                if cands:
                    # Pick the candidate with smallest total tax difference to ITC
                    itc_v = itc_lookup.get(key, {"CGST": 0, "SGST": 0, "IGST": 0})
                    best_c = min(cands, key=lambda c: abs(itc_v["CGST"] - c["CGST"]) + abs(itc_v["SGST"] - c["SGST"]) + abs(itc_v["IGST"] - c["IGST"]))
                    vals_2a_lookup[key] = {'CGST': best_c['CGST'], 'SGST': best_c['SGST'], 'IGST': best_c['IGST'], 'TYPE': type_lookup.get(best_c['key'], ''), 'BM': booking_month_2a_lookup.get(best_c['key'], '')}

    # Map status back to each original ITC row
    result['Status'] = result['GSTINinvoice_norm'].map(status_lookup).fillna('Unmatched')

    # Add 2A columns (CGST/SGST/IGST as per 2A, Type)
    not_found_label = 'Not found in 2A'
    result['CGST as per 2A'] = result['GSTINinvoice_norm'].map(
        lambda k: vals_2a_lookup[k]['CGST'] if k in vals_2a_lookup else not_found_label)
    result['SGST as per 2A'] = result['GSTINinvoice_norm'].map(
        lambda k: vals_2a_lookup[k]['SGST'] if k in vals_2a_lookup else not_found_label)
    result['IGST as per 2A'] = result['GSTINinvoice_norm'].map(
        lambda k: vals_2a_lookup[k]['IGST'] if k in vals_2a_lookup else not_found_label)
    result['Type'] = result['GSTINinvoice_norm'].map(
        lambda k: vals_2a_lookup[k]['TYPE'] if k in vals_2a_lookup else not_found_label)

    # Add Booking Month columns
    # 'Booking Month as per GSTR-2A': from the matched 2A file's booking month/period
    result['Booking Month as per GSTR-2A'] = result['GSTINinvoice_norm'].map(
        lambda k: vals_2a_lookup[k]['BM'] if k in vals_2a_lookup and vals_2a_lookup[k].get('BM') else not_found_label)
    # 'Booking Month as per ITC': from the ITC file's own booking month column
    itc_bm_col = find_booking_month_column(itc_df)
    if itc_bm_col and itc_bm_col in result.columns:
        result['Booking Month as per ITC'] = result[itc_bm_col].astype(str).str.strip()
    else:
        result['Booking Month as per ITC'] = ''

    # Add Remarks column (empty by default, populated during debug matching)
    result['Remarks'] = ''

    # Clean helper columns
    result = result.drop(columns=[c for c in ['GSTINinvoice_norm','_norm_gstin','_norm_inv'] if c in result.columns])

    # Summary
    total = len(result)
    matched_count = (result['Status'] == 'Matched').sum()
    unmatched_count = (result['Status'] == 'Unmatched').sum()
    higher_count = (result['Status'] == 'Higher in 2A').sum()
    lower_count = (result['Status'] == 'Lower in 2A').sum()
    not_found_count = (result['Status'] == 'Not found in 2A').sum()

    if log_callback:
        log_callback(f"ITC Results: Total: {total}, Matched: {matched_count}, Unmatched: {unmatched_count}, Higher in 2A: {higher_count}, Lower in 2A: {lower_count}, Not found in 2A: {not_found_count}")
        log_callback(f"ITC result table created with {total} records (mapped back to original ITC line items)")

    return result


def match_cdnr_negatives(cdnr_df, cdnra_df, itc_result_df, log_callback=None):
    """Match CDNR/CDNRA entries with negative ITC values by GSTIN + tax amount key.

    CDNR/CDNRA have positive tax values; corresponding ITC entries have negative values.
    This creates a composite key (GSTIN + IGST or CGST) to match them.
    Only upgrades unmatched ITC rows to 'Matched' — never downgrades already-matched rows.
    """
    if itc_result_df is None or itc_result_df.empty:
        return itc_result_df

    # Combine CDNR + remaining CDNRA
    cdnr_frames = []
    if cdnr_df is not None and not cdnr_df.empty:
        cdnr_frames.append(cdnr_df)
    if cdnra_df is not None and not cdnra_df.empty:
        cdnr_frames.append(cdnra_df)
    if not cdnr_frames:
        if log_callback:
            log_callback("CDNR matching: No CDNR/CDNRA data available, skipping")
        return itc_result_df

    combined_cdnr = pd.concat(cdnr_frames, ignore_index=True)

    # Find CDNR columns
    cdnr_gstin_col = None
    cdnr_cgst_col, cdnr_sgst_col, cdnr_igst_col = find_tax_amount_columns(combined_cdnr)
    for col in combined_cdnr.columns:
        cl = col.lower().strip()
        if 'gstin' in cl or 'gstn' in cl:
            cdnr_gstin_col = col
            break

    if not cdnr_gstin_col:
        if log_callback:
            log_callback("CDNR matching: Could not find GSTIN column in CDNR, skipping")
        return itc_result_df

    # Find SGST column in CDNR for 2A column values
    cdnr_sgst_col = None
    for col in combined_cdnr.columns:
        cl = col.lower().strip()
        if cl == 'sgst' or ('sgst' in cl and 'amount' in cl):
            cdnr_sgst_col = col
            break

    # Find booking month column in CDNR
    cdnr_bm_col = find_booking_month_column(combined_cdnr)

    # Build CDNR key dict: key → list of {original values} (for one-to-one consumption)
    cdnr_keys = {}
    for idx, row in combined_cdnr.iterrows():
        gstin = normalize_gstin(str(row.get(cdnr_gstin_col, '')))
        igst = safe_numeric_conversion(row.get(cdnr_igst_col, 0)) if cdnr_igst_col else 0
        cgst = safe_numeric_conversion(row.get(cdnr_cgst_col, 0)) if cdnr_cgst_col else 0
        sgst = safe_numeric_conversion(row.get(cdnr_sgst_col, 0)) if cdnr_sgst_col else 0

        # Store original (positive) values for 2A columns
        orig_igst = igst
        orig_cgst = cgst
        orig_sgst = sgst

        # Negate positive values for matching key
        if igst > 0:
            igst = -igst
        if cgst > 0:
            cgst = -cgst

        # Determine key tax: IGST if non-zero, else CGST
        if igst != 0:
            key_tax = igst
        elif cgst != 0:
            key_tax = cgst
        else:
            continue  # both zero, skip

        # Determine source table
        src_type = 'CDNR'
        if idx >= len(cdnr_frames[0]) if len(cdnr_frames) > 1 else False:
            src_type = 'CDNRA'

        key = gstin + '|' + str(round(key_tax, 2))
        if key not in cdnr_keys:
            cdnr_keys[key] = []
        cdnr_bm = str(row.get(cdnr_bm_col, '')).strip() if cdnr_bm_col else ''
        cdnr_keys[key].append({
            'CGST': orig_cgst, 'SGST': orig_sgst, 'IGST': orig_igst, 'TYPE': src_type, 'BM': cdnr_bm
        })

    if not cdnr_keys:
        if log_callback:
            log_callback("CDNR matching: No valid CDNR entries with non-zero tax, skipping")
        return itc_result_df

    # Find ITC columns
    itc_gstin_col = None
    itc_inv_col = None
    itc_cgst_col, itc_sgst_col, itc_igst_col = find_tax_amount_columns(itc_result_df)
    for col in itc_result_df.columns:
        cl = col.lower().strip()
        if 'vendor' in cl and 'gstn' in cl:
            itc_gstin_col = col
        elif 'vendor inv' in cl or 'external doc' in cl:
            itc_inv_col = col

    if not itc_gstin_col or not itc_inv_col:
        if log_callback:
            log_callback("CDNR matching: Could not find GSTIN/Invoice columns in ITC, skipping")
        return itc_result_df

    # Statuses that should NOT be changed (already matched)
    protected_statuses = {'Matched', 'Matched but invoice number is not accurate'}

    # Track which GSTINinvoice_norm keys to upgrade
    matched_norm_keys = set()
    matched_cdnr_info = {}  # inv_key → {'CGST', 'SGST', 'IGST', 'TYPE'}
    result = itc_result_df.copy()

    # Build normalized keys for grouping sibling rows
    result['_norm_gstin'] = result[itc_gstin_col].apply(normalize_gstin)
    result['_norm_inv'] = result[itc_inv_col].apply(normalize_invoice)
    result['_gstin_inv_norm'] = result['_norm_gstin'] + '|' + result['_norm_inv']

    # Process unique invoice groups (not individual line items)
    seen_inv_keys = set()
    for idx, row in result.iterrows():
        inv_key = row['_gstin_inv_norm']
        if inv_key in seen_inv_keys:
            continue
        seen_inv_keys.add(inv_key)

        # Skip already matched rows
        if row.get('Status', '') in protected_statuses:
            continue

        gstin = row['_norm_gstin']
        igst = safe_numeric_conversion(row.get(itc_igst_col, 0)) if itc_igst_col else 0
        cgst = safe_numeric_conversion(row.get(itc_cgst_col, 0)) if itc_cgst_col else 0

        # Only consider rows with negative IGST or CGST (credit note entries)
        if igst >= 0 and cgst >= 0:
            continue

        # Determine key tax: IGST if non-zero, else CGST
        if igst != 0:
            key_tax = igst
        elif cgst != 0:
            key_tax = cgst
        else:
            continue

        key = gstin + '|' + str(round(key_tax, 2))

        # Check if CDNR has a matching entry
        if key in cdnr_keys and len(cdnr_keys[key]) > 0:
            cdnr_info = cdnr_keys[key].pop()  # consume one CDNR entry
            matched_norm_keys.add(inv_key)
            matched_cdnr_info[inv_key] = cdnr_info

    # Update all sibling rows for matched invoices
    has_2a_cols = 'CGST as per 2A' in result.columns
    if matched_norm_keys:
        mask = result['_gstin_inv_norm'].isin(matched_norm_keys)
        result.loc[mask, 'Status'] = 'Matched'
        # Update 2A columns with CDNR values
        if has_2a_cols:
            for inv_key, info in matched_cdnr_info.items():
                row_mask = result['_gstin_inv_norm'] == inv_key
                result.loc[row_mask, 'CGST as per 2A'] = info['CGST']
                result.loc[row_mask, 'SGST as per 2A'] = info['SGST']
                result.loc[row_mask, 'IGST as per 2A'] = info['IGST']
                result.loc[row_mask, 'Type'] = info['TYPE']
                if 'Booking Month as per GSTR-2A' in result.columns and info.get('BM'):
                    result.loc[row_mask, 'Booking Month as per GSTR-2A'] = info['BM']

    # Clean helper columns
    result = result.drop(columns=['_norm_gstin', '_norm_inv', '_gstin_inv_norm'])

    matched_count = len(matched_norm_keys)
    matched_rows = mask.sum() if matched_norm_keys else 0
    if log_callback:
        log_callback(f"CDNR matching: {matched_count} invoices ({matched_rows} rows) matched via CDNR/CDNRA negative matching")

    return result


def create_comprehensive_report(itc_df, itc_register, merged_df, comparison_df, log_callback=None):
    """Step 8: Create comprehensive merged report with all tables side by side"""
    if itc_df is None or itc_df.empty:
        if log_callback:
            log_callback("Error: ITC table is empty")
        return pd.DataFrame()

    invoice_col = None
    for col in itc_df.columns:
        col_lower = col.lower().strip()
        if 'vendor inv' in col_lower or 'external doc' in col_lower:
            invoice_col = col
            break

    if not invoice_col:
        if log_callback:
            log_callback("Error: Could not find Invoice column in ITC")
        return pd.DataFrame()

    itc_columns = {}
    for col in itc_df.columns:
        itc_columns[col] = f"ITC_{col}"
    itc_prepared = itc_df.rename(columns=itc_columns)
    itc_invoice_col = f"ITC_{invoice_col}"

    itc_reg_prepared = itc_register.copy()
    itc_reg_columns = {}
    for col in itc_reg_prepared.columns:
        itc_reg_columns[col] = f"as_per_itc_register_{col}"
    itc_reg_prepared = itc_reg_prepared.rename(columns=itc_reg_columns)

    gstr2a_prepared = merged_df.copy()
    gstr2a_columns = {}
    for col in gstr2a_prepared.columns:
        gstr2a_columns[col] = f"as_per_GSTR-2A_{col}"
    gstr2a_prepared = gstr2a_prepared.rename(columns=gstr2a_columns)

    vendor_gstn_col = None
    vendor_inv_col = None
    for col in itc_df.columns:
        col_lower = col.lower().strip()
        if 'vendor' in col_lower and 'gstn' in col_lower:
            vendor_gstn_col = col
        elif 'vendor inv' in col_lower or 'external doc' in col_lower:
            vendor_inv_col = col

    if vendor_gstn_col and vendor_inv_col:
        itc_prepared['_merge_key'] = itc_df[vendor_gstn_col].astype(str) + itc_df[vendor_inv_col].astype(str)

    itc_reg_prepared['_merge_key'] = itc_register['GSTINinvoice']
    gstr2a_prepared['_merge_key'] = merged_df['GSTN'].astype(str) + merged_df['Document_number'].astype(str)

    comprehensive = pd.merge(
        itc_prepared,
        itc_reg_prepared,
        on='_merge_key',
        how='outer'
    )

    comprehensive = pd.merge(
        comprehensive,
        gstr2a_prepared,
        on='_merge_key',
        how='outer'
    )

    if not comparison_df.empty:
        # Include the reconciliation Status as well so comprehensive report follows the same MATCH/MISMATCH decision
        comparison_subset = comparison_df[['GSTINinvoice', 'CGST_Difference', 'SGST_Difference', 'IGST_Difference', 'Status']].copy()
        comparison_subset = comparison_subset.rename(columns={'GSTINinvoice': '_merge_key', 'Status': 'Reconciliation_Status'})
        comprehensive = pd.merge(
            comprehensive,
            comparison_subset,
            on='_merge_key',
            how='left'
        )

    tolerance = 10.0

    # Derive numeric CGST/SGST/IGST from displayed ITC and GSTR columns (prefer Amount columns if present)
    def _find_col(cols, keyword):
        """Return the best column name from cols containing keyword. Prefer columns with 'amount' else the first match."""
        candidates = [c for c in cols if keyword in c.lower()]
        if not candidates:
            return None
        for c in candidates:
            if 'amount' in c.lower():
                return c
        return candidates[0]

    # Build lists of candidate columns (these are defined later originally but we need them now to pick numeric columns)
    itc_cols = [col for col in comprehensive.columns if col.startswith('ITC_')]
    gstr2a_cols = [col for col in comprehensive.columns if col.startswith('as_per_GSTR-2A_')]

    itc_cgst_col = _find_col(itc_cols, 'cgst')
    itc_sgst_col = _find_col(itc_cols, 'sgst')
    itc_igst_col = _find_col(itc_cols, 'igst')

    gstr_cgst_col = _find_col(gstr2a_cols, 'cgst')
    gstr_sgst_col = _find_col(gstr2a_cols, 'sgst')
    gstr_igst_col = _find_col(gstr2a_cols, 'igst')

    # Compute rounded numeric columns to use for determining Status (HALF_UP rounding)
    if itc_igst_col and gstr_igst_col:
        comprehensive['ITC_IGST_num'] = pd.to_numeric(comprehensive[itc_igst_col].astype(str).str.replace(',',''), errors='coerce').fillna(0).apply(lambda x: round_rupee(x))
        comprehensive['GSTR_IGST_num'] = pd.to_numeric(comprehensive[gstr_igst_col].astype(str).str.replace(',',''), errors='coerce').fillna(0).apply(lambda x: round_rupee(x))
        comprehensive['IGST_Diff_num'] = comprehensive['ITC_IGST_num'] - comprehensive['GSTR_IGST_num']
    else:
        comprehensive['IGST_Diff_num'] = comprehensive.get('IGST_Difference', 0)

    if itc_cgst_col and gstr_cgst_col:
        comprehensive['ITC_CGST_num'] = pd.to_numeric(comprehensive[itc_cgst_col].astype(str).str.replace(',',''), errors='coerce').fillna(0).apply(lambda x: round_rupee(x))
        comprehensive['GSTR_CGST_num'] = pd.to_numeric(comprehensive[gstr_cgst_col].astype(str).str.replace(',',''), errors='coerce').fillna(0).apply(lambda x: round_rupee(x))
        comprehensive['CGST_Diff_num'] = comprehensive['ITC_CGST_num'] - comprehensive['GSTR_CGST_num']
    else:
        comprehensive['CGST_Diff_num'] = comprehensive.get('CGST_Difference', 0)

    if itc_sgst_col and gstr_sgst_col:
        comprehensive['ITC_SGST_num'] = pd.to_numeric(comprehensive[itc_sgst_col].astype(str).str.replace(',',''), errors='coerce').fillna(0).apply(lambda x: round_rupee(x))
        comprehensive['GSTR_SGST_num'] = pd.to_numeric(comprehensive[gstr_sgst_col].astype(str).str.replace(',',''), errors='coerce').fillna(0).apply(lambda x: round_rupee(x))
        comprehensive['SGST_Diff_num'] = comprehensive['ITC_SGST_num'] - comprehensive['GSTR_SGST_num']
    else:
        comprehensive['SGST_Diff_num'] = comprehensive.get('SGST_Difference', 0)

    # Prefer differences from comparison_df (invoice-level) when available, otherwise fallback to per-line numeric diffs
    def _choose_diff(row, comp_col, num_col):
        # comp_col is e.g. 'CGST_Difference' inserted from comparison_df; num_col is per-line computed diff
        comp_val = row.get(comp_col, None)
        if pd.notna(comp_val):
            # Round comp_val to nearest rupee using HALF_UP
            try:
                return round_rupee(comp_val)
            except Exception:
                return 0
        # else, use num_col
        return row.get(num_col, 0) if pd.notna(row.get(num_col, 0)) else 0

    def get_detailed_status(row):
        # If reconciliation already determined a MATCH at invoice level, prefer that decision (consistent with comparison_df)
        recon_status = row.get('Reconciliation_Status', None)
        if recon_status and str(recon_status).upper() == 'MATCH':
            return 'MATCH'

        cgst_diff = _choose_diff(row, 'CGST_Difference', 'CGST_Diff_num')
        sgst_diff = _choose_diff(row, 'SGST_Difference', 'SGST_Diff_num')
        igst_diff = _choose_diff(row, 'IGST_Difference', 'IGST_Diff_num')

        cgst_match = abs(cgst_diff) <= tolerance
        sgst_match = abs(sgst_diff) <= tolerance
        igst_match = abs(igst_diff) <= tolerance

        # If both per-row displayed ITC and GSTR numeric values exist and are effectively equal, prefer them as match
        try:
            itc_c = row.get('ITC_CGST_num', None)
            gstr_c = row.get('GSTR_CGST_num', None)
            if itc_c is not None and gstr_c is not None and abs(itc_c - gstr_c) <= tolerance:
                cgst_match = True
                cgst_diff = 0
        except Exception:
            pass
        try:
            itc_s = row.get('ITC_SGST_num', None)
            gstr_s = row.get('GSTR_SGST_num', None)
            if itc_s is not None and gstr_s is not None and abs(itc_s - gstr_s) <= tolerance:
                sgst_match = True
                sgst_diff = 0
        except Exception:
            pass
        try:
            itc_i = row.get('ITC_IGST_num', None)
            gstr_i = row.get('GSTR_IGST_num', None)
            if itc_i is not None and gstr_i is not None and abs(itc_i - gstr_i) <= tolerance:
                igst_match = True
                igst_diff = 0
        except Exception:
            pass

        if cgst_match and sgst_match and igst_match:
            return 'MATCH'

        mismatches = []
        if not cgst_match:
            if cgst_diff > 0:
                mismatches.append('CGST Lower in 2A')
            else:
                mismatches.append('CGST Higher in 2A')
        if not sgst_match:
            if sgst_diff > 0:
                mismatches.append('SGST Lower in 2A')
            else:
                mismatches.append('SGST Higher in 2A')
        if not igst_match:
            if igst_diff > 0:
                mismatches.append('IGST Lower in 2A')
            else:
                mismatches.append('IGST Higher in 2A')

        if len(mismatches) > 1:
            return 'MISMATCH'
        elif len(mismatches) == 1:
            return mismatches[0]

        return 'MISMATCH'
    comprehensive['Status'] = comprehensive.apply(get_detailed_status, axis=1)

    comprehensive = comprehensive.rename(columns={
        'CGST_Difference': 'CGST',
        'SGST_Difference': 'SGST',
        'IGST_Difference': 'IGST'
    })

    if '_merge_key' in comprehensive.columns:
        comprehensive = comprehensive.drop(columns=['_merge_key'])

    itc_cols = [col for col in comprehensive.columns if col.startswith('ITC_')]
    itc_reg_cols = [col for col in comprehensive.columns if col.startswith('as_per_itc_register_')]
    gstr2a_cols = [col for col in comprehensive.columns if col.startswith('as_per_GSTR-2A_')]
    status_cols = ['Status', 'CGST', 'SGST', 'IGST']
    status_cols = [col for col in status_cols if col in comprehensive.columns]

    ordered_cols = itc_cols + itc_reg_cols + gstr2a_cols + status_cols
    remaining_cols = [col for col in comprehensive.columns if col not in ordered_cols]
    ordered_cols = ordered_cols + remaining_cols

    comprehensive = comprehensive[ordered_cols]

    if log_callback:
        log_callback(f"Step 8: Created comprehensive report with {len(comprehensive)} records")

    return comprehensive


def to_excel(df):
    """Convert dataframe to Excel file for download"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Reconciliation')
    output.seek(0)
    return output.getvalue()


def to_excel_with_highlight(df):
    """Convert dataframe to Excel file with mismatch highlighting"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Comprehensive_Report')

        workbook = writer.book
        worksheet = writer.sheets['Comprehensive_Report']

        from openpyxl.styles import PatternFill
        red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')

        status_col_idx = None
        for idx, col in enumerate(df.columns, 1):
            if col == 'Status':
                status_col_idx = idx
                break

        if status_col_idx:
            for row_idx, row in enumerate(df.itertuples(), 2):
                status_value = getattr(row, 'Status', None)
                if status_value and status_value != 'MATCH':
                    for col_idx in range(1, len(df.columns) + 1):
                        worksheet.cell(row=row_idx, column=col_idx).fill = red_fill

    output.seek(0)
    return output.getvalue()


class GSTReconciliationApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window setup
        self.title("GST Reconciliation Tool")
        self.geometry("1200x800")
        self.minsize(1000, 700)

        # Try to set icon
        try:
            ico_path = get_resource_path("app_icon.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
            else:
                icon_path = get_resource_path("logo small.png")
                if os.path.exists(icon_path):
                    icon_image = Image.open(icon_path)
                    self.iconphoto(True, ctk.CTkImage(light_image=icon_image, size=(32, 32))._light_image)
        except Exception:
            pass

        # File paths storage
        self.file_paths = {
            'B2B': None,
            'B2BA': None,
            'CDNR': None,
            'CDNRA': None,
            'IMPG': None,
            'IMPGSEZ': None,
            'ITC': None
        }

        # Results storage
        self.comparison_df = None
        self.itc_result_df = None
        self.unmatched_2a_df = None
        self.merged_df = None

        # Create main layout
        self.create_widgets()

    def create_widgets(self):
        # Main scrollable container
        self.main_container = ctk.CTkScrollableFrame(self)
        self.main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # Header with logo
        header_frame = ctk.CTkFrame(self.main_container, fg_color=THEME_PRIMARY, corner_radius=10)
        header_frame.pack(fill="x", padx=5, pady=5)

        # Header content frame (for logo and title side by side)
        header_content = ctk.CTkFrame(header_frame, fg_color="transparent")
        header_content.pack(pady=15)

        # Load and display logo
        try:
            logo_path = get_resource_path("logo small.png")
            if os.path.exists(logo_path):
                logo_image = Image.open(logo_path)
                self.logo_ctk = ctk.CTkImage(light_image=logo_image, dark_image=logo_image, size=(50, 50))
                logo_label = ctk.CTkLabel(header_content, image=self.logo_ctk, text="")
                logo_label.pack(side="left", padx=10)
        except Exception:
            pass

        title_label = ctk.CTkLabel(
            header_content,
            text="GST Reconciliation Tool",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="white"
        )
        title_label.pack(side="left", padx=10)

        # Instructions frame
        instructions_frame = ctk.CTkFrame(self.main_container, fg_color=THEME_LIGHT, corner_radius=10)
        instructions_frame.pack(fill="x", padx=5, pady=5)

        instructions_text = (
            "HOW TO USE THIS APP\n\n"
            "Step 1 :  Prepare your files\n"
            "    - Download your GSTR-2A data from the GST portal (B2B, B2BA, CDNR, CDNRA, IMPG, IMPGSEZ).\n"
            "    - Keep your Purchase Register / ITC Register ready.\n"
            "    - Make sure each file has columns named  'Invoice No',  'GSTIN',  'CGST',  'SGST'  and  'IGST'.\n"
            "    - Tip: Download the Template below and paste your data into it for an error-free experience.\n\n"
            "Step 2 :  Upload your files\n"
            "    - Option A  -  Upload CSV files one by one using the Browse buttons below.\n"
            "    - Option B  -  Put all sheets in one Excel file and click  'Upload Excel File (All Sheets)'.\n"
            "         Sheets must be named B2B, B2BA, CDNR, CDNRA, IMPG, IMPGSEZ, ITC.\n\n"
            "Step 3 :  Click  'Process and Reconcile'\n"
            "    - The app will match your ITC Register against GSTR-2A and show the results.\n"
            "    - You can then view or download the reconciliation report.\n\n"
            "Which file comes from where?\n"
            "    From GSTR-2A (GST Portal)  :  2A-B2B, 2A-B2BA, 2A-CDNR, 2A-CDNRA, 2A-IMPG, 2A-IMPGSEZ\n"
            "    From your records               :  Register-ITC  (your Purchase / ITC Register)"
        )
        instructions_label = ctk.CTkLabel(
            instructions_frame,
            text=instructions_text,
            font=ctk.CTkFont(size=12),
            justify="left",
            text_color=THEME_DARK,
            anchor="w"
        )
        instructions_label.pack(pady=10, padx=15, fill="x")

        # Template buttons frame
        template_buttons_frame = ctk.CTkFrame(instructions_frame, fg_color="transparent")
        template_buttons_frame.pack(pady=10)

        # Template download button
        template_btn = ctk.CTkButton(
            template_buttons_frame,
            text="Download Template.xlsx",
            command=self.download_template,
            fg_color=THEME_PRIMARY,
            hover_color=THEME_HOVER,
            width=200
        )
        template_btn.pack(side="left", padx=5)

        # Upload Excel template button
        upload_excel_btn = ctk.CTkButton(
            template_buttons_frame,
            text="Upload Excel File (All Sheets)",
            command=self.upload_excel_template,
            fg_color="#9C27B0",  # Purple for variety
            hover_color="#7B1FA2",
            width=220
        )
        upload_excel_btn.pack(side="left", padx=5)

        # Clear All button
        clear_all_btn = ctk.CTkButton(
            template_buttons_frame,
            text="Clear All",
            command=self.clear_all_files,
            fg_color="#757575",  # Gray
            hover_color="#616161",
            width=100
        )
        clear_all_btn.pack(side="left", padx=5)

        # File upload section
        upload_frame = ctk.CTkFrame(self.main_container)
        upload_frame.pack(fill="x", padx=5, pady=10)

        upload_title = ctk.CTkLabel(
            upload_frame,
            text="Upload Files",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        upload_title.pack(pady=10)

        # Display name mapping (internal key → UI label)
        self.display_names = {
            'B2B': '2A-B2B',
            'B2BA': '2A-B2BA',
            'CDNR': '2A-CDNR',
            'CDNRA': '2A-CDNRA',
            'IMPG': '2A-IMPG',
            'IMPGSEZ': '2A-IMPGSEZ',
            'ITC': 'Register-ITC'
        }

        # Create two columns for file uploads
        files_container = ctk.CTkFrame(upload_frame, fg_color="transparent")
        files_container.pack(fill="x", padx=10, pady=5)

        # Left column
        left_col = ctk.CTkFrame(files_container, fg_color="transparent")
        left_col.pack(side="left", fill="both", expand=True, padx=5)

        # Right column
        right_col = ctk.CTkFrame(files_container, fg_color="transparent")
        right_col.pack(side="right", fill="both", expand=True, padx=5)

        # File upload buttons
        self.file_labels = {}
        left_files = ['B2B', 'B2BA', 'CDNR', 'CDNRA']
        right_files = ['IMPG', 'IMPGSEZ', 'ITC']

        for i, file_type in enumerate(left_files):
            self.create_file_upload(left_col, file_type, i + 1)

        for i, file_type in enumerate(right_files):
            self.create_file_upload(right_col, file_type, i + 5)

        # Process button
        self.process_btn = ctk.CTkButton(
            self.main_container,
            text="Process and Reconcile",
            command=self.start_processing,
            fg_color=THEME_PRIMARY,
            hover_color=THEME_HOVER,
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50
        )
        self.process_btn.pack(fill="x", padx=5, pady=10)

        # Progress section
        self.progress_frame = ctk.CTkFrame(self.main_container)
        self.progress_frame.pack(fill="x", padx=5, pady=5)

        self.progress_label = ctk.CTkLabel(
            self.progress_frame,
            text="Ready to process",
            font=ctk.CTkFont(size=12)
        )
        self.progress_label.pack(pady=5)

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, width=400, progress_color=THEME_PRIMARY)
        self.progress_bar.pack(pady=5)
        self.progress_bar.set(0)

        # Log area
        log_frame = ctk.CTkFrame(self.main_container)
        log_frame.pack(fill="x", padx=5, pady=5)

        log_title = ctk.CTkLabel(
            log_frame,
            text="Processing Log",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        log_title.pack(pady=5)

        self.log_text = ctk.CTkTextbox(log_frame, height=200, font=ctk.CTkFont(size=11))
        self.log_text.pack(fill="x", padx=5, pady=5)

        # Results buttons frame (initially hidden)
        self.results_frame = ctk.CTkFrame(self.main_container)

        results_label = ctk.CTkLabel(
            self.results_frame,
            text="Download Results",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        results_label.pack(pady=5)

        buttons_container = ctk.CTkFrame(self.results_frame, fg_color="transparent")
        buttons_container.pack(fill="x", pady=5)

        self.download_reconciliation_btn = ctk.CTkButton(
            buttons_container,
            text="Download ITC Results",
            command=self.download_reconciliation,
            fg_color=THEME_PRIMARY,
            hover_color=THEME_HOVER,
            width=250
        )
        self.download_reconciliation_btn.pack(side="left", padx=10, pady=5)

        self.download_unmatched_itc_btn = ctk.CTkButton(
            buttons_container,
            text="Download Unmatched in ITC",
            command=self.download_unmatched_itc,
            fg_color="#E65100",
            hover_color="#BF360C",
            width=250
        )
        self.download_unmatched_itc_btn.pack(side="left", padx=10, pady=5)

        self.download_unmatched_2a_btn = ctk.CTkButton(
            buttons_container,
            text="Download Unmatched in 2A",
            command=self.download_unmatched_2a,
            fg_color="#1565C0",
            hover_color="#0D47A1",
            width=250
        )
        self.download_unmatched_2a_btn.pack(side="left", padx=10, pady=5)

        self.view_results_btn = ctk.CTkButton(
            buttons_container,
            text="View Results",
            command=self.view_results,
            fg_color="#9C27B0",
            hover_color="#7B1FA2",
            width=150
        )
        self.view_results_btn.pack(side="left", padx=10, pady=5)

        self.debug_matching_btn = ctk.CTkButton(
            buttons_container,
            text="Debug Matching",
            command=self.debug_matching,
            fg_color="#FF6F00",
            hover_color="#E65100",
            width=180
        )
        self.debug_matching_btn.pack(side="left", padx=10, pady=5)

        # Footer
        footer_frame = ctk.CTkFrame(self.main_container, fg_color=THEME_DARK, corner_radius=10)
        footer_frame.pack(fill="x", padx=5, pady=5)

        footer_label = ctk.CTkLabel(
            footer_frame,
            text="Copyright (c) GSC in time 2026 | Contact: info@gscintime.com | https://www.gscintime.com | +91-22-4612 5600",
            font=ctk.CTkFont(size=11),
            text_color="white"
        )
        footer_label.pack(pady=10)

    def create_file_upload(self, parent, file_type, number):
        """Create a file upload row"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=3)

        display_name = self.display_names.get(file_type, file_type)
        label = ctk.CTkLabel(
            frame,
            text=f"{number}. {display_name}:",
            font=ctk.CTkFont(size=12),
            width=110
        )
        label.pack(side="left", padx=5)

        self.file_labels[file_type] = ctk.CTkLabel(
            frame,
            text="No file selected",
            font=ctk.CTkFont(size=11),
            text_color="gray",
            width=200
        )
        self.file_labels[file_type].pack(side="left", padx=5, fill="x", expand=True)

        btn = ctk.CTkButton(
            frame,
            text="Browse",
            command=lambda ft=file_type: self.browse_file(ft),
            width=80,
            height=28,
            fg_color=THEME_PRIMARY,
            hover_color=THEME_HOVER
        )
        btn.pack(side="right", padx=5)

        clear_btn = ctk.CTkButton(
            frame,
            text="X",
            command=lambda ft=file_type: self.clear_file(ft),
            width=30,
            height=28,
            fg_color="#757575",
            hover_color="#616161"
        )
        clear_btn.pack(side="right", padx=2)

    def browse_file(self, file_type):
        """Open file dialog to select a CSV file"""
        display_name = self.display_names.get(file_type, file_type)
        filepath = filedialog.askopenfilename(
            title=f"Select {display_name} CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filepath:
            self.file_paths[file_type] = filepath
            filename = os.path.basename(filepath)
            self.file_labels[file_type].configure(text=filename, text_color=THEME_PRIMARY)
            self.log(f"Selected {display_name}: {filename}")

    def clear_file(self, file_type):
        """Clear selected file"""
        self.file_paths[file_type] = None
        self.file_labels[file_type].configure(text="No file selected", text_color="gray")

    def clear_all_files(self):
        """Clear all selected files"""
        for file_type in self.file_paths.keys():
            self.file_paths[file_type] = None
            self.file_labels[file_type].configure(text="No file selected", text_color="gray")
        self.log("All files cleared.")

    def log(self, message):
        """Add message to log"""
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def download_template(self):
        """Download template Excel file - copies the existing Template all.xlsx"""
        try:
            # Try to find the existing template file
            template_source = get_resource_path("Template all.xlsx")

            if not os.path.exists(template_source):
                # Fallback to Template.xlsx
                template_source = get_resource_path("Template.xlsx")

            filepath = filedialog.asksaveasfilename(
                title="Save Template",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="Template.xlsx"
            )
            if filepath:
                if os.path.exists(template_source):
                    # Copy existing template file
                    import shutil
                    shutil.copy2(template_source, filepath)
                else:
                    # Generate template if source doesn't exist
                    template_data = create_template_excel()
                    with open(filepath, 'wb') as f:
                        f.write(template_data)
                self.log(f"Template saved to: {filepath}")
                messagebox.showinfo("Success", f"Template saved successfully!\n\nLocation: {filepath}")
        except Exception as e:
            self.log(f"Error saving template: {str(e)}")
            messagebox.showerror("Error", f"Failed to save template: {str(e)}")

    def upload_excel_template(self):
        """Upload a single Excel file with all sheets (B2B, B2BA, CDNR, CDNRA, IMPG, IMPGSEZ, ITC)"""
        filepath = filedialog.askopenfilename(
            title="Select Excel File with All Sheets",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filepath:
            try:
                # Read all sheets from the Excel file
                excel_file = pd.ExcelFile(filepath)
                sheet_names = excel_file.sheet_names
                self.log(f"Loading Excel file: {os.path.basename(filepath)}")
                self.log(f"Found sheets: {', '.join(sheet_names)}")

                # Map sheet names to file types (case-insensitive)
                sheet_mapping = {}
                for sheet in sheet_names:
                    sheet_upper = sheet.upper().strip()
                    if sheet_upper in ['B2B', 'B2BA', 'CDNR', 'CDNRA', 'IMPG', 'IMPGSEZ', 'ITC']:
                        sheet_mapping[sheet_upper] = sheet

                # Load each recognized sheet
                loaded_count = 0
                for file_type in ['B2B', 'B2BA', 'CDNR', 'CDNRA', 'IMPG', 'IMPGSEZ', 'ITC']:
                    if file_type in sheet_mapping:
                        # Store the Excel file path and sheet name
                        self.file_paths[file_type] = (filepath, sheet_mapping[file_type])
                        self.file_labels[file_type].configure(
                            text=f"{os.path.basename(filepath)} [{sheet_mapping[file_type]}]",
                            text_color=THEME_PRIMARY
                        )
                        loaded_count += 1
                        dn = self.display_names.get(file_type, file_type)
                        self.log(f"  - {dn}: Loaded from sheet '{sheet_mapping[file_type]}'")

                if loaded_count > 0:
                    messagebox.showinfo("Success", f"Loaded {loaded_count} sheets from Excel file!\n\nSheets found: {', '.join(sheet_mapping.keys())}")
                else:
                    messagebox.showwarning("Warning", "No recognized sheets found!\n\nExpected sheet names: B2B, B2BA, CDNR, CDNRA, IMPG, IMPGSEZ, ITC")

            except Exception as e:
                self.log(f"Error loading Excel file: {str(e)}")
                messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")

    def start_processing(self):
        """Start the reconciliation process in a separate thread"""
        # Check if ITC is provided (can be string path or tuple for Excel)
        if not self.file_paths['ITC']:
            messagebox.showerror("Error", "Register-ITC file/sheet is mandatory!")
            return

        self.process_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.log_text.delete("1.0", "end")
        self.results_frame.pack_forget()

        # Run processing in a thread
        thread = threading.Thread(target=self.process_files)
        thread.start()

    def process_files(self):
        """Process all files and perform reconciliation"""
        try:
            tables = {}

            # Load files
            self.update_progress(0.1, "Loading files...")
            for name, filepath in self.file_paths.items():
                dn = self.display_names.get(name, name)
                if filepath:
                    try:
                        # Check if it's a tuple (Excel file with sheet name) or string (CSV file)
                        if isinstance(filepath, tuple):
                            # Excel file: (filepath, sheet_name)
                            excel_path, sheet_name = filepath
                            # Read Excel with all columns as strings first (like CSV does)
                            # This ensures consistent behavior between Excel and CSV
                            df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)
                            # Replace 'nan' and 'None' strings with empty string
                            df = df.fillna('')
                            df = df.replace('nan', '')
                            df = df.replace('None', '')
                            tables[name] = df
                            self.log(f"Loaded {dn} from Excel sheet '{sheet_name}': {len(tables[name])} rows")
                        else:
                            # CSV file
                            try:
                                tables[name] = pd.read_csv(filepath, encoding='utf-8', dtype=str)
                                tables[name] = tables[name].fillna('')
                                self.log(f"Loaded {dn}: {len(tables[name])} rows")
                            except:
                                tables[name] = pd.read_csv(filepath, encoding='latin-1', dtype=str)
                                tables[name] = tables[name].fillna('')
                                self.log(f"Loaded {dn}: {len(tables[name])} rows")
                    except Exception as e:
                        self.log(f"Error loading {dn}: {str(e)}")
                        tables[name] = pd.DataFrame()
                else:
                    tables[name] = pd.DataFrame()

            # Normalize columns
            self.update_progress(0.2, "Normalizing columns...")
            if 'ITC' in tables and not tables['ITC'].empty:
                tables['ITC'] = normalize_itc_columns(tables['ITC'])
            if 'CDNR' in tables and not tables['CDNR'].empty:
                tables['CDNR'] = normalize_cdnr_columns(tables['CDNR'])
            if 'CDNRA' in tables and not tables['CDNRA'].empty:
                tables['CDNRA'] = normalize_cdnr_columns(tables['CDNRA'])

            # Step 1: Merge duplicates (keep original ITC to preserve line items)
            self.update_progress(0.3, "Step 1: Merging duplicates...")
            # Keep a copy of the original ITC (line items) before collapsing duplicates for matching
            original_itc = tables['ITC'].copy()
            tables['ITC'] = merge_duplicate_vendor_invoices(tables['ITC'], self.log)  # merged per invoice

            # Step 2: Match B2B with B2BA
            self.update_progress(0.4, "Step 2: Matching B2B with B2BA...")
            if 'B2B' in tables and 'B2BA' in tables:
                b2b_inv_col = None
                b2ba_inv_col = None

                for col in tables['B2B'].columns:
                    if col.lower().strip() == 'invoice no':
                        b2b_inv_col = col
                        break

                if not tables['B2BA'].empty:
                    for col in tables['B2BA'].columns:
                        if col.lower().strip() == 'invoice no':
                            b2ba_inv_col = col
                            break

                if b2b_inv_col and b2ba_inv_col:
                    tables['B2B'], tables['B2BA'] = match_and_update(
                        tables['B2B'], tables['B2BA'],
                        b2b_inv_col, b2ba_inv_col, 'B2B', self.log
                    )

            # Step 3: Match CDNR with CDNRA
            self.update_progress(0.5, "Step 3: Matching CDNR with CDNRA...")
            if 'CDNR' in tables and 'CDNRA' in tables:
                cdnr_inv_col = None
                cdnra_inv_col = None

                if not tables['CDNR'].empty:
                    for col in tables['CDNR'].columns:
                        cl = col.lower().strip()
                        if cl == 'note no' or cl == 'invoice no' or cl == 'boe no':
                            cdnr_inv_col = col
                            break

                if not tables['CDNRA'].empty:
                    for col in tables['CDNRA'].columns:
                        cl = col.lower().strip()
                        if cl == 'note no' or cl == 'invoice no' or cl == 'boe no':
                            cdnra_inv_col = col
                            break

                if cdnr_inv_col and cdnra_inv_col:
                    tables['CDNR'], tables['CDNRA'] = match_and_update(
                        tables['CDNR'], tables['CDNRA'],
                        cdnr_inv_col, cdnra_inv_col, 'CDNR', self.log
                    )

            # Step 4: Create MERGED table
            self.update_progress(0.6, "Step 4: Creating MERGED table...")
            merged_df = create_merged_table(tables, self.log)
            self.merged_df = merged_df

            # Step 5: Create ITC register
            self.update_progress(0.7, "Step 5: Creating ITC register...")
            # Use merged ITC for register (invoice-level sums)
            itc_register = create_itc_register(tables['ITC'], self.log)

            # Step 6: Create GSTR 2A table
            self.update_progress(0.8, "Step 6: Creating GSTR-2A table...")
            gstr_2a = create_gstr_2a(merged_df, self.log)

            # Step 7: Compare tables
            self.update_progress(0.85, "Step 7: Comparing tables...")
            self.comparison_df = compare_tables(itc_register, gstr_2a, self.log)

            # Create ITC result table with Status column
            self.update_progress(0.9, "Creating ITC results with Status...")
            # Pass original (pre-merged) ITC so all original line items are present in results
            # Use the aggregated comparison dataframe to ensure ITC Results reflect reconciliation (match counts)
            self.itc_result_df = create_itc_result(original_itc, itc_register, gstr_2a, self.comparison_df, merged_df, self.log)

            # Step 7.5: CDNR/CDNRA negative value matching
            self.update_progress(0.92, "Step 7.5: Matching CDNR/CDNRA negatives with ITC...")
            self.itc_result_df = match_cdnr_negatives(
                tables.get('CDNR', pd.DataFrame()),
                tables.get('CDNRA', pd.DataFrame()),
                self.itc_result_df, self.log
            )

            # Compute unmatched GSTR 2A items (in merged/2A but not in ITC)
            self.update_progress(0.95, "Finding unmatched 2A items...")
            self.unmatched_2a_df = self._compute_unmatched_2a(merged_df, original_itc)

            # Complete
            self.update_progress(1.0, "Processing completed!")
            self.log("\n" + "=" * 50)
            self.log("PROCESSING COMPLETED SUCCESSFULLY!")
            self.log("=" * 50)

            if self.itc_result_df is not None and not self.itc_result_df.empty:
                total = len(self.itc_result_df)
                matched = len(self.itc_result_df[self.itc_result_df['Status'] == 'Matched'])
                unmatched = len(self.itc_result_df[self.itc_result_df['Status'] == 'Unmatched'])
                higher = len(self.itc_result_df[self.itc_result_df['Status'] == 'Higher in 2A'])
                lower = len(self.itc_result_df[self.itc_result_df['Status'] == 'Lower in 2A'])
                not_found = len(self.itc_result_df[self.itc_result_df['Status'] == 'Not found in 2A'])
                self.log(f"\nITC RESULTS SUMMARY:")
                self.log(f"  Total Records: {total}")
                self.log(f"  Matched: {matched} ({matched/total*100:.1f}%)")
                self.log(f"  Unmatched: {unmatched} ({unmatched/total*100:.1f}%)")
                self.log(f"  Higher in 2A: {higher} ({higher/total*100:.1f}%)")
                self.log(f"  Lower in 2A: {lower} ({lower/total*100:.1f}%)")
                self.log(f"  Not found in 2A: {not_found} ({not_found/total*100:.1f}%)")

            # Show results frame
            self.after(0, self.show_results_frame)

        except Exception as e:
            self.log(f"\nError during processing: {str(e)}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.after(0, lambda: self.process_btn.configure(state="normal"))

    def update_progress(self, value, message):
        """Update progress bar and label"""
        self.after(0, lambda: self.progress_bar.set(value))
        self.after(0, lambda: self.progress_label.configure(text=message))
        self.log(message)

    def show_results_frame(self):
        """Show the results download frame"""
        self.results_frame.pack(fill="x", padx=5, pady=5, before=self.progress_frame)

    def _compute_unmatched_2a(self, merged_df, original_itc):
        """Find rows in merged/GSTR 2A that have no matching invoice in ITC."""
        if merged_df is None or merged_df.empty:
            return pd.DataFrame()
        if original_itc is None or original_itc.empty:
            return merged_df.copy()

        # Build set of normalized GSTIN+Invoice keys from ITC
        vendor_gstn_col = None
        vendor_inv_col = None
        for col in original_itc.columns:
            cl = col.lower().strip()
            if 'vendor' in cl and 'gstn' in cl:
                vendor_gstn_col = col
            elif 'vendor inv' in cl or 'external doc' in cl:
                vendor_inv_col = col

        if not vendor_gstn_col or not vendor_inv_col:
            return pd.DataFrame()

        itc_keys = set()
        for _, row in original_itc.iterrows():
            key = normalize_gstin(row[vendor_gstn_col]) + '|' + normalize_invoice(row[vendor_inv_col])
            itc_keys.add(key)

        # Check each merged row
        merged_keys = merged_df.apply(
            lambda r: normalize_gstin(str(r.get('GSTN', ''))) + '|' + normalize_invoice(str(r.get('Document_number', ''))),
            axis=1
        )
        mask_unmatched = ~merged_keys.isin(itc_keys)
        return merged_df[mask_unmatched].reset_index(drop=True)

    def download_reconciliation(self):
        """Download ITC results report with unmatched sheets"""
        download_df = self.itc_result_df if self.itc_result_df is not None and not self.itc_result_df.empty else self.comparison_df
        if download_df is None or download_df.empty:
            messagebox.showwarning("Warning", "No data available!")
            return

        try:
            filepath = filedialog.asksaveasfilename(
                title="Save ITC Results",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="GST_ITC_Results.xlsx"
            )
            if filepath:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    download_df.to_excel(writer, sheet_name='ITC_Results', index=False)

                    # Unmatched in ITC (includes Unmatched, Higher in 2A, Lower in 2A)
                    if self.itc_result_df is not None and 'Status' in self.itc_result_df.columns:
                        unmatched_itc = self.itc_result_df[self.itc_result_df['Status'].isin(['Unmatched', 'Higher in 2A', 'Lower in 2A', 'Not found in 2A'])]
                        if not unmatched_itc.empty:
                            unmatched_itc.to_excel(writer, sheet_name='Unmatched_in_ITC', index=False)

                    # Unmatched in 2A
                    if self.unmatched_2a_df is not None and not self.unmatched_2a_df.empty:
                        self.unmatched_2a_df.to_excel(writer, sheet_name='Unmatched_in_2A', index=False)

                output.seek(0)
                with open(filepath, 'wb') as f:
                    f.write(output.getvalue())
                self.log(f"ITC Results saved to: {filepath}")
                messagebox.showinfo("Success", f"ITC Results saved successfully!\n\nLocation: {filepath}")
        except Exception as e:
            self.log(f"Error saving report: {str(e)}")
            messagebox.showerror("Error", f"Failed to save report: {str(e)}")

    def download_unmatched_itc(self):
        """Download unmatched ITC items as Excel"""
        if self.itc_result_df is None or self.itc_result_df.empty:
            messagebox.showwarning("Warning", "No ITC data available!")
            return

        unmatched_itc = self.itc_result_df[self.itc_result_df['Status'].isin(['Unmatched', 'Higher in 2A', 'Lower in 2A', 'Not found in 2A'])]
        if unmatched_itc.empty:
            messagebox.showinfo("Info", "No unmatched ITC items found!")
            return

        try:
            filepath = filedialog.asksaveasfilename(
                title="Save Unmatched in ITC",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="Unmatched_in_ITC.xlsx"
            )
            if filepath:
                unmatched_itc.to_excel(filepath, index=False, engine='openpyxl')
                self.log(f"Unmatched in ITC saved to: {filepath}")
                messagebox.showinfo("Success", f"Unmatched in ITC saved successfully!\n\nRecords: {len(unmatched_itc)}\nLocation: {filepath}")
        except Exception as e:
            self.log(f"Error saving report: {str(e)}")
            messagebox.showerror("Error", f"Failed to save report: {str(e)}")

    def download_unmatched_2a(self):
        """Download unmatched GSTR 2A items as Excel"""
        if self.unmatched_2a_df is None or self.unmatched_2a_df.empty:
            messagebox.showinfo("Info", "No unmatched GSTR 2A items found!")
            return

        try:
            filepath = filedialog.asksaveasfilename(
                title="Save Unmatched in 2A",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="Unmatched_in_2A.xlsx"
            )
            if filepath:
                self.unmatched_2a_df.to_excel(filepath, index=False, engine='openpyxl')
                self.log(f"Unmatched in 2A saved to: {filepath}")
                messagebox.showinfo("Success", f"Unmatched in 2A saved successfully!\n\nRecords: {len(self.unmatched_2a_df)}\nLocation: {filepath}")
        except Exception as e:
            self.log(f"Error saving report: {str(e)}")
            messagebox.showerror("Error", f"Failed to save report: {str(e)}")

    def find_debug_candidate_pairs(self):
        """Find ITC-2A pairs where GSTIN and tax amounts match but invoice numbers only partially match."""
        candidates = []

        if self.itc_result_df is None or self.itc_result_df.empty:
            return candidates
        if self.unmatched_2a_df is None or self.unmatched_2a_df.empty:
            return candidates

        # Find ITC columns
        vendor_gstn_col = None
        vendor_inv_col = None
        for col in self.itc_result_df.columns:
            col_lower = col.lower().strip()
            if 'vendor' in col_lower and 'gstn' in col_lower:
                vendor_gstn_col = col
            elif 'vendor inv' in col_lower or 'external doc' in col_lower:
                vendor_inv_col = col

        itc_cgst_col, itc_sgst_col, itc_igst_col = find_tax_amount_columns(self.itc_result_df)

        if not vendor_gstn_col or not vendor_inv_col:
            return candidates

        # Find taxable value and invoice date columns in ITC
        taxable_col = None
        itc_date_col = None
        for col in self.itc_result_df.columns:
            if 'taxable' in col.lower() and 'value' in col.lower() and not taxable_col:
                taxable_col = col
            if col.lower().strip() == 'invoice date' and not itc_date_col:
                itc_date_col = col

        # Filter to only truly unmatched ITC rows (not found in 2A at all)
        unmatched_itc = self.itc_result_df[
            self.itc_result_df['Status'].isin(['Unmatched', 'Not found in 2A'])
        ]

        if unmatched_itc.empty:
            return candidates

        # Build 2A lookup from unmatched 2A rows only, grouped by normalized GSTIN
        from collections import defaultdict
        twoa_by_gstin = defaultdict(list)
        for idx, row in self.unmatched_2a_df.iterrows():
            twoa_by_gstin[normalize_gstin(str(row.get('GSTN', '')))].append({
                'idx': idx,
                'norm_inv': normalize_invoice(str(row.get('Document_number', ''))),
                'raw_inv': str(row.get('Document_number', '')),
                'gstn': str(row.get('GSTN', '')),
                'cgst': safe_numeric_conversion(row.get('CGST', 0)),
                'sgst': safe_numeric_conversion(row.get('SGST', 0)),
                'igst': safe_numeric_conversion(row.get('IGST', 0)),
                'tax': safe_numeric_conversion(row.get('TAX', 0)),
                'date': str(row.get('Invoice_Date', '')),
            })

        used_2a = set()

        for itc_idx, itc_row in unmatched_itc.iterrows():
            itc_gstin_norm = normalize_gstin(str(itc_row[vendor_gstn_col]))
            itc_inv_norm = normalize_invoice(str(itc_row[vendor_inv_col]))
            itc_cgst = safe_numeric_conversion(itc_row.get(itc_cgst_col, 0)) if itc_cgst_col else 0
            itc_sgst = safe_numeric_conversion(itc_row.get(itc_sgst_col, 0)) if itc_sgst_col else 0
            itc_igst = safe_numeric_conversion(itc_row.get(itc_igst_col, 0)) if itc_igst_col else 0
            itc_taxable = safe_numeric_conversion(itc_row.get(taxable_col, 0)) if taxable_col else 0

            best = None
            best_sim = 0.0

            for cand in twoa_by_gstin.get(itc_gstin_norm, []):
                if cand['idx'] in used_2a:
                    continue
                # Taxable value, IGST, CGST, SGST must match within Rs 10 tolerance
                if (abs(itc_taxable - cand['tax']) > 10 or
                    abs(itc_igst - cand['igst']) > 10 or
                    abs(itc_cgst - cand['cgst']) > 10 or
                    abs(itc_sgst - cand['sgst']) > 10):
                    continue
                # Invoice similarity >= 10%
                sim = similarity(itc_inv_norm, cand['norm_inv'])
                if sim >= 0.1 and sim > best_sim:
                    best_sim = sim
                    best = cand

            if best is not None:
                used_2a.add(best['idx'])
                candidates.append({
                    'itc_index': itc_idx,
                    'itc_invoice': str(itc_row[vendor_inv_col]),
                    'itc_gstin': str(itc_row[vendor_gstn_col]),
                    'itc_date': str(itc_row[itc_date_col]) if itc_date_col else '',
                    'itc_taxable': itc_taxable,
                    'itc_cgst': itc_cgst,
                    'itc_sgst': itc_sgst,
                    'itc_igst': itc_igst,
                    'twoa_invoice': best['raw_inv'],
                    'twoa_gstin': best['gstn'],
                    'twoa_date': best['date'],
                    'twoa_taxable': best['tax'],
                    'twoa_cgst': best['cgst'],
                    'twoa_sgst': best['sgst'],
                    'twoa_igst': best['igst'],
                    'similarity': best_sim,
                    'itc_inv_norm': itc_inv_norm,
                })

        candidates.sort(key=lambda x: x['similarity'], reverse=True)
        return candidates

    def debug_matching(self):
        """Open debug matching window to review partial invoice matches."""
        if self.itc_result_df is None or self.itc_result_df.empty:
            messagebox.showwarning("Warning", "No ITC results available! Run reconciliation first.")
            return
        if self.unmatched_2a_df is None or self.unmatched_2a_df.empty:
            messagebox.showwarning("Warning", "No unmatched 2A data available!")
            return

        candidates = self.find_debug_candidate_pairs()

        if not candidates:
            messagebox.showinfo("Debug Matching",
                                "No candidates found.\n\n"
                                "This means there are no unmatched pairs where:\n"
                                "- GSTIN matches exactly\n"
                                "- Taxable Value, IGST, CGST, SGST match within Rs 10\n"
                                "- Invoice numbers are at least 10% similar")
            return

        self._open_debug_window(candidates)

    def _open_debug_window(self, candidates):
        """Display candidate pairs one at a time for user to match or skip."""
        debug_win = ctk.CTkToplevel(self)
        debug_win.title("Debug Matching - Review Partial Invoice Matches")
        debug_win.geometry("900x500")
        debug_win.lift()
        debug_win.focus_force()
        debug_win.after(100, lambda: debug_win.attributes('-topmost', False))
        debug_win.attributes('-topmost', True)

        state = {'idx': 0, 'matched': 0, 'skipped': 0, 'total': len(candidates)}

        # Find vendor columns for sibling-row updates
        vendor_gstn_col = None
        vendor_inv_col = None
        for col in self.itc_result_df.columns:
            cl = col.lower().strip()
            if 'vendor' in cl and 'gstn' in cl:
                vendor_gstn_col = col
            elif 'vendor inv' in cl or 'external doc' in cl:
                vendor_inv_col = col

        # Header
        header = ctk.CTkFrame(debug_win, fg_color="#FFF3E0")
        header.pack(fill="x", padx=10, pady=(10, 5))

        progress_lbl = ctk.CTkLabel(header, text=f"Pair 1 of {state['total']}",
                                    font=ctk.CTkFont(size=14, weight="bold"))
        progress_lbl.pack(pady=(5, 0))

        sim_lbl = ctk.CTkLabel(header, text="", font=ctk.CTkFont(size=12), text_color="#E65100")
        sim_lbl.pack(pady=(0, 5))

        # Field labels
        fields = ['Invoice No', 'Invoice Date', 'GSTIN', 'Taxable Value', 'IGST', 'CGST', 'SGST']

        # ITC panel
        itc_frame = ctk.CTkFrame(debug_win)
        itc_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(itc_frame, text="ITC Row", font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="#E91E63").pack(anchor="w", padx=10, pady=(5, 0))
        itc_grid = ctk.CTkFrame(itc_frame, fg_color="transparent")
        itc_grid.pack(fill="x", padx=10, pady=5)
        itc_vals = {}
        for i, f in enumerate(fields):
            ctk.CTkLabel(itc_grid, text=f"{f}:", font=ctk.CTkFont(size=11, weight="bold")).grid(
                row=0, column=i, padx=8, pady=2, sticky="w")
            lbl = ctk.CTkLabel(itc_grid, text="", font=ctk.CTkFont(size=11))
            lbl.grid(row=1, column=i, padx=8, pady=2, sticky="w")
            itc_vals[f] = lbl

        # 2A panel
        twoa_frame = ctk.CTkFrame(debug_win)
        twoa_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(twoa_frame, text="2A Row", font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="#1565C0").pack(anchor="w", padx=10, pady=(5, 0))
        twoa_grid = ctk.CTkFrame(twoa_frame, fg_color="transparent")
        twoa_grid.pack(fill="x", padx=10, pady=5)
        twoa_vals = {}
        for i, f in enumerate(fields):
            ctk.CTkLabel(twoa_grid, text=f"{f}:", font=ctk.CTkFont(size=11, weight="bold")).grid(
                row=0, column=i, padx=8, pady=2, sticky="w")
            lbl = ctk.CTkLabel(twoa_grid, text="", font=ctk.CTkFont(size=11))
            lbl.grid(row=1, column=i, padx=8, pady=2, sticky="w")
            twoa_vals[f] = lbl

        # Remarks
        remarks_frame = ctk.CTkFrame(debug_win, fg_color="transparent")
        remarks_frame.pack(fill="x", padx=10, pady=(5, 0))

        ctk.CTkLabel(remarks_frame, text="Remarks:",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(10, 5))

        remarks_entry = ctk.CTkEntry(remarks_frame, placeholder_text="Enter remarks for this invoice...",
                                     font=ctk.CTkFont(size=12), height=32)
        remarks_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        # Buttons
        btn_frame = ctk.CTkFrame(debug_win, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)

        match_btn = ctk.CTkButton(btn_frame, text="Match", fg_color="#2E7D32",
                                  hover_color="#1B5E20", font=ctk.CTkFont(size=14, weight="bold"),
                                  width=200, height=40, command=lambda: on_match())
        match_btn.pack(side="left", padx=20)

        skip_btn = ctk.CTkButton(btn_frame, text="Skip", fg_color="#757575",
                                 hover_color="#616161", font=ctk.CTkFont(size=14, weight="bold"),
                                 width=200, height=40, command=lambda: on_skip())
        skip_btn.pack(side="left", padx=20)

        counter_lbl = ctk.CTkLabel(debug_win, text="Matched: 0 | Skipped: 0",
                                   font=ctk.CTkFont(size=11))
        counter_lbl.pack(pady=5)

        def _save_remarks(p, action):
            """Save the remarks text to all sibling ITC rows for this invoice."""
            remark_text = remarks_entry.get().strip()
            if not remark_text:
                return
            # Prepend action to remark
            full_remark = f"[{action}] {remark_text}"
            norm_key = normalize_gstin(p['itc_gstin']) + '|' + p['itc_inv_norm']
            if 'Remarks' not in self.itc_result_df.columns:
                self.itc_result_df['Remarks'] = ''
            if vendor_gstn_col and vendor_inv_col:
                for row_idx, row in self.itc_result_df.iterrows():
                    row_key = normalize_gstin(str(row[vendor_gstn_col])) + '|' + normalize_invoice(str(row[vendor_inv_col]))
                    if row_key == norm_key:
                        self.itc_result_df.at[row_idx, 'Remarks'] = full_remark
            else:
                self.itc_result_df.at[p['itc_index'], 'Remarks'] = full_remark

        def display_pair(i):
            if i >= state['total']:
                show_summary()
                return
            p = candidates[i]
            progress_lbl.configure(text=f"Pair {i + 1} of {state['total']}")
            sim_lbl.configure(text=f"Invoice Similarity: {p['similarity']:.0%}")

            itc_vals['Invoice No'].configure(text=p['itc_invoice'])
            itc_vals['Invoice Date'].configure(text=p.get('itc_date', ''))
            itc_vals['GSTIN'].configure(text=p['itc_gstin'])
            itc_vals['Taxable Value'].configure(text=f"{p['itc_taxable']:,.2f}")
            itc_vals['IGST'].configure(text=f"{p['itc_igst']:,.2f}")
            itc_vals['CGST'].configure(text=f"{p['itc_cgst']:,.2f}")
            itc_vals['SGST'].configure(text=f"{p['itc_sgst']:,.2f}")

            twoa_vals['Invoice No'].configure(text=p['twoa_invoice'])
            twoa_vals['Invoice Date'].configure(text=p.get('twoa_date', ''))
            twoa_vals['GSTIN'].configure(text=p['twoa_gstin'])
            twoa_vals['Taxable Value'].configure(text=f"{p['twoa_taxable']:,.2f}")
            twoa_vals['IGST'].configure(text=f"{p['twoa_igst']:,.2f}")
            twoa_vals['CGST'].configure(text=f"{p['twoa_cgst']:,.2f}")
            twoa_vals['SGST'].configure(text=f"{p['twoa_sgst']:,.2f}")

            # Clear remarks for new pair
            remarks_entry.delete(0, 'end')

            counter_lbl.configure(text=f"Matched: {state['matched']} | Skipped: {state['skipped']}")

        def on_match():
            p = candidates[state['idx']]
            # Save remarks before moving on
            _save_remarks(p, 'Matched')
            # Update this row and all sibling rows with same GSTIN+Invoice
            norm_key = normalize_gstin(p['itc_gstin']) + '|' + p['itc_inv_norm']
            if vendor_gstn_col and vendor_inv_col:
                for row_idx, row in self.itc_result_df.iterrows():
                    row_key = normalize_gstin(str(row[vendor_gstn_col])) + '|' + normalize_invoice(str(row[vendor_inv_col]))
                    if row_key == norm_key:
                        self.itc_result_df.at[row_idx, 'Status'] = 'Matched but invoice number is not accurate'
            else:
                self.itc_result_df.at[p['itc_index'], 'Status'] = 'Matched but invoice number is not accurate'

            state['matched'] += 1
            state['idx'] += 1
            self.log(f"Debug Match: '{p['itc_invoice']}' -> '{p['twoa_invoice']}' (similarity: {p['similarity']:.0%})")
            display_pair(state['idx'])

        def on_skip():
            p = candidates[state['idx']]
            # Save remarks before moving on
            _save_remarks(p, 'Skipped')
            state['skipped'] += 1
            state['idx'] += 1
            display_pair(state['idx'])

        def show_summary():
            match_btn.configure(state="disabled")
            skip_btn.configure(state="disabled")
            remarks_entry.configure(state="disabled")
            progress_lbl.configure(text="Review Complete!")
            sim_lbl.configure(text="")
            for lbl in itc_vals.values():
                lbl.configure(text="")
            for lbl in twoa_vals.values():
                lbl.configure(text="")
            counter_lbl.configure(text=f"DONE - Matched: {state['matched']} | Skipped: {state['skipped']}")
            self.log(f"Debug Matching complete: {state['matched']} matched, {state['skipped']} skipped out of {state['total']}")
            messagebox.showinfo("Debug Matching Complete",
                                f"Total pairs reviewed: {state['total']}\n"
                                f"Matched by user: {state['matched']}\n"
                                f"Skipped: {state['skipped']}", parent=debug_win)

        display_pair(0)

    def view_results(self):
        """Open results viewer window"""
        has_itc = self.itc_result_df is not None and not self.itc_result_df.empty
        if not has_itc:
            messagebox.showwarning("Warning", "No data available to view!")
            return

        # Create results window
        results_window = ctk.CTkToplevel(self)
        results_window.title("ITC Results")
        results_window.geometry("1000x600")
        results_window.lift()
        results_window.focus_force()
        results_window.after(100, lambda: results_window.attributes('-topmost', False))
        results_window.attributes('-topmost', True)

        # Notebook for tabs
        tabview = ctk.CTkTabview(results_window)
        tabview.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab 1: ITC Results with Status
        tab1 = tabview.add("ITC Results")
        self.create_data_table(tab1, self.itc_result_df, status_field='Status',
                               status_values={'Matched': '#c8e6c9', 'Unmatched': '#ffcdd2',
                                              'Higher in 2A': '#fff9c4', 'Lower in 2A': '#bbdefb',
                                              'Not found in 2A': '#e1bee7',
                                              'Matched but invoice number is not accurate': '#ffe0b2'})

        # Tab 2: Unmatched in ITC (includes Unmatched, Higher in 2A, Lower in 2A)
        unmatched_itc = self.itc_result_df[self.itc_result_df['Status'].isin(['Unmatched', 'Higher in 2A', 'Lower in 2A', 'Not found in 2A'])]
        tab2 = tabview.add("Unmatched in ITC")
        if not unmatched_itc.empty:
            self.create_data_table(tab2, unmatched_itc, status_field='Status',
                                   status_values={'Unmatched': '#ffcdd2', 'Higher in 2A': '#fff9c4', 'Lower in 2A': '#bbdefb', 'Not found in 2A': '#e1bee7'})
        else:
            ctk.CTkLabel(tab2, text="No unmatched ITC items found!", font=ctk.CTkFont(size=16)).pack(pady=50)

        # Tab 3: Unmatched in 2A
        tab3 = tabview.add("Unmatched in 2A")
        if self.unmatched_2a_df is not None and not self.unmatched_2a_df.empty:
            self.create_data_table(tab3, self.unmatched_2a_df)
        else:
            ctk.CTkLabel(tab3, text="No unmatched GSTR 2A items found!", font=ctk.CTkFont(size=16)).pack(pady=50)

    def create_data_table(self, parent, df, status_field=None, status_values=None):
        """Create a data table with treeview.
        status_field: column name to use for row coloring (e.g. 'Status')
        status_values: dict mapping status text to background color (e.g. {'Matched': '#c8e6c9'})
        """
        frame = ctk.CTkFrame(parent)
        frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Summary
        total = len(df)
        if status_field and status_values and status_field in df.columns:
            summary_parts = [f"Total: {total}"]
            for status_name in status_values:
                count = len(df[df[status_field] == status_name])
                summary_parts.append(f"{status_name}: {count}")
            summary_text = " | ".join(summary_parts)
        else:
            matched = len(df[df['Status'] == 'MATCH']) if 'Status' in df.columns else 0
            mismatched = total - matched
            summary_text = f"Total: {total} | Matched: {matched} | Mismatched: {mismatched}"

        summary_label = ctk.CTkLabel(
            frame,
            text=summary_text,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        summary_label.pack(pady=5)

        # Create treeview with scrollbars
        tree_frame = ctk.CTkFrame(frame)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Scrollbars
        y_scroll = ttk.Scrollbar(tree_frame)
        y_scroll.pack(side="right", fill="y")

        x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal")
        x_scroll.pack(side="bottom", fill="x")

        # Treeview
        columns = list(df.columns)
        tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            yscrollcommand=y_scroll.set,
            xscrollcommand=x_scroll.set
        )

        y_scroll.config(command=tree.yview)
        x_scroll.config(command=tree.xview)

        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, minwidth=50)

        # Highlight GSTIN and Invoice columns if overall matched ratio >= 90%
        highlight_note = None
        if 'Status' in df.columns and total > 0:
            matched = len(df[df['Status'] == 'MATCH']) if 'MATCH' in df['Status'].values else len(df[df['Status'] == 'Matched'])
            # Accept both 'MATCH' and 'Matched' text variants
            if matched == 0:
                # try case-insensitive count
                matched = len(df[df['Status'].str.lower() == 'match']) if df['Status'].dtype == 'object' else matched
            matched_ratio = matched / total
            if matched_ratio >= 0.9:
                # Find likely GSTIN and Invoice columns
                gstin_cols = [c for c in columns if 'gstin' in c.lower() or 'gstn' in c.lower() or ("vendor" in c.lower() and 'gstn' in c.lower())]
                inv_cols = [c for c in columns if 'invoice' in c.lower() or 'inv' in c.lower() or 'external doc' in c.lower() or 'document' in c.lower()]
                for c in gstin_cols + inv_cols:
                    # append a visual marker to heading text
                    tree.heading(c, text=f"{c} ✅90%")
                highlight_note = 'Note: GSTIN & Invoice columns are 90%+ matched (highlighted)'

        # Configure tag colors for status-based coloring
        if status_field and status_values:
            for status_name, color in status_values.items():
                tag_name = status_name.lower().replace(' ', '_')
                tree.tag_configure(tag_name, background=color)

        # Add data (limit to 1000 rows for performance)
        display_df = df.head(1000)
        for idx, row in display_df.iterrows():
            values = [str(v) for v in row.values]
            if status_field and status_values and status_field in df.columns:
                status_val = str(row.get(status_field, ''))
                tag_name = status_val.lower().replace(' ', '_') if status_val in status_values else ''
                tree.insert('', 'end', values=values, tags=(tag_name,))
            else:
                tag = 'mismatch' if row.get('Status', '') != 'MATCH' else ''
                tree.insert('', 'end', values=values, tags=(tag,))

        # Show highlight note if applicable
        if highlight_note:
            ctk.CTkLabel(
                frame,
                text=highlight_note,
                font=ctk.CTkFont(size=11, weight='bold'),
                text_color='green'
            ).pack(pady=3)

        # Configure default mismatch tag if not using custom status
        if not status_field:
            tree.tag_configure('mismatch', background='#ffcccc')

        tree.pack(fill="both", expand=True)

        if len(df) > 1000:
            ctk.CTkLabel(
                frame,
                text=f"Showing first 1000 of {len(df)} rows. Download Excel for full data.",
                font=ctk.CTkFont(size=11),
                text_color="orange"
            ).pack(pady=5)


if __name__ == "__main__":
    app = GSTReconciliationApp()
    app.mainloop()
