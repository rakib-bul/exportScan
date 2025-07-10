import pandas as pd
import numpy as np
import os
from tkinter import Tk, messagebox
import tkinter as tk
import logging
from typing import Tuple, Optional
from constants import REQUIRED_COLS_DF1, REQUIRED_COLS_DF2, STATUS_COLUMN


def clean_column_name(col: str) -> str:
    return str(col).strip().lower().replace(' ', '').replace('_', '').replace('-', '')

def get_excel_engine(file_path: str) -> Optional[str]:
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.xlsx':
        return 'openpyxl'
    elif ext == '.xls':
        return 'xlrd'
    return None

def validate_file(file_path: str) -> Tuple[bool, str]:
    if not os.path.exists(file_path):
        return False, f"File does not exist: {file_path}"
    if not os.access(file_path, os.R_OK):
        return False, f"File is not readable: {file_path}"
    return True, ""

def compare_excel_files(file1_path: str, file2_path: str, status_var, result_text) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    try:
        logging.info(f"Starting comparison between {file1_path} and {file2_path}")
        
        # Validate files
        for path in [file1_path, file2_path]:
            is_valid, error = validate_file(path)
            if not is_valid:
                return None, error

        # Get appropriate engines
        engine1 = get_excel_engine(file1_path)
        engine2 = get_excel_engine(file2_path)
        
        # Read Excel files
        df1 = pd.read_excel(file1_path, engine=engine1)
        df2 = pd.read_excel(file2_path, engine=engine2)
        
        # Clean column names
        df1.columns = [clean_column_name(col) for col in df1.columns]
        df2.columns = [clean_column_name(col) for col in df2.columns]
        
        # Check required columns
        required_cols = {'df1': REQUIRED_COLS_DF1, 'df2': REQUIRED_COLS_DF2}
        for df_name, cols in required_cols.items():
            df = df1 if df_name == 'df1' else df2
            missing = [col for col in cols if col not in df.columns]
            if missing:
                error_msg = f"Missing columns in {df_name}: {', '.join(missing)}"
                logging.error(error_msg)
                return None, error_msg
        
        # Initialize status column and counters
        df2[STATUS_COLUMN] = "Not Checked"
        match_counts = {
            'po_match': 0, 'job_po_match': 0, 'style_color_match': 0,
            'no_match': 0, 'qty_mismatch': 0, 'less_shipment': 0,
            'over_shipment': 0, 'no_shipment': 0
        }
        
        # Create last4 columns for jobno
        df1['jobno_last4'] = df1['jobno'].astype(str).str.strip().str[-4:]
        df2['jobno_last4'] = df2['jobno'].astype(str).str.strip().str[-4:]
        
        # Matching logic
        # FIRST PRIORITY MATCHING: PO Number only
        result_text.insert(tk.END, "\n1. Matching by PO Number...\n")
        result_text.update()
        po_mask = (
            df2['ponumber'].notna() &
            df2['ponumber'].isin(df1['ponumber'])
        )

        # Group df1 by PO number to sum exfactoryqty for POs with multiple entries
        po_grouped = df1.groupby('ponumber')['exfactoryqty'].sum().reset_index()

        for idx in df2[po_mask].index:
            if df2.at[idx, STATUS_COLUMN] != "Not Checked":
                continue
                    
            po_val = df2.at[idx, 'ponumber']
            # Get the total exfactoryqty for this PO
            total_exfactory = po_grouped[po_grouped['ponumber'] == po_val]['exfactoryqty'].values
            
            if len(total_exfactory) > 0:
                exfactory_qty = total_exfactory[0]
                ship_qty = df2.at[idx, 'shipqty']
                        
                # Check if export quantity is empty
                if pd.isna(exfactory_qty) or exfactory_qty == '':
                    df2.at[idx, STATUS_COLUMN] = 'No Shipment (PO Match)'
                    match_counts['no_shipment'] += 1
                elif exfactory_qty == ship_qty:
                    df2.at[idx, STATUS_COLUMN] = 'Ok (PO Match)'
                    match_counts['po_match'] += 1
                elif exfactory_qty < ship_qty:
                    df2.at[idx, STATUS_COLUMN] = f'Over Shipment (PO Match: {exfactory_qty} vs {ship_qty})'
                    match_counts['over_shipment'] += 1
                else:
                    df2.at[idx, STATUS_COLUMN] = f'Less Shipment (PO Match: {exfactory_qty} vs {ship_qty})'
                    match_counts['less_shipment'] += 1
        
        # SECOND PRIORITY MATCHING: Job No (last4) + PO Number
        result_text.insert(tk.END, "2. Matching by Job No (last4) + PO Number...\n")
        result_text.update()
        job_po_mask = (
            (df2[STATUS_COLUMN] == "Not Checked") &
            df2['jobno_last4'].notna() & 
            df2['ponumber'].notna()
        )

        # Create a grouped version of df1 for Job+PO matching
        df1['job_po_key'] = df1['jobno_last4'] + '_' + df1['ponumber'].astype(str)
        job_po_grouped = df1.groupby('job_po_key')['exfactoryqty'].sum().reset_index()

        for idx in df2[job_po_mask].index:
            job_val = df2.at[idx, 'jobno_last4']
            po_val = df2.at[idx, 'ponumber']
            job_po_key = f"{job_val}_{po_val}"
            
            # Get the total exfactoryqty for this Job+PO combination
            total_exfactory = job_po_grouped[job_po_grouped['job_po_key'] == job_po_key]['exfactoryqty'].values
            
            if len(total_exfactory) > 0:
                exfactory_qty = total_exfactory[0]
                ship_qty = df2.at[idx, 'shipqty']
                        
                # Check if export quantity is empty
                if pd.isna(exfactory_qty) or exfactory_qty == '':
                    df2.at[idx, STATUS_COLUMN] = 'No Shipment (Job+PO Match)'
                    match_counts['no_shipment'] += 1
                elif exfactory_qty == ship_qty:
                    df2.at[idx, STATUS_COLUMN] = 'Ok (Job+PO Match)'
                    match_counts['job_po_match'] += 1
                elif exfactory_qty < ship_qty:
                    df2.at[idx, STATUS_COLUMN] = f'Over Shipment (Job+PO: {exfactory_qty} vs {ship_qty})'
                    match_counts['over_shipment'] += 1
                else:
                    df2.at[idx, STATUS_COLUMN] = f'Less Shipment (Job+PO: {exfactory_qty} vs {ship_qty})'
                    match_counts['less_shipment'] += 1
        
        # THIRD PRIORITY MATCHING: Style Ref + Color
        result_text.insert(tk.END, "3. Matching by Style Ref + Color...\n")
        result_text.update()
        style_color_mask = (df2[STATUS_COLUMN] == "Not Checked")
        secondary_df = df2[style_color_mask].copy()
        
        for idx in secondary_df.index:
            style_val = secondary_df.at[idx, 'stylerefno']
            color_val = secondary_df.at[idx, 'color']
            
            match = df1[
                (df1['stylerefno'] == style_val) &
                (df1['color'] == color_val)
            ]
            
            if not match.empty:
                exfactory_qty = match.iloc[0]['exfactoryqty']
                ship_qty = secondary_df.at[idx, 'shipqty']
                
                # Check if export quantity is empty
                if pd.isna(exfactory_qty) or exfactory_qty == '':
                    df2.at[idx, STATUS_COLUMN] = 'No Shipment (Style+Color Match)'
                    match_counts['no_shipment'] += 1
                elif exfactory_qty == ship_qty:
                    df2.at[idx, STATUS_COLUMN] = 'Ok (Style+Color Match)'
                    match_counts['style_color_match'] += 1
                elif exfactory_qty < ship_qty:
                    df2.at[idx, STATUS_COLUMN] = f'Over Shipment (Style+Color: {exfactory_qty} vs {ship_qty})'
                    match_counts['over_shipment'] += 1
                else:
                    df2.at[idx, STATUS_COLUMN] = f'Less Shipment (Style+Color: {exfactory_qty} vs {ship_qty})'
                    match_counts['less_shipment'] += 1
            else:
                df2.at[idx, STATUS_COLUMN] = 'No Match Found'
                match_counts['no_match'] += 1
        
        # Drop temporary columns before returning
        df2 = df2.drop(columns=['jobno_last4'])
        if 'job_po_key' in df1.columns:
            df1 = df1.drop(columns=['job_po_key'])
        
        # Generate summary report
        summary = (
            f"\n=== Matching Summary ===\n"
            f"Perfect Matches: {match_counts['po_match'] + match_counts['job_po_match'] + match_counts['style_color_match']}\n"
            f"Less Shipment Cases: {match_counts['less_shipment']}\n"
            f"Over Shipment Cases: {match_counts['over_shipment']}\n"
            f"No Shipment Cases: {match_counts['no_shipment']}\n"
            f"No Matches Found: {match_counts['no_match']}\n"
            f"Total Records Processed: {len(df2)}\n"
        )
        result_text.insert('end', summary)
        result_text.update()
        
        logging.info("Comparison completed successfully")
        return df2, None
    
    except ImportError as e:
        error_msg = "Reading .xls files requires xlrd. Install with: pip install xlrd" if 'xlrd' in str(e).lower() else f"Import Error: {str(e)}"
        logging.error(error_msg)
        return None, error_msg
    except Exception as e:
        error_msg = f"Error during processing: {str(e)}"
        logging.error(error_msg, exc_info=True)
        return None, error_msg
    

def show_summary_stats(df: pd.DataFrame) -> dict:
    return {
        'Total Records': len(df),
        'Perfect Matches': len(df[df[STATUS_COLUMN].str.startswith('Ok')]),
        'Quantity Mismatches': len(df[df[STATUS_COLUMN].str.contains('Qty Mismatch')]),
        'No Matches': len(df[df[STATUS_COLUMN] == 'No Match Found'])
    }