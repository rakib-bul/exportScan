import pandas as pd
import numpy as np
import os
import tkinter as tk
import logging
from typing import Tuple, Optional
from constants import REQUIRED_COLS_DF1, REQUIRED_COLS_DF2, STATUS_COLUMN, BUYER_SPECIFIC_BUYERS

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

def compare_excel_files(
    file1_path: str, 
    file2_path: str, 
    status_var, 
    result_text, 
    buyer_specific=False, 
    combine_po_in="df1"
) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    try:
        logging.info(f"Starting comparison with buyer_specific={buyer_specific}, combine_po_in={combine_po_in}")
        
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
        
        # Convert key columns to uppercase and strip
        text_cols = ['jobno', 'ponumber', 'stylerefno', 'color', 'buyer']
        for col in text_cols:
            if col in df1.columns:
                df1[col] = df1[col].astype(str).str.strip().str.upper()
            if col in df2.columns:
                df2[col] = df2[col].astype(str).str.strip().str.upper()
        
        # Handle buyer-specific option
        if buyer_specific:
            if 'buyer' not in df2.columns:
                result_text.insert(tk.END, "\nWARNING: Buyer column missing - ")
                result_text.insert(tk.END, "disabling buyer-specific matching\n")
                buyer_specific = False
            else:
                df2['buyer_normalized'] = df2['buyer'].astype(str).str.strip().str.upper()
        
        # Check required columns
        required_cols = {'df1': REQUIRED_COLS_DF1, 'df2': REQUIRED_COLS_DF2}
        for df_name, cols in required_cols.items():
            df = df1 if df_name == 'df1' else df2
            missing = [col for col in cols if col not in df.columns]
            if missing:
                error_msg = f"Missing columns in {df_name}: {', '.join(missing)}"
                logging.error(error_msg)
                return None, error_msg
        
        # Combine StyleRefNo and PO if requested
        if buyer_specific:
            if combine_po_in == "df1":
                df1['combined_po'] = df1['stylerefno'].astype(str) + '-' + df1['ponumber'].astype(str)
                result_text.insert(tk.END, "\nCreated combined PO-StyleRef in Source File\n")
            else:
                df2['combined_po'] = df2['stylerefno'].astype(str) + '-' + df2['ponumber'].astype(str)
                result_text.insert(tk.END, "\nCreated combined PO-StyleRef in Target File\n")
            result_text.update()
        
        # Function to get last4 of jobno
        def get_last4(jobno):
            s = str(jobno).strip()
            return s[-4:] if len(s) >= 4 else s
        
        # Create last4 columns for jobno
        df1['jobno_last4'] = df1['jobno'].apply(get_last4)
        df2['jobno_last4'] = df2['jobno'].apply(get_last4)
        
        # Initialize status column and counters
        df2[STATUS_COLUMN] = "Not Checked"
        match_counts = {
            'po_match': 0, 'job_po_match': 0, 'style_color_match': 0,
            'no_match': 0, 'qty_mismatch': 0, 'less_shipment': 0,
            'over_shipment': 0, 'no_shipment': 0,
            'buyer_po_job': 0, 'buyer_combined': 0
        }
        
        # Prepare buyer-specific masks
        if buyer_specific:
            buyer_mask = df2['buyer_normalized'].isin(BUYER_SPECIFIC_BUYERS)
            non_buyer_mask = ~buyer_mask
        else:
            buyer_mask = pd.Series(False, index=df2.index)
            non_buyer_mask = pd.Series(True, index=df2.index)
        
        # ===== 1. STANDARD MATCHING (for non-buyer-specific rows) =====
        result_text.insert(tk.END, "\n=== Standard Matching ===\n")
        result_text.update()
        
        # FIRST PRIORITY: PO Number only
        result_text.insert(tk.END, "1. Matching by PO Number...\n")
        result_text.update()
        po_mask = (
            non_buyer_mask & 
            (df2[STATUS_COLUMN] == "Not Checked") &
            df2['ponumber'].notna() &
            (df2['ponumber'] != '') &
            df2['ponumber'].isin(df1['ponumber'])
        )

        # Group df1 by PO number to sum quantities
        po_grouped = df1.groupby('ponumber')['exfactoryqty'].sum().reset_index()

        for idx in df2[po_mask].index:
            po_val = df2.at[idx, 'ponumber']
            total_exfactory = po_grouped.loc[po_grouped['ponumber'] == po_val, 'exfactoryqty'].values
            
            if len(total_exfactory) > 0:
                exfactory_qty = total_exfactory[0]
                ship_qty = df2.at[idx, 'shipqty']
                
                # Handle empty/zero quantities
                if pd.isna(exfactory_qty) or exfactory_qty in ['', 0, '0']:
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
            
            # Update progress every 100 rows
            if idx % 100 == 0:
                result_text.insert(tk.END, f"Processed {idx} rows...\n")
                result_text.see(tk.END)
                result_text.update_idletasks()
        
        # SECOND PRIORITY: Job No (last4) + PO Number
        result_text.insert(tk.END, "2. Matching by Job No (last4) + PO Number...\n")
        result_text.update()
        job_po_mask = (
            non_buyer_mask & 
            (df2[STATUS_COLUMN] == "Not Checked") &
            df2['jobno_last4'].notna() & 
            (df2['jobno_last4'] != '') &
            df2['ponumber'].notna() &
            (df2['ponumber'] != '')
        )

        # Create grouped version for Job+PO matching
        df1['job_po_key'] = df1['jobno_last4'] + '_' + df1['ponumber'].astype(str)
        job_po_grouped = df1.groupby('job_po_key')['exfactoryqty'].sum().reset_index()

        for idx in df2[job_po_mask].index:
            job_val = df2.at[idx, 'jobno_last4']
            po_val = df2.at[idx, 'ponumber']
            job_po_key = f"{job_val}_{po_val}"
            
            total_exfactory = job_po_grouped.loc[job_po_grouped['job_po_key'] == job_po_key, 'exfactoryqty'].values
            
            if len(total_exfactory) > 0:
                exfactory_qty = total_exfactory[0]
                ship_qty = df2.at[idx, 'shipqty']
                
                # Handle empty/zero quantities
                if pd.isna(exfactory_qty) or exfactory_qty in ['', 0, '0']:
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
            
            # Update progress every 100 rows
            if idx % 100 == 0:
                result_text.insert(tk.END, f"Processed {idx} rows...\n")
                result_text.see(tk.END)
                result_text.update_idletasks()
        
        # THIRD PRIORITY: Style Ref + Color
        result_text.insert(tk.END, "3. Matching by Style Ref + Color...\n")
        result_text.update()
        style_color_mask = (
            non_buyer_mask & 
            (df2[STATUS_COLUMN] == "Not Checked")
        )
        secondary_df = df2[style_color_mask].copy()
        
        for idx in secondary_df.index:
            style_val = secondary_df.at[idx, 'stylerefno']
            color_val = secondary_df.at[idx, 'color']
            
            style_match = df1[
                (df1['stylerefno'] == style_val) &
                (df1['color'] == color_val)
            ]
            
            if not style_match.empty:
                exfactory_qty = style_match['exfactoryqty'].sum()  # Sum all matches
                ship_qty = secondary_df.at[idx, 'shipqty']
                
                # Handle empty/zero quantities
                if pd.isna(exfactory_qty) or exfactory_qty in ['', 0, '0']:
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
            
            # Update progress every 100 rows
            if idx % 100 == 0:
                result_text.insert(tk.END, f"Processed {idx} rows...\n")
                result_text.see(tk.END)
                result_text.update_idletasks()
        
        # ===== 2. BUYER-SPECIFIC MATCHING =====
        if buyer_specific:
            result_text.insert(tk.END, "\n=== Buyer-Specific Matching ===\n")
            result_text.update()
            
            # FIRST: Match by PO + Job Number
            result_text.insert(tk.END, "1. Buyer Matching by PO + Job Number...\n")
            result_text.update()
            buyer_po_job_mask = (
                buyer_mask &
                (df2[STATUS_COLUMN] == "Not Checked") &
                df2['ponumber'].notna() &
                (df2['ponumber'] != '') &
                df2['jobno_last4'].notna() &
                (df2['jobno_last4'] != '')
            )
            
            # Create grouped version for PO+Job matching
            df1['po_job_key'] = df1['ponumber'].astype(str) + '_' + df1['jobno_last4']
            po_job_grouped = df1.groupby('po_job_key')['exfactoryqty'].sum().reset_index()
            
            for idx in df2[buyer_po_job_mask].index:
                po_val = df2.at[idx, 'ponumber']
                job_val = df2.at[idx, 'jobno_last4']
                po_job_key = f"{po_val}_{job_val}"
                
                total_exfactory = po_job_grouped.loc[po_job_grouped['po_job_key'] == po_job_key, 'exfactoryqty'].values
                
                if len(total_exfactory) > 0:
                    exfactory_qty = total_exfactory[0]
                    ship_qty = df2.at[idx, 'shipqty']
                    
                    # Handle empty/zero quantities
                    if pd.isna(exfactory_qty) or exfactory_qty in ['', 0, '0']:
                        df2.at[idx, STATUS_COLUMN] = 'No Shipment (Buyer PO+Job)'
                        match_counts['no_shipment'] += 1
                    elif exfactory_qty == ship_qty:
                        df2.at[idx, STATUS_COLUMN] = 'Ok (Buyer PO+Job)'
                        match_counts['buyer_po_job'] += 1
                    elif exfactory_qty < ship_qty:
                        df2.at[idx, STATUS_COLUMN] = f'Over Shipment (Buyer PO+Job: {exfactory_qty} vs {ship_qty})'
                        match_counts['over_shipment'] += 1
                    else:
                        df2.at[idx, STATUS_COLUMN] = f'Less Shipment (Buyer PO+Job: {exfactory_qty} vs {ship_qty})'
                        match_counts['less_shipment'] += 1
                
                # Update progress every 100 rows
                if idx % 100 == 0:
                    result_text.insert(tk.END, f"Processed {idx} rows...\n")
                    result_text.see(tk.END)
                    result_text.update_idletasks()
            
            # SECOND: Match by Combined PO
            result_text.insert(tk.END, "2. Buyer Matching by Combined PO-StyleRef...\n")
            result_text.update()
            buyer_combined_mask = (
                buyer_mask &
                (df2[STATUS_COLUMN] == "Not Checked")
            )
            
            for idx in df2[buyer_combined_mask].index:
                # Get combined PO value
                if combine_po_in == "df1":
                    key = df2.at[idx, 'stylerefno'] + '-' + df2.at[idx, 'ponumber']
                    match = df1[df1['combined_po'] == key]
                else:
                    key = df2.at[idx, 'combined_po']
                    match = df1[df1['combined_po'] == key]
                
                if not match.empty:
                    exfactory_qty = match['exfactoryqty'].sum()
                    ship_qty = df2.at[idx, 'shipqty']
                    
                    # Handle empty/zero quantities
                    if pd.isna(exfactory_qty) or exfactory_qty in ['', 0, '0']:
                        df2.at[idx, STATUS_COLUMN] = 'No Shipment (Buyer Combined)'
                        match_counts['no_shipment'] += 1
                    elif exfactory_qty == ship_qty:
                        df2.at[idx, STATUS_COLUMN] = 'Ok (Buyer Combined)'
                        match_counts['buyer_combined'] += 1
                    elif exfactory_qty < ship_qty:
                        df2.at[idx, STATUS_COLUMN] = f'Over Shipment (Buyer Combined: {exfactory_qty} vs {ship_qty})'
                        match_counts['over_shipment'] += 1
                    else:
                        df2.at[idx, STATUS_COLUMN] = f'Less Shipment (Buyer Combined: {exfactory_qty} vs {ship_qty})'
                        match_counts['less_shipment'] += 1
                
                # Update progress every 100 rows
                if idx % 100 == 0:
                    result_text.insert(tk.END, f"Processed {idx} rows...\n")
                    result_text.see(tk.END)
                    result_text.update_idletasks()
        
        # Handle unmatched buyer-specific rows
        if buyer_specific:
            unmatched_buyer_mask = (
                buyer_mask &
                (df2[STATUS_COLUMN] == "Not Checked")
            )
            df2.loc[unmatched_buyer_mask, STATUS_COLUMN] = 'No Match (Buyer)'
            match_counts['no_match'] += unmatched_buyer_mask.sum()

        # Handle remaining unmatched rows
        unmatched_mask = (df2[STATUS_COLUMN] == "Not Checked")
        df2.loc[unmatched_mask, STATUS_COLUMN] = 'No Match Found'
        match_counts['no_match'] += unmatched_mask.sum()

        # Drop temporary columns
        temp_cols = ['jobno_last4', 'buyer_normalized', 'job_po_key', 'po_job_key', 'combined_po']
        for col in temp_cols:
            if col in df2.columns:
                df2.drop(col, axis=1, inplace=True)
            if col in df1.columns:
                df1.drop(col, axis=1, inplace=True)

        # Generate summary report
        summary = (
            f"\n=== Matching Summary ===\n"
            f"Standard Perfect Matches: {match_counts['po_match'] + match_counts['job_po_match'] + match_counts['style_color_match']}\n"
            f"Buyer PO+Job Matches: {match_counts['buyer_po_job']}\n"
            f"Buyer Combined PO Matches: {match_counts['buyer_combined']}\n"
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
        'Quantity Mismatches': len(df[df[STATUS_COLUMN].str.contains('Shipment')]),
        'No Matches': len(df[df[STATUS_COLUMN] == 'No Match Found'])
    }