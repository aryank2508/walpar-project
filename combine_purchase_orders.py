"""
Production-Safe Purchase Order Excel File Combiner

This script combines multiple Excel files containing Purchase Order (PO) data
from yearly folders into single combined Excel files per year.

IMPORTANT: All output files will have ONLY the specified columns in the exact order provided.
Missing columns in a year will be filled with empty values.

Author: Senior Python Data Engineer
Date: 2025
"""

import pandas as pd
import os
import re
from typing import Dict, List, Optional
import warnings

# Suppress warnings for cleaner output
warnings.filterwarnings('ignore', category=UserWarning)

# Configuration
BASE_DIR = "PO"
OUTPUT_DIR = "combined-data"

# File patterns to skip (temporary files, hidden files, etc.)
SKIP_PATTERNS = ['~$', '.ini', '.lnk', 'desktop', 'Autosaved', 'AutoRecovered', 'Recovered', 'INDEX', '.tmp']

# EXACT COLUMN LIST - This is the only columns that will be in output files
# Order must be maintained exactly as specified
# Updated to match actual Excel structure after user changes
REQUIRED_COLUMNS = [
    'Folder',
    'Product Type (NEW)',
    'Source_File',
    'Sheet_Name',
    'Production Order Form',
    'Reference Format No',
    'PO Reference',
    'Reference Order No',
    'Order Form',
    'Order Type',
    'Order Date',
    'O.F. Date',
    'PO Date',
    'Generic Name',
    'Brand Name',
    'Quantity',
    'Pack Size',
    'Product Rate (NEW)',
    'Packing Size',
    'Packing Type',
    'Strength',
    'Net Content',
    'M.R.P.',
    'M.R.P. (Per Strip)',
    'Exp. Date',
    'Shelf Life',
    'Flavour',
    'Colour',
    'Packaging',
    'Bottle Type',
    'Measuring Cap',
    'Spray Pump',
    'ROPP Cap',
    'Closure',
    'Cap Color',
    'Label Provided By',
    'Carton Provided By',
    'Dropper',
    'PVC Size',
    'PVC Colour',
    'PVC Color & Type',
    'Blister Type',
    'Capsules Size',
    'Capsules Color',
    'Capsule Type',
    'Punch Size',
    'Tablet Color',
    'Coating Type',
    'Jar',
    'Jar Type',
    'Jar Lid Color',
    'Tin Type',
    'Tin / Jar Type',
    'Tin / Jar Rim Type',
    'Corrugated Box',
    'Corrugated Box Pack Size',
    'Corrugated Tin',
    'Cap / Lid / Spoon',
    'Spoon',
    'Leaflet',
    'Dropper (if not counted earlier)',
    'Strapping',
    'Wrapping',
    'Shrinking',
    'Dispatched Detail',
    'Date Of Dispatch',
    'Company Name',
    'Name Of Client',
    'Name Of Marketed By',
    'Contact Person',
    'Contact Number',
    'Contact Detail Of Dispatch',
    'Specific Requirement If Any',
    'Artwork Revision',
    'Artwork Revision No',
    'Prepared By',
    'Checked By',
    'Approved Date',
    'Issued By Sign & Date',
    'Checked & Authorised Sign & Date',
    'Authorised Sign',
    'Designation',
    'Rate',
    'Due Payment',
    'Remark',
    'Product Batch Number',
]


def is_valid_excel_file(file_path: str) -> bool:
    """Check if a file is a valid Excel file that should be processed."""
    file_name = os.path.basename(file_path).lower()
    
    for pattern in SKIP_PATTERNS:
        if pattern.lower() in file_name:
            return False
    
    return file_path.endswith('.xlsx')


def get_all_excel_files(folder_path: str, recursive: bool = True) -> List[str]:
    """
    Get all valid Excel files from a folder.
    
    Args:
        folder_path: Path to the folder to search
        recursive: If True, search subdirectories as well (default: True)
    
    Returns:
        List of Excel file paths
    """
    excel_files = []
    
    if not os.path.exists(folder_path):
        return excel_files
    
    if recursive:
        # Recursive search - walk through all subdirectories
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                if is_valid_excel_file(file_path):
                    excel_files.append(file_path)
    else:
        # Non-recursive - only immediate folder
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path) and is_valid_excel_file(item_path):
                excel_files.append(item_path)
    
    return sorted(excel_files)


def normalize_text(text: str) -> str:
    """Normalize text for matching (lowercase, remove extra spaces/punctuation)."""
    if not text or pd.isna(text):
        return ''
    text = str(text).strip().lower()
    # Remove extra spaces, punctuation variations
    text = re.sub(r'[_\s\.\-\(\)]+', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def match_column_name(source_name: str, target_columns: List[str]) -> Optional[str]:
    """Match a source column name to one of the target columns."""
    if not source_name:
        return None
    
    source_normalized = normalize_text(source_name)
    
    # Exact match first
    for target in target_columns:
        if normalize_text(target) == source_normalized:
            return target
    
    # Fuzzy matching - check if all words match
    source_words = set(source_normalized.split())
    
    best_match = None
    best_score = 0
    
    for target in target_columns:
        target_normalized = normalize_text(target)
        target_words = set(target_normalized.split())
        
        # Check if source words are subset of target or vice versa
        if source_words == target_words:
            return target
        elif source_words.issubset(target_words) or target_words.issubset(source_words):
            score = len(source_words & target_words)
            if score > best_score:
                best_score = score
                best_match = target
    
    # Check for common variations
    variations = {
        'reference format no': 'Reference Format No',
        'ref format no': 'Reference Format No',
        'reference format': 'Reference Format No',
        'po reference': 'PO Reference',
        'po ref': 'PO Reference',
        'reference order no': 'Reference Order No',
        'order form number': 'Order Form',
        'order form no': 'Order Form',
        'order form': 'Order Form',
        'order type': 'Order Type',
        'order date': 'Order Date',
        'of date': 'O.F. Date',
        'o.f. date': 'O.F. Date',
        'po date': 'PO Date',
        'generic name': 'Generic Name',
        'brand name': 'Brand Name',
        'quantity': 'Quantity',
        'qty': 'Quantity',
        'qty uom': 'Qty (UOM)',
        'qty (uom)': 'Qty (UOM)',
        'pack size': 'Pack Size',
        'packing size': 'Packing Size',
        'packing type': 'Packing Type',
        'strength': 'Strength',
        'net content': 'Net Content',
        'm.r.p': 'M.R.P.',
        'mrp': 'M.R.P.',
        'm.r.p.': 'M.R.P.',
        'm.r.p. per strip': 'M.R.P. (Per Strip)',
        'exp date': 'Exp. Date',
        'expiry date': 'Exp. Date',
        'shelf life': 'Shelf Life',
        'flavour': 'Flavour',
        'flavor': 'Flavour',
        'colour': 'Colour',
        'color': 'Colour',
        'packaging': 'Packaging',
        'bottle type': 'Bottle Type',
        'measuring cap': 'Measuring Cap',
        'spray pump': 'Spray Pump',
        'ropp cap': 'ROPP Cap',
        'closure': 'Closure',
        'cap color': 'Cap Color',
        'cap colour': 'Cap Color',
        'label provided by': 'Label Provided By',
        'carton provided by': 'Carton Provided By',
        'dropper': 'Dropper',
        'pvc size': 'PVC Size',
        'pvc colour': 'PVC Colour',
        'pvc color': 'PVC Colour',
        'pvc color type': 'PVC Color & Type',
        'pvc color & type': 'PVC Color & Type',
        'blister type': 'Blister Type',
        'capsules size': 'Capsules Size',
        'capsule size': 'Capsules Size',
        'capsules color': 'Capsules Color',
        'capsule color': 'Capsules Color',
        'capsules colour': 'Capsules Color',
        'capsule type': 'Capsule Type',
        'tablet color': 'Tablet Color',
        'tablet colour': 'Tablet Color',
        'coating type': 'Coating Type',
        'punch size': 'Punch Size',
        'jar': 'Jar',
        'jar type': 'Jar Type',
        'tin type': 'Tin Type',
        'tin jar type': 'Tin / Jar Type',
        'tin / jar type': 'Tin / Jar Type',
        'tin jar rim type': 'Tin / Jar Rim Type',
        'tin / jar rim type': 'Tin / Jar Rim Type',
        'jar lid color': 'Jar Lid Color',
        'corrugated box': 'Corrugated Box',
        'corrugated box pack size': 'Corrugated Box Pack Size',
        'corrugated tin': 'Corrugated Tin',
        'cap lid spoon': 'Cap / Lid / Spoon',
        'cap / lid / spoon': 'Cap / Lid / Spoon',
        'spoon': 'Spoon',
        'leaflet': 'Leaflet',
        'strapping': 'Strapping',
        'wrapping': 'Wrapping',
        'shrinking': 'Shrinking',
        'dispatched detail': 'Dispatched Detail',
        'company name': 'Company Name',
        'name of client': 'Name Of Client',
        'name of marketed by': 'Name Of Marketed By',
        'marketed by': 'Name Of Marketed By',
        'contact person': 'Contact Person',
        'contact number': 'Contact Number',
        'contact detail of dispatch': 'Contact Detail Of Dispatch',
        'date of dispatch': 'Date Of Dispatch',
        'specific requirement if any': 'Specific Requirement If Any',
        'artwork revision': 'Artwork Revision',
        'artwork revision no': 'Artwork Revision No',
        'prepared by': 'Prepared By',
        'checked by': 'Checked By',
        'approved date': 'Approved Date',
        'issued by sign date': 'Issued By Sign & Date',
        'checked authorised sign date': 'Checked & Authorised Sign & Date',
        'authorised sign': 'Authorised Sign',
        'designation': 'Designation',
        'rate': 'Rate',
        'due payment': 'Due Payment',
        'remark': 'Remark',
        'product batch number': 'Product Batch Number',
    }
    
    if source_normalized in variations:
        return variations[source_normalized]
    
    return best_match


def safe_extract(df, row_index: int, col_index: int, pattern: str) -> Optional[str]:
    """Safely extract data using regex pattern from a specific cell."""
    try:
        if row_index < len(df) and col_index < len(df.columns):
            cell_value = df.iloc[row_index, col_index]
            if pd.notna(cell_value):
                match = re.search(pattern, str(cell_value), re.IGNORECASE)
                if match:
                    return match.group(1).strip()
    except:
        pass
    return None


def extract_data_from_sheet(df_sheet: pd.DataFrame, file_name: str, sheet_name: str) -> Optional[Dict]:
    """Extract Purchase Order data from a single sheet."""
    if df_sheet.empty or len(df_sheet) < 5:
        return None
    
    data = {}
    
    # Extract header information
    try:
        for row_idx in range(3, min(10, len(df_sheet))):
            if len(df_sheet.columns) == 0:
                continue
            row_val = str(df_sheet.iloc[row_idx, 0])
            
            # Reference Format No
            if 'reference format' in row_val.lower():
                ref_match = safe_extract(df_sheet, row_idx, 0, r'Reference format no\.\*rev no\.:\s*(.*)')
                if ref_match:
                    data['Reference Format No'] = ref_match
            
            # Order Form / Order Form Number
            if 'order form number' in row_val.lower() or 'order form' in row_val.lower():
                order_form_match = safe_extract(df_sheet, row_idx, 0, r'Order form number\s*:\s*(\w+)')
                if order_form_match:
                    data['Order Form'] = order_form_match
            
            # O.F. Date
            if 'o.f. date' in row_val.lower() or 'of date' in row_val.lower():
                of_date_match = safe_extract(df_sheet, row_idx, 0, r'O\.F\. Date\s*:\s*([\d\-/]+)')
                if of_date_match:
                    data['O.F. Date'] = of_date_match
            
            # PO Date
            if 'po date' in row_val.lower():
                po_date_match = safe_extract(df_sheet, row_idx, 0, r'PO Date\s*:\s*([\d\-/]+)')
                if po_date_match:
                    data['PO Date'] = po_date_match
            
            # PO Reference
            if 'poreference' in row_val.lower() or 'po reference' in row_val.lower():
                po_ref_match = safe_extract(df_sheet, row_idx, 0, r'POreference\s*:\s*(.*)')
                if po_ref_match:
                    data['PO Reference'] = po_ref_match
            
            # Order Type
            if 'order type' in row_val.lower():
                order_type_match = safe_extract(df_sheet, row_idx, 0, r'Order type\s*:\s*(.*)')
                if order_type_match:
                    data['Order Type'] = order_type_match
            
            # Contact Person
            if 'contact person' in row_val.lower():
                contact_match = safe_extract(df_sheet, row_idx, 0, r'Contact person\s*\([^)]*\)\s*:\s*(.*)')
                if contact_match:
                    data['Contact Person'] = contact_match
    except:
        pass
    
    # Extract Title-Details pairs
    try:
        for row_idx in range(8, len(df_sheet)):
            try:
                if len(df_sheet.columns) < 3:
                    continue
                
                title_cell = df_sheet.iloc[row_idx, 1]
                
                if pd.notna(title_cell):
                    title = str(title_cell).strip()
                    
                    if title and not title.isdigit() and title.lower() not in ['no.', 'no', 'title', 'details']:
                        details_cell = df_sheet.iloc[row_idx, 2] if len(df_sheet.columns) > 2 else None
                        details_value = ''
                        
                        if pd.notna(details_cell):
                            details_value = str(details_cell).strip()
                        
                        # Check column 4 for additional info (like MRP)
                        additional_value = None
                        if len(df_sheet.columns) > 4:
                            additional_cell = df_sheet.iloc[row_idx, 4]
                            if pd.notna(additional_cell):
                                additional_value = str(additional_cell).strip()
                        
                        # Match title to required column
                        matched_column = match_column_name(title, REQUIRED_COLUMNS)
                        
                        if matched_column and matched_column not in data:
                            if details_value and details_value.upper() not in ['NA', 'N/A', '']:
                                # Handle MRP - sometimes value is in column 4
                                if 'm.r.p' in matched_column.lower():
                                    if additional_value:
                                        mrp_clean = re.sub(r'^Rs\.?\s*', '', additional_value, flags=re.IGNORECASE).strip()
                                        mrp_clean = re.sub(r'[,\s]', '', mrp_clean)
                                        if mrp_clean and mrp_clean.upper() not in ['EXPORT', 'NA', 'PHY.SAMPLE', 'TO BE CONFIRM', 'PERUNITPRICE', '']:
                                            data[matched_column] = mrp_clean
                                        else:
                                            data[matched_column] = details_value
                                    else:
                                        data[matched_column] = details_value
                                # Handle Quantity - extract numeric part
                                elif 'quantity' in matched_column.lower() or 'qty' in matched_column.lower():
                                    qty_match = re.match(r'[\"\']?([\d,]+)', details_value)
                                    if qty_match:
                                        qty_clean = qty_match.group(1).replace(',', '').strip()
                                        data[matched_column] = qty_clean
                                    else:
                                        data[matched_column] = details_value
                                else:
                                    data[matched_column] = details_value
            except:
                continue
    except:
        pass
    
    # Check if we have minimum required data
    if 'Order Form' not in data and 'Reference Format No' not in data and 'PO Reference' not in data:
        return None
    
    return data


def get_processed_files_from_tracking(year_name: str, tracking_dir: str = None) -> set:
    """
    Get set of source files that have already been processed.
    Uses a tracking file to remember processed files.
    """
    processed_files = set()
    
    if tracking_dir is None:
        tracking_dir = OUTPUT_DIR
    
    # Create tracking file path
    tracking_file = os.path.join(tracking_dir, f".{year_name}_processed.txt")
    
    if os.path.exists(tracking_file):
        try:
            with open(tracking_file, 'r', encoding='utf-8') as f:
                processed_files = {line.strip() for line in f if line.strip()}
        except Exception:
            pass
    
    return processed_files


def mark_file_as_processed(file_name: str, year_name: str, tracking_dir: str = None):
    """Mark a file as processed in the tracking file."""
    if tracking_dir is None:
        tracking_dir = OUTPUT_DIR
    
    os.makedirs(tracking_dir, exist_ok=True)
    tracking_file = os.path.join(tracking_dir, f".{year_name}_processed.txt")
    
    try:
        with open(tracking_file, 'a', encoding='utf-8') as f:
            f.write(f"{file_name}\n")
    except Exception:
        pass


def process_year_folder(year_folder: str, output_dir: str, skip_processed: bool = True) -> bool:
    """
    Process all Excel files in a year folder and create combined output.
    
    Args:
        year_folder: Path to the year folder
        output_dir: Output directory path
        skip_processed: If True, skip files already in output (default: True)
    """
    year_name = os.path.basename(year_folder)
    
    # Convert year name to output format
    year_parts = year_name.split('-')
    if len(year_parts) == 2:
        output_year_name = f"{year_parts[0]}-{year_parts[1][-2:]}"
    else:
        output_year_name = year_name
    
    output_file = os.path.join(output_dir, f"{output_year_name}.xlsx")
    
    print(f"\n{'='*70}")
    print(f"Processing {year_name} folder")
    print(f"{'='*70}\n")
    
    # Check for already processed files using tracking file
    processed_files = set()
    
    if skip_processed:
        processed_files = get_processed_files_from_tracking(output_year_name, output_dir)
        if processed_files:
            print(f"Found {len(processed_files)} already processed file(s). Will skip them.")
    
    excel_files = get_all_excel_files(year_folder)
    
    if not excel_files:
        print(f"⚠️ No Excel files found in {year_name} folder")
        return False
    
    print(f"Found {len(excel_files)} Excel file(s)")
    
    all_records = []
    total_sheets_processed = 0
    skipped_files = 0
    
    for file_idx, file_path in enumerate(excel_files, 1):
        file_name = os.path.basename(file_path)
        
        # Check if file should be skipped
        if skip_processed and file_name in processed_files:
            print(f"\n[{file_idx}/{len(excel_files)}] >> Skipping (already processed): {file_name}", flush=True)
            skipped_files += 1
            continue
        
        print(f"\n[{file_idx}/{len(excel_files)}] Processing: {file_name}", flush=True)
        
        try:
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
            print(f"    Opening file... {len(sheet_names)} sheets found", flush=True)
            
            sheets_with_data = 0
            for sheet_idx, sheet_name in enumerate(sheet_names, 1):
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    extracted = extract_data_from_sheet(df, file_name, sheet_name)
                    
                    if extracted:
                        all_records.append(extracted)
                        total_sheets_processed += 1
                        sheets_with_data += 1
                    
                    if len(sheet_names) > 50 and sheet_idx % 50 == 0:
                        print(f"    Progress: {sheet_idx}/{len(sheet_names)} sheets...", flush=True)
                
                except Exception as e:
                    if sheet_idx <= 3:
                        print(f"    ! Error in sheet '{sheet_name}': {e}", flush=True)
                    continue
            
            if sheets_with_data > 0:
                print(f"    + Extracted data from {sheets_with_data} sheet(s)", flush=True)
                # Mark file as processed if we successfully extracted data
                if skip_processed:
                    mark_file_as_processed(file_name, output_year_name, output_dir)
        
        except Exception as e:
            print(f"    X Error processing file: {e}", flush=True)
            continue
    
    if not all_records:
        print(f"\nX No data extracted from {year_name} folder")
        return False
    
    print(f"\nCombining {len(all_records)} records...")
    
    # Convert to DataFrame
    df = pd.DataFrame(all_records)
    
    # Create DataFrame with ONLY required columns in exact order
    # Initialize with all required columns
    result_data = {col: [] for col in REQUIRED_COLUMNS}
    
    # Map extracted data to required columns
    for record in all_records:
        for req_col in REQUIRED_COLUMNS:
            # Find matching value in record
            value = None
            for key, val in record.items():
                if match_column_name(key, [req_col]) == req_col:
                    value = val
                    break
            
            result_data[req_col].append(value)
    
    # Create final DataFrame with exact columns in exact order
    df_final = pd.DataFrame(result_data)
    
    # Remove duplicates
    print("Removing duplicates...")
    original_count = len(df_final)
    
    def create_key(row):
        parts = []
        if pd.notna(row.get('Order Form')):
            parts.append(str(row['Order Form']))
        if pd.notna(row.get('PO Date')):
            parts.append(str(row['PO Date']))
        if pd.notna(row.get('Brand Name')):
            parts.append(str(row['Brand Name']))
        return '|'.join(parts) if parts else f"ROW_{id(row)}"
    
    duplicate_keys = df_final.apply(create_key, axis=1)
    df_final = df_final.assign(_duplicate_key=duplicate_keys).copy()
    df_dedup = df_final.drop_duplicates(subset=['_duplicate_key'], keep='first')
    df_dedup = df_dedup.drop(columns=['_duplicate_key']).copy()
    
    duplicates_removed = original_count - len(df_dedup)
    if duplicates_removed > 0:
        print(f"  Removed {duplicates_removed} duplicate record(s)")
    
    # Sort by date if available
    if 'PO Date' in df_dedup.columns:
        try:
            sort_dates = pd.to_datetime(df_dedup['PO Date'], errors='coerce')
            df_dedup = df_dedup.assign(_sort_date=sort_dates).sort_values(by='_sort_date', na_position='last')
            df_dedup = df_dedup.drop(columns=['_sort_date']).copy()
        except:
            pass
    
    # Ensure columns are in exact order
    df_dedup = df_dedup[REQUIRED_COLUMNS]
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Save to Excel
    output_file = os.path.join(output_dir, f"{output_year_name}.xlsx")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_dedup.to_excel(writer, sheet_name='Combined_PO_Data', index=False)
    
    print(f"\n+ Successfully created: {output_file}")
    print(f"   Total records: {len(df_dedup)}")
    print(f"   Total columns: {len(df_dedup.columns)}")
    
    return True


def main():
    """Main execution function."""
    print("\n" + "="*70)
    print("Purchase Order Excel File Combiner")
    print("All output files will have ONLY the specified columns in exact order")
    print("="*70)
    
    # Check base directory
    if not os.path.exists(BASE_DIR):
        print(f"\n❌ Error: Base directory '{BASE_DIR}' not found!")
        return
    
    # Find all year folders
    year_folders = []
    for item in os.listdir(BASE_DIR):
        item_path = os.path.join(BASE_DIR, item)
        if os.path.isdir(item_path):
            if '-' in item and len(item.split('-')) == 2:
                try:
                    parts = item.split('-')
                    int(parts[0])
                    int(parts[1])
                    year_folders.append(item_path)
                except ValueError:
                    continue
    
    year_folders = sorted(year_folders)
    
    if not year_folders:
        print(f"\n❌ No year folders found in '{BASE_DIR}'")
        return
    
    print(f"\nFound {len(year_folders)} year folder(s) to process")
    print(f"Required columns: {len(REQUIRED_COLUMNS)}")
    
    # Process each year folder
    successful = 0
    failed = 0
    
    for year_folder in year_folders:
        if process_year_folder(year_folder, OUTPUT_DIR):
            successful += 1
        else:
            failed += 1
    
    # Summary
    print(f"\n{'='*70}")
    print("Processing Complete")
    print(f"{'='*70}")
    print(f"+ Successfully processed: {successful} year(s)")
    if failed > 0:
        print(f"X Failed: {failed} year(s)")
    print(f"\nOutput files saved to: {os.path.abspath(OUTPUT_DIR)}")
    print(f"All output files have {len(REQUIRED_COLUMNS)} columns in the exact order specified")
    print("="*70 + "\n")


if __name__ == "__main__":
    main()
