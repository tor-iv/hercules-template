"""
Carta Cap Table Transformer
Transforms Carta exports into firm's internal cap table format.
Can be called standalone or via xlwings from Excel.
"""

import pandas as pd
from openpyxl import load_workbook
from pathlib import Path
from datetime import datetime
import re


def parse_carta_export(carta_path: str) -> dict:
    """Parse Carta export and extract relevant data."""
    
    # Read raw to find header row (look for 'Stakeholder ID' or 'Name')
    raw = pd.read_excel(carta_path, sheet_name='Detailed Cap', header=None)
    
    header_row = None
    for i in range(min(10, len(raw))):
        row_values = [str(v).lower() for v in raw.iloc[i].tolist()]
        if 'name' in row_values or 'stakeholder' in ' '.join(row_values):
            header_row = i
            break
    
    if header_row is None:
        raise ValueError("Could not find header row in Carta export")
    
    # Re-read with correct header
    df = pd.read_excel(carta_path, sheet_name='Detailed Cap', header=header_row)
    
    # Parse company name from title (row 0)
    title_cell = str(raw.iloc[0, 0])
    company_name = title_cell.replace('Detailed Capitalization Table', '').strip()
    
    # Parse date from "As of" row (usually row 1)
    cap_table_date = None
    for i in range(min(5, len(raw))):
        cell_val = str(raw.iloc[i, 0])
        if 'As of' in cell_val:
            date_match = re.search(r'(\d{2}/\d{2}/\d{4})', cell_val)
            if date_match:
                cap_table_date = datetime.strptime(date_match.group(1), '%m/%d/%Y')
                break
    
    # Identify share class columns (exclude ID, Name, and summary columns)
    share_class_cols = []
    for col in df.columns:
        col_lower = col.lower()
        if any(x in col_lower for x in ['class', 'series']) and 'units' in col_lower:
            share_class_cols.append(col)
    
    # Identify options columns
    options_cols = [col for col in df.columns if 'option' in col.lower() or 'rsu' in col.lower()]
    
    # Filter to stakeholder rows only (exclude summary rows at bottom)
    # Summary rows have specific text patterns
    summary_patterns = ['total', 'outstanding', 'available', 'fully diluted', 'percentage', 'price per']
    
    def is_stakeholder_row(row):
        name = str(row.get('Name', '')).lower()
        return name and not any(pattern in name for pattern in summary_patterns) and name != 'nan'
    
    stakeholders_df = df[df.apply(is_stakeholder_row, axis=1)].copy()
    
    # Calculate total shares per stakeholder
    stakeholders_df['_total_shares'] = 0
    for col in share_class_cols:
        stakeholders_df['_total_shares'] += pd.to_numeric(stakeholders_df[col], errors='coerce').fillna(0)
    
    # Sum options columns
    stakeholders_df['_total_options'] = 0
    for col in options_cols:
        stakeholders_df['_total_options'] += pd.to_numeric(stakeholders_df[col], errors='coerce').fillna(0)
    
    # Extract validation totals from summary rows
    validation = {}
    for idx, row in df.iterrows():
        name = str(row.get('Name', '')).lower()
        if 'total units outstanding' in name:
            validation['total_outstanding'] = row.get('Outstanding Units', 0)
        elif 'fully diluted' in name and 'units' in name:
            for col in share_class_cols:
                validation[f'fully_diluted_{col}'] = row.get(col, 0)
    
    return {
        'company_name': company_name,
        'cap_table_date': cap_table_date or datetime.now(),
        'stakeholders': stakeholders_df,
        'share_class_cols': share_class_cols,
        'options_cols': options_cols,
        'validation': validation
    }


def transform_to_template(carta_data: dict, template_path: str, output_path: str) -> dict:
    """Transform Carta data into the firm's cap table template."""
    
    # Load template (preserves formulas)
    wb = load_workbook(template_path)
    inputs_sheet = wb['Inputs']
    
    # 1. Set company name and date
    inputs_sheet['I6'] = carta_data['company_name']
    inputs_sheet['I7'] = carta_data['cap_table_date']
    
    # 2. Update share class headers in row 30 (columns F-I for up to 4 classes)
    share_classes = carta_data['share_class_cols']
    class_col_mapping = {}  # Maps class name to column letter
    
    for i, class_name in enumerate(share_classes[:4]):  # Max 4 classes
        col_letter = chr(ord('F') + i)  # F, G, H, I
        col_num = 6 + i  # Column numbers: 6, 7, 8, 9
        
        # Clean up class name for header (e.g., "Class A Units (CA)" -> "Class A")
        clean_name = re.sub(r'\s*Units.*$', '', class_name).strip()
        clean_name = re.sub(r'\s*\([^)]*\)', '', clean_name).strip()
        
        inputs_sheet.cell(row=30, column=col_num).value = clean_name
        class_col_mapping[class_name] = col_num
    
    # 3. Sort stakeholders by total shares (descending) for top 9
    stakeholders = carta_data['stakeholders'].copy()
    stakeholders = stakeholders.sort_values('_total_shares', ascending=False)
    
    # Split into top 9 and "other"
    top_investors = stakeholders.head(9)
    other_investors = stakeholders.iloc[9:] if len(stakeholders) > 9 else pd.DataFrame()
    
    # 4. Populate investor rows 31-39
    for i, (idx, row) in enumerate(top_investors.iterrows()):
        excel_row = 31 + i
        
        # Investor name (column D)
        inputs_sheet.cell(row=excel_row, column=4).value = row['Name']
        
        # Share classes (columns F-I based on mapping)
        for class_name, col_num in class_col_mapping.items():
            value = pd.to_numeric(row.get(class_name, 0), errors='coerce')
            if pd.notna(value) and value != 0:
                inputs_sheet.cell(row=excel_row, column=col_num).value = int(value)
            else:
                inputs_sheet.cell(row=excel_row, column=col_num).value = 0
        
        # Common shares (column P) - Carta doesn't separate, so 0
        inputs_sheet.cell(row=excel_row, column=16).value = 0
        
        # Options (column Q) - sum of all options columns
        options_total = row['_total_options']
        inputs_sheet.cell(row=excel_row, column=17).value = int(options_total) if pd.notna(options_total) else 0
    
    # Clear remaining investor rows (if less than 9 investors)
    for i in range(len(top_investors), 9):
        excel_row = 31 + i
        inputs_sheet.cell(row=excel_row, column=4).value = None
        for col in range(6, 18):  # F through Q
            inputs_sheet.cell(row=excel_row, column=col).value = 0
    
    # 5. Populate "Other Investors" row (row 40)
    if len(other_investors) > 0:
        inputs_sheet.cell(row=40, column=4).value = "Other Investors"
        
        for class_name, col_num in class_col_mapping.items():
            total = pd.to_numeric(other_investors[class_name], errors='coerce').fillna(0).sum()
            inputs_sheet.cell(row=40, column=col_num).value = int(total) if total > 0 else 0
        
        inputs_sheet.cell(row=40, column=16).value = 0  # Common
        inputs_sheet.cell(row=40, column=17).value = int(other_investors['_total_options'].sum())  # Options
    else:
        inputs_sheet.cell(row=40, column=4).value = "Other Investors"
        for col in range(6, 18):
            inputs_sheet.cell(row=40, column=col).value = 0
    
    # 6. Management row (row 41) - sum all Common shares
    # In Carta exports, individual names typically represent management/employees
    # For now, we'll leave Management row for manual adjustment
    # The template has this at row 41 or uses a different structure
    
    # 7. Save output
    wb.save(output_path)
    
    # 8. Generate validation summary
    total_by_class = {}
    for class_name in share_classes[:4]:
        col_total = pd.to_numeric(stakeholders[class_name], errors='coerce').fillna(0).sum()
        total_by_class[class_name] = col_total
    
    return {
        'output_path': output_path,
        'investors_processed': len(stakeholders),
        'top_investors': len(top_investors),
        'other_investors_count': len(other_investors),
        'share_classes_mapped': list(class_col_mapping.keys()),
        'totals_by_class': total_by_class,
        'validation': carta_data['validation']
    }


def run_transformation(carta_path: str, template_path: str, output_dir: str = None) -> dict:
    """Main entry point for the transformation."""
    
    carta_path = Path(carta_path)
    template_path = Path(template_path)
    
    if not carta_path.exists():
        raise FileNotFoundError(f"Carta export not found: {carta_path}")
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")
    
    # Parse Carta data
    carta_data = parse_carta_export(str(carta_path))
    
    # Generate output filename
    company_clean = re.sub(r'[^\w\s-]', '', carta_data['company_name']).strip()
    date_str = carta_data['cap_table_date'].strftime('%Y%m%d')
    output_filename = f"{company_clean}_Cap_Table_{date_str}.xlsx"
    
    if output_dir:
        output_path = Path(output_dir) / output_filename
    else:
        output_path = carta_path.parent / output_filename
    
    # Run transformation
    result = transform_to_template(carta_data, str(template_path), str(output_path))
    
    return result


# xlwings entry point
def main():
    """Called by xlwings button click."""
    import xlwings as xw
    
    # Get the calling workbook
    wb = xw.Book.caller()
    
    # File picker for Carta export
    import tkinter as tk
    from tkinter import filedialog
    
    root = tk.Tk()
    root.withdraw()
    
    carta_path = filedialog.askopenfilename(
        title="Select Carta Export",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    
    if not carta_path:
        xw.Book.caller().sheets[0].range('A1').value = "Cancelled - no file selected"
        return
    
    # Template path (in templates/ folder relative to project root)
    script_dir = Path(__file__).parent
    template_path = script_dir.parent / "templates" / "Cap Table Template.xlsx"
    
    if not template_path.exists():
        # Try looking in common locations
        template_path = filedialog.askopenfilename(
            title="Select Cap Table Template",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not template_path:
            return
    
    try:
        result = run_transformation(carta_path, str(template_path))
        
        # Show success message
        msg = f"Success!\n\nOutput: {result['output_path']}\n"
        msg += f"Investors processed: {result['investors_processed']}\n"
        msg += f"Top investors: {result['top_investors']}\n"
        msg += f"Other investors: {result['other_investors_count']}"
        
        tk.messagebox.showinfo("Carta Transformer", msg)
        
        # Open the output file
        xw.Book(result['output_path'])
        
    except Exception as e:
        tk.messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    # For testing without xlwings
    import sys
    if len(sys.argv) >= 3:
        result = run_transformation(sys.argv[1], sys.argv[2])
        print(f"Output: {result['output_path']}")
        print(f"Investors: {result['investors_processed']}")
    else:
        print("Usage: python carta_to_cap_table.py <carta_export.xlsx> <template.xlsx>")
