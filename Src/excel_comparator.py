"""
Excel Comparator Tool

A comprehensive Python tool for comparing two Excel files cell-by-cell
and identifying all differences with detailed context information.

Author: Shreyes Shalgar
"""

import openpyxl
from openpyxl.utils import get_column_letter
import sys
import os
from pathlib import Path


class ExcelComparator:
    def __init__(self, file1_path, file2_path):
        """Initialize the comparator with two Excel files."""
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.differences = []
        
        try:
            self.workbook1 = openpyxl.load_workbook(file1_path, data_only=True)
            self.workbook2 = openpyxl.load_workbook(file2_path, data_only=True)
        except Exception as e:
            print(f"Error loading Excel files: {e}")
            sys.exit(1)
    
    def compare(self):
        """Compare both Excel files."""
        print(f"Comparing '{Path(self.file1_path).name}' and '{Path(self.file2_path).name}'...\n")
        
        sheets1 = self.workbook1.sheetnames
        sheets2 = self.workbook2.sheetnames
        
        # Check for sheet name differences
        if set(sheets1) != set(sheets2):
            missing_in_file2 = set(sheets1) - set(sheets2)
            missing_in_file1 = set(sheets2) - set(sheets1)
            
            if missing_in_file2:
                print(f"⚠ Sheets in File 1 but not in File 2: {', '.join(missing_in_file2)}\n")
            if missing_in_file1:
                print(f"⚠ Sheets in File 2 but not in File 1: {', '.join(missing_in_file1)}\n")
        
        # Compare common sheets
        common_sheets = set(sheets1) & set(sheets2)
        
        for sheet_name in sorted(common_sheets):
            self._compare_sheets(sheet_name)
        
        self._print_results()
    
    def _get_error_name(self, value):
        """Extract error name from cell value if it's an Excel error."""
        if value is None:
            return None
        # Check if value is an Excel error
        value_str = str(value)
        excel_errors = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!']
        for error in excel_errors:
            if value_str.strip() == error:
                return error
        return None
    
    def _compare_sheets(self, sheet_name):
        """Compare two sheets with the same name."""
        ws1 = self.workbook1[sheet_name]
        ws2 = self.workbook2[sheet_name]
        
        # Get the dimensions of both sheets
        max_row1 = ws1.max_row
        max_col1 = ws1.max_column
        max_row2 = ws2.max_row
        max_col2 = ws2.max_column
        
        max_row = max(max_row1, max_row2)
        max_col = max(max_col1, max_col2)
        
        # Compare cells
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell1 = ws1.cell(row, col)
                cell2 = ws2.cell(row, col)
                
                value1 = cell1.value
                value2 = cell2.value
                
                # Detect error names
                error_name1 = self._get_error_name(value1)
                error_name2 = self._get_error_name(value2)
                
                # Get column D value for this row from both files
                d_cell1 = ws1.cell(row, 4)  # Column D is the 4th column
                d_cell2 = ws2.cell(row, 4)
                d_value1 = d_cell1.value
                d_value2 = d_cell2.value
                
                # Compare values
                if value1 != value2:
                    col_letter = get_column_letter(col)
                    cell_address = f"{col_letter}{row}"
                    
                    # Get the header cell (row 1) from the same column where difference is found
                    header_cell1 = ws1.cell(1, col)
                    header_cell2 = ws2.cell(1, col)
                    header_value1 = header_cell1.value
                    header_value2 = header_cell2.value
                    
                    # Use actual error name if present, otherwise generic name
                    error_name = error_name1 or error_name2 or "Value Mismatch"
                    
                    self.differences.append({
                        'sheet': sheet_name,
                        'cell': cell_address,
                        'row': row,
                        'column': col,
                        'file1_value': value1,
                        'file2_value': value2,
                        'error_name': error_name,
                        'd_value1': d_value1,
                        'd_value2': d_value2,
                        'header_value1': header_value1,
                        'header_value2': header_value2
                    })
    
    def _print_results(self):
        """Print the comparison results."""
        if not self.differences:
            print("✓ No differences found! Both files are identical.\n")
            return
        
        print(f"Found {len(self.differences)} difference(s):\n")
        print("=" * 100)
        
        # Group differences by sheet
        by_sheet = {}
        for diff in self.differences:
            sheet = diff['sheet']
            if sheet not in by_sheet:
                by_sheet[sheet] = []
            by_sheet[sheet].append(diff)
        
        for sheet_name in sorted(by_sheet.keys()):
            diffs = by_sheet[sheet_name]
            print(f"\nSheet: '{sheet_name}' ({len(diffs)} difference(s))")
            print("-" * 100)
            
            for diff in sorted(diffs, key=lambda x: (x['row'], x['column'])):
                cell = diff['cell']
                val1 = diff['file1_value']
                val2 = diff['file2_value']
                
                print(f"  Cell: {cell}")
                print(f"    File 1: {repr(val1)}")
                print(f"    File 2: {repr(val2)}")
                print()
        
        print("=" * 100)
        print(f"\nTotal differences: {len(self.differences)}\n")
    
    def export_to_csv(self, output_file):
        """Export differences to a CSV file."""
        if not self.differences:
            print("No differences to export.")
            return
        
        import csv
        
        try:
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Sheet', 'Cell', 'Error_name_1', 'Error_name_2', 'Column', 'Column Header (File 1)', 'Column Header (File 2)', 'File 1 Value', 'File 2 Value'])
                
                for diff in sorted(self.differences, key=lambda x: (x['sheet'], x['row'], x['column'])):
                    col_letter = get_column_letter(diff['column'])
                    
                    writer.writerow([
                        diff['sheet'],
                        diff['cell'],
                        diff.get('d_value1'),
                        diff.get('d_value2'),
                        col_letter,
                        diff.get('header_value1'),
                        diff.get('header_value2'),
                        diff['file1_value'],
                        diff['file2_value']
                    ])
            
            print(f"Differences exported to '{output_file}'")
        except Exception as e:
            print(f"Error exporting to CSV: {e}")
    
    def export_to_excel(self, output_file):
        """Export differences to an Excel file."""
        if not self.differences:
            print("No differences to export.")
            return
        
        try:
            # Create a new workbook for the results
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Differences"
            
            # Write header row
            headers = ['Sheet', 'Cell', 'Error_name_1', 'Error_name_2', 'Column', 'Column Header (File 1)', 'Column Header (File 2)', 'File 1 Value', 'File 2 Value']
            ws.append(headers)
            
            # Style header row
            from openpyxl.styles import Font, PatternFill
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            
            # Write data rows
            for diff in sorted(self.differences, key=lambda x: (x['sheet'], x['row'], x['column'])):
                col_letter = get_column_letter(diff['column'])
                
                ws.append([
                    diff['sheet'],
                    diff['cell'],
                    diff.get('d_value1'),
                    diff.get('d_value2'),
                    col_letter,
                    diff.get('header_value1'),
                    diff.get('header_value2'),
                    diff['file1_value'],
                    diff['file2_value']
                ])
            
            # Adjust column widths
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 15
            ws.column_dimensions['D'].width = 15
            ws.column_dimensions['E'].width = 12
            ws.column_dimensions['F'].width = 20
            ws.column_dimensions['G'].width = 20
            ws.column_dimensions['H'].width = 20
            ws.column_dimensions['I'].width = 20
            
            # Save the workbook
            wb.save(output_file)
            print(f"Differences exported to '{output_file}'")
        except Exception as e:
            print(f"Error exporting to Excel: {e}")

def get_file_path(default_path, file_type):
    """Get file path from user input with validation."""
    while True:
        if default_path and os.path.exists(default_path):
            print(f"\033[92m✓ {file_type} File Found at:\033[0m {os.path.abspath(default_path)}")
            return os.path.abspath(default_path)
        
        file_path = input(f"\033[38;5;208mEnter path for {file_type} Excel file:\033[0m ")
        
        if not file_path.strip():
            print("\033[91m✗ Error: Path cannot be empty.\033[0m")
            continue
        
        file_path = os.path.abspath(file_path.strip())
        
        if not os.path.exists(file_path):
            print(f"\033[91m✗ Error: File not found at {file_path}\033[0m")
            continue
        
        if not file_path.lower().endswith(('.xlsx', '.xls', '.xlsm')):
            print("\\033[91m\u2717 Error: File must be an Excel file (.xlsx, .xls, or .xlsm)\\033[0m")
            continue
        
        print(f"\033[92m✓ File loaded: {file_path}\033[0m")
        return file_path
    
def main():
    """Main function."""
    os.system("title 'Excel Comparator Tool'")  # Set console title (Windows)
    print("\n" + "="*100)
    print("\033[96mExcel File Comparator Tool\033[0m")
    print("\033[90mAuthor: Shreyes Shalgar\033[0m")
    print("="*100 + "\n")
    
    # Get file paths from user input
    path1 = sys.argv[1] if len(sys.argv) > 1 else ""
    path2 = sys.argv[2] if len(sys.argv) > 2 else ""
    print(path1,path2)
    file1 = get_file_path(path1, "First")
    print()
    file2 = get_file_path(path2, "Second")
    
    if file1 == file2:
        print("\n\033[91m✗ Error: Cannot compare a file with itself!\033[0m")
        sys.exit(1)
    
    print()
    
    # Create comparator and run comparison
    try:
        comparator = ExcelComparator(file1, file2)
        comparator.compare()
        
        # Ask if user wants to export results
        if comparator.differences:
            export_choice = input("\nWould you like to export differences? (yes/no): ").strip().lower()
            if export_choice in ['yes', 'y']:
                format_choice = input("Export format - (1) Excel or (2) CSV? (1/2): ").strip()
                output_file = input("Enter output filename (default: differences): ").strip()
                
                if not output_file:
                    output_file = "differences"
                
                if format_choice == '2':
                    if not output_file.endswith('.csv'):
                        output_file += '.csv'
                    comparator.export_to_csv(output_file)
                else:
                    if not output_file.endswith('.xlsx'):
                        output_file += '.xlsx'
                    comparator.export_to_excel(output_file)
    except Exception as e:
        print(f"\033[91m✗ Error during comparison: {e}\033[0m")
        sys.exit(1)


if __name__ == "__main__":
    main()
