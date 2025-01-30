"""Verify the output Excel file contents."""
import openpyxl
from pathlib import Path

def verify_output():
    """Print the contents of the output Excel file."""
    data_dir = Path(__file__).parent / "data"
    output_file = data_dir / "test_output.xlsx"
    
    print("📊 Verifying output file contents...")
    print("=" * 60)
    
    wb = openpyxl.load_workbook(output_file, read_only=True)
    
    # Check CustomerSummary sheet
    if "CustomerSummary" in wb.sheetnames:
        ws = wb["CustomerSummary"]
        print("\n🧑 CustomerSummary Sheet:")
        print("-" * 40)
        
        # Print headers
        headers = [cell.value for cell in ws[1] if cell.value]
        print("Headers:", ", ".join(headers))
        
        # Print first few rows
        print("\nSample Data (first 3 rows):")
        for row in list(ws.rows)[1:4]:  # Skip header, take next 3 rows
            row_data = [str(cell.value) for cell in row if cell.value is not None]
            print(" | ".join(row_data))
    
    # Check TransactionSummary sheet
    if "TransactionSummary" in wb.sheetnames:
        ws = wb["TransactionSummary"]
        print("\n💰 TransactionSummary Sheet:")
        print("-" * 40)
        
        # Print headers
        headers = [cell.value for cell in ws[1] if cell.value]
        print("Headers:", ", ".join(headers))
        
        # Print first few rows
        print("\nSample Data (first 3 rows):")
        for row in list(ws.rows)[1:4]:  # Skip header, take next 3 rows
            row_data = [str(cell.value) for cell in row if cell.value is not None]
            print(" | ".join(row_data))
    
    wb.close()
    print("\n" + "=" * 60)

if __name__ == "__main__":
    verify_output()