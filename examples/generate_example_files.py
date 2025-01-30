"""Generate example Excel files for testing."""
from pathlib import Path
import openpyxl
from datetime import datetime, timedelta
import random

def generate_example_files():
    """Generate source and target Excel files."""
    data_dir = Path(__file__).parent / "data"
    data_dir.mkdir(exist_ok=True)
    
    # Generate source file
    source_wb = openpyxl.Workbook()
    
    # CustomerData sheet
    ws = source_wb.active
    ws.title = "CustomerData"
    ws.append([
        "CustomerID", "FirstName", "LastName", "Email",
        "RegistrationDate", "LastLoginDate", "Status"
    ])
    
    # Generate sample customer data
    for i in range(1, 21):  # 20 customers
        reg_date = datetime(2023, 1, 1) + timedelta(days=random.randint(0, 365))
        last_login = reg_date + timedelta(days=random.randint(0, 30))
        ws.append([
            f"CUST{i:04d}",
            f"FirstName{i}",
            f"LastName{i}",
            f"customer{i}@example.com",
            reg_date.strftime("%Y-%m-%d"),
            last_login.strftime("%Y-%m-%d %H:%M:%S"),
            random.choice(["Active", "Inactive"])
        ])
    
    # Transactions sheet
    ws = source_wb.create_sheet("Transactions")
    ws.append([
        "TransactionID", "CustomerID", "Date", "Amount",
        "Type", "Status", "Notes"
    ])
    
    # Generate sample transaction data
    for i in range(1, 51):  # 50 transactions
        trans_date = datetime(2023, 1, 1) + timedelta(days=random.randint(0, 365))
        ws.append([
            f"TRX{i:04d}",
            f"CUST{random.randint(1, 20):04d}",
            trans_date.strftime("%Y-%m-%d %H:%M:%S"),
            round(random.uniform(10, 1000), 2),
            random.choice(["Purchase", "Refund", "Credit"]),
            random.choice(["Completed", "Pending", "Failed"]),
            f"Transaction note {i}"
        ])
    
    source_wb.save(data_dir / "source.xlsx")
    
    # Generate target file
    target_wb = openpyxl.Workbook()
    
    # CustomerSummary sheet
    ws = target_wb.active
    ws.title = "CustomerSummary"
    ws.append([
        "CustomerID", "FullName", "Email", "DaysSinceRegistration",
        "LastLoginDate", "IsActive", "TransactionCount", "TotalSpent"
    ])
    
    # TransactionSummary sheet
    ws = target_wb.create_sheet("TransactionSummary")
    ws.append([
        "CustomerID", "TransactionCount", "TotalAmount",
        "AverageAmount", "LastTransactionDate", "SuccessRate"
    ])
    
    target_wb.save(data_dir / "target.xlsx")
    
    print("✨ Example files generated successfully!")
    print(f"📊 Source file: {data_dir / 'source.xlsx'}")
    print(f"🎯 Target file: {data_dir / 'target.xlsx'}")

if __name__ == "__main__":
    generate_example_files()