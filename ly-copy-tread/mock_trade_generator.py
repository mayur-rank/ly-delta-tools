import xlwings as xw
import time
import random
import os

# --- CONFIGURATION ---
EXCEL_FILE = "LiveTrade.xlsx"
SHEET_NAME = "Sheet1"
SYMBOLS = ["NIFTY23MAR17500CE", "BANKNIFTY23MAR40000PE", "RELIANCE", "TCS", "INFY"]
SIDES = ["BUY", "SELL"]

def create_mock_excel():
    """Creates a new Excel file with headers if it doesn't exist."""
    if not os.path.exists(EXCEL_FILE):
        print(f"Creating {EXCEL_FILE}...")
        wb = xw.Book()
        sheet = wb.sheets[0]
        sheet.name = SHEET_NAME
        # Adding headers for clarity (Leader script assumes A=Sym, B=Side, C=Qty)
        sheet.range("A1").value = ["Symbol", "Side", "Qty"]
        wb.save(EXCEL_FILE)
        # wb.close() # Keep open for live updates? Leader needs it open.
        return wb
    else:
        print(f"Opening existing {EXCEL_FILE}...")
        return xw.Book(EXCEL_FILE)

def start_mock_generator():
    wb = create_mock_excel()
    sheet = wb.sheets[SHEET_NAME]
    
    print("Mock Generator Started. Press Ctrl+C to stop.")
    print("Adding a new trade every 2-5 seconds...")
    
    try:
        while True:
            # Find the next empty row
            current_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row + 1
            
            # Generate random trade data
            symbol = random.choice(SYMBOLS)
            side = random.choice(SIDES)
            qty = random.randint(1, 10) * 50 # Multiples of 50 for lot sizes
            
            # Write to Excel
            sheet.range(f"A{current_row}").value = [symbol, side, qty]
            print(f"Added Row {current_row}: {symbol}, {side}, {qty}")
            
            # Save so Leader script sees it
            wb.save()
            
            # Wait for random interval
            time.sleep(random.uniform(2, 5))
            
    except KeyboardInterrupt:
        print("\nMock Generator Stopped.")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    start_mock_generator()
