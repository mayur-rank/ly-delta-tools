import pywinauto
from pywinauto import Desktop
import time
import os
import sys
import csv

# --- CONFIGURATION (Change these after tomorrow's first run!) ---
# If your ODIN columns are different, change these numbers based on the discovery log.
COL_STRIKE = 1
COL_EXPIRY = 0
COL_QTY = 6      # Net Qty column

# Set this to "Downloads" to test on your current PC.
# Set this to "[Integrated Net Position]" for the real ODIN.
TARGET_WINDOW_TITLE = "Downloads"  
# The exact name of the Table/List inside the window
TARGET_GRID_TITLE = "Items View" 

LOG_FILE = "odin_position_history.csv"
POLLING_INTERVAL = 0.5 # 500ms

def get_row_data(item):
    """
    Extracts all text columns from a row. 
    Different Windows controls handle this differently.
    """
    try:
        # Try getting all sub-elements (cells)
        cells = item.children()
        if cells:
            return [c.window_text() for c in cells]
        
        # Fallback: Just return the whole text split by tabs
        return item.window_text().split("\t")
    except:
        return []

def start_logging():
    print(f"--- ODIN DATA LOGGER & DELTA TRACKER ---")
    print(f"Targeting Window: '{TARGET_WINDOW_TITLE}'")
    print(f"Logging to: {os.path.abspath(LOG_FILE)}")
    
    # State tracking: { (strike, expiry): last_qty }
    position_state = {}

    # Initialize CSV with headers
    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Time", "Strike", "Expiry", "Action", "Qty_Change", "New_Net_Qty"])

    try:
        print("Waiting for window... (Press Ctrl+C to stop)")
        while True:
            try:
                # Use Desktop to find the window (MDI windows can be tricky otherwise)
                desktop = Desktop(backend="uia")
                win = desktop.window(title_re=f".*{TARGET_WINDOW_TITLE}.*")
                
                if win.exists():
                    grid = win.child_window(title=TARGET_GRID_TITLE, control_type="List")
                    
                    rows = grid.items()
                    
                    for row in rows:
                        data = get_row_data(row)
                        
                        # Basic validation: Skip empty rows
                        if len(data) < 2: continue
                        
                        try:
                            # 1. Extract Values
                            strike = data[COL_STRIKE]
                            expiry = data[COL_EXPIRY]
                            
                            # Handle empty or non-numeric Net Qty
                            raw_qty = data[COL_QTY].replace(",", "").strip()
                            net_qty = int(raw_qty) if raw_qty and raw_qty != "" else 0
                            
                            pos_key = (strike, expiry)
                            
                            # 2. Check for Changes (Deltas)
                            if pos_key in position_state:
                                old_qty = position_state[pos_key]
                                if net_qty != old_qty:
                                    delta = net_qty - old_qty
                                    action = "BUY" if delta > 0 else "SELL"
                                    
                                    msg = f"[{time.strftime('%H:%M:%S')}] {action} {abs(delta)} | {strike} | Net: {net_qty}"
                                    print(msg)
                                    
                                    # Log to CSV
                                    with open(LOG_FILE, "a", newline="") as f:
                                        writer = csv.writer(f)
                                        writer.writerow([time.strftime('%H:%M:%S'), strike, expiry, action, abs(delta), net_qty])
                                    
                                    # Update state
                                    position_state[pos_key] = net_qty
                            else:
                                # First time seeing this position
                                position_state[pos_key] = net_qty
                                print(f"Registered Position: {strike} (Net Qty: {net_qty})")
                                
                        except Exception as parse_err:
                            # Usually happens on header rows or empty lines
                            # Uncomment below to debug column indices:
                            # print(f"Parse Error on row: {data}")
                            continue
                            
                else:
                    # Heartbeat
                    print(".", end="", flush=True)
                    
            except Exception as loop_error:
                # This could be window closing or element disappearing
                time.sleep(1)
                
            time.sleep(POLLING_INTERVAL)

    except KeyboardInterrupt:
        print("\nExiting and saving final state...")

if __name__ == "__main__":
    start_logging()
