import pywinauto
from pywinauto import Desktop
import sys

def run_diagnostic():
    print("--- ODIN COLUMN DIAGNOSTIC ---")
    print("Searching for the '[Integrated Net Position]' window...")
    
    try:
        # We use UIA as it is better for detailed grid data
        desktop = Desktop(backend="uia")
        
        # ODIN titles often have specific bracket formatting
        # We try to find the window by title keyword
        win = None
        for w in desktop.windows():
            if "Integrated Net Position" in w.window_text():
                win = w
                break
        
        if win:
            print(f"SUCCESS! Found Window: '{win.window_text()}'")
            
            # Find the first 'List' or 'DataGrid'
            # Most ODIN windows use 'List' for their main tables
            grid = win.child_window(control_type="List")
            
            if not grid.exists():
                print("Could not find a 'List' control. Trying 'DataGrid'...")
                grid = win.child_window(control_type="DataGrid")
            
            items = grid.items()
            if not items:
                print("Found the table, but it appears to be EMPTY.")
                return

            print(f"Found {len(items)} items in the table.")
            print("\n" + "="*30)
            print("CALIBRATION DATA (First 2 Rows):")
            print("="*30)
            
            for i, item in enumerate(items[:2]):
                print(f"\n[ROW {i} RAW TEXT]: {item.window_text()}")
                
                # Check for sub-cells
                children = item.children()
                if children:
                    print(f"[ROW {i} COLUMNS]:")
                    for j, child in enumerate(children):
                        print(f"  Column {j}: '{child.window_text()}'")
                else:
                    print(f"[ROW {i}]: No sub-columns found. Item might be a single string.")
            
            print("\n" + "="*30)
            print("Please COPY AND PASTE the text above back to me.")
            
        else:
            print("\n[!] ERROR: Could not find window containing 'Integrated Net Position'.")
            print("Visible Windows:")
            for w in desktop.windows():
                if w.window_text():
                    print(f" - {w.window_text()}")
            
    except Exception as e:
        print(f"\n[!] CRITICAL ERROR: {e}")
        print("Tip: Ensure you are running this script as ADMINISTRATOR.")

if __name__ == "__main__":
    run_diagnostic()
