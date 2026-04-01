import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading
import time
import os
import csv
import pywinauto
from pywinauto import Desktop

# --- CALIBRATION (Will update after your diagnostic!) ---
COL_STRIKE = 2
COL_EXPIRY = 1
COL_QTY = 7
TARGET_WINDOW_TITLE = "Integrated Net Position"
TARGET_GRID_TITLE = "List" # Or "DataGrid"
LOG_FILE = "odin_trade_log.csv"

class OdinScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ODIN Trade Fetcher v1.0")
        self.root.geometry("600x450")
        self.root.configure(bg="#2c3e50")

        # Control Variables
        self.running = False
        self.stop_event = threading.Event()
        self.position_state = {}

        self.setup_ui()

    def setup_ui(self):
        # Header
        header = tk.Label(self.root, text="ODIN REAL-TIME FETCH PANEL", 
                         font=("Arial", 16, "bold"), bg="#2c3e50", fg="#ecf0f1", pady=20)
        header.pack()

        # Status Indicator
        self.status_frame = tk.Frame(self.root, bg="#2c3e50")
        self.status_frame.pack(pady=10)
        
        self.lbl_status = tk.Label(self.status_frame, text="STATUS: IDLE", 
                                  font=("Arial", 12, "bold"), bg="#2c3e50", fg="#bdc3c7")
        self.lbl_status.pack(side=tk.LEFT, padx=10)

        # Control Buttons
        btn_frame = tk.Frame(self.root, bg="#2c3e50")
        btn_frame.pack(pady=20)

        self.btn_start = tk.Button(btn_frame, text="▶ START FETCHING", command=self.start_scraping,
                                  font=("Arial", 10, "bold"), bg="#27ae60", fg="white", 
                                  width=18, height=2, relief=tk.FLAT)
        self.btn_start.pack(side=tk.LEFT, padx=10)

        self.btn_stop = tk.Button(btn_frame, text="■ STOP", command=self.stop_scraping,
                                 font=("Arial", 10, "bold"), bg="#c0392b", fg="white", 
                                 width=15, height=2, relief=tk.FLAT, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=10)

        self.btn_calibrate = tk.Button(btn_frame, text="🔍 CALIBRATE", command=self.run_calibration,
                                 font=("Arial", 10, "bold"), bg="#f39c12", fg="white", 
                                 width=15, height=2, relief=tk.FLAT)
        self.btn_calibrate.pack(side=tk.LEFT, padx=10)

        # Log Window
        log_label = tk.Label(self.root, text="Recent Activity:", font=("Arial", 10), 
                            bg="#2c3e50", fg="#bdc3c7")
        log_label.pack(anchor="w", padx=40)
        
        self.log_area = scrolledtext.ScrolledText(self.root, width=65, height=10, 
                                                 bg="#34495e", fg="#ecf0f1", font=("Consolas", 9))
        self.log_area.pack(pady=5, padx=40)

        # Initialize CSV
        if not os.path.exists(LOG_FILE):
            with open(LOG_FILE, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Time", "Strike", "Expiry", "Action", "Qty_Change", "New_Net_Qty"])

    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        formatted_msg = f"[{timestamp}] {message}\n"
        self.log_area.insert(tk.END, formatted_msg)
        self.log_area.see(tk.END)

    def update_status(self, text, color):
        self.lbl_status.config(text=f"STATUS: {text}", fg=color)

    def run_calibration(self):
        self.log("CALIBRATION STARTED...")
        CAL_FILE = "calibration_results.txt"
        
        try:
            desktop = Desktop(backend="uia")
            win = None
            for w in desktop.windows():
                if TARGET_WINDOW_TITLE in w.window_text():
                    win = w
                    break
            
            if win:
                grid = win.child_window(control_type="List")
                if not grid.exists():
                    grid = win.child_window(control_type="DataGrid")
                
                items = grid.items()
                if not items:
                    self.log("Table found but it is EMPTY.")
                    return

                with open(CAL_FILE, "w") as f:
                    f.write(f"--- CALIBRATION DATA ---\n")
                    f.write(f"Window: {win.window_text()}\n\n")
                    
                    for i, item in enumerate(items[:3]): # Log first 3 rows
                        f.write(f"ROW {i} TEXT: {item.window_text()}\n")
                        children = item.children()
                        for j, child in enumerate(children):
                            f.write(f"  Col {j}: '{child.window_text()}'\n")
                        f.write("-" * 30 + "\n")
                
                self.log(f"SUCCESS! Result saved to {CAL_FILE}")
                messagebox.showinfo("Calibration Done", f"Please send me the file:\n{os.path.abspath(CAL_FILE)}")
            else:
                self.log("ERROR: Could not find ODIN window!")
                messagebox.showerror("Error", "Could not find ODIN window. Make sure 'Integrated Net Position' is open.")
        except Exception as e:
            self.log(f"FAILED: {e}")

    def start_scraping(self):
        self.running = True
        self.stop_event.clear()
        self.btn_start.config(state=tk.DISABLED, bg="#7f8c8d")
        self.btn_stop.config(state=tk.NORMAL, bg="#e74c3c")
        self.update_status("RUNNING - SEARCHING...", "#2ecc71")
        self.log("Scraping engine started.")

        # Start background thread
        self.thread = threading.Thread(target=self.scraping_loop, daemon=True)
        self.thread.start()

    def stop_scraping(self):
        self.running = False
        self.stop_event.set()
        self.btn_start.config(state=tk.NORMAL, bg="#27ae60")
        self.btn_stop.config(state=tk.DISABLED, bg="#95a5a6")
        self.update_status("STOPPED", "#e74c3c")
        self.log("Scraping engine stopped.")

    def scraping_loop(self):
        while not self.stop_event.is_set():
            try:
                desktop = Desktop(backend="uia")
                # Flexible window finding
                win = None
                for w in desktop.windows():
                    if TARGET_WINDOW_TITLE in w.window_text():
                        win = w
                        break
                
                if win and win.exists():
                    self.update_status("RUNNING - CONNECTED", "#2ecc71")
                    
                    # Try to find the grid
                    grid = win.child_window(control_type="List") # Default for ODIN
                    if not grid.exists():
                        grid = win.child_window(control_type="DataGrid")

                    items = grid.items()
                    for row in items:
                        cells = row.children()
                        if len(cells) > max(COL_STRIKE, COL_QTY):
                            strike = cells[COL_STRIKE].window_text()
                            expiry = cells[COL_EXPIRY].window_text()
                            
                            # Clean and convert Qty
                            raw_qty = cells[COL_QTY].window_text().replace(",", "").strip()
                            net_qty = int(raw_qty) if raw_qty and raw_qty.lstrip('-').isdigit() else 0
                            
                            pos_key = (strike, expiry)
                            
                            if pos_key in self.position_state:
                                old_qty = self.position_state[pos_key]
                                if net_qty != old_qty:
                                    delta = net_qty - old_qty
                                    action = "BUY" if delta > 0 else "SELL"
                                    
                                    msg = f"{action} {abs(delta)} | {strike} | Net: {net_qty}"
                                    self.log(msg)
                                    
                                    with open(LOG_FILE, "a", newline="") as f:
                                        writer = csv.writer(f)
                                        writer.writerow([time.strftime('%H:%M:%S'), strike, expiry, action, abs(delta), net_qty])
                                    
                                    self.position_state[pos_key] = net_qty
                            else:
                                if strike: # Register non-empty rows
                                    self.position_state[pos_key] = net_qty
                                    self.log(f"Synced: {strike[:15]}... ({net_qty})")
                else:
                    self.update_status("RUNNING - ODIN NOT FOUND", "#f1c40f")
                    
            except Exception as e:
                # Silently log errors to the internal console occasionally
                pass
                
            time.sleep(1.0) # Check every second for stability

if __name__ == "__main__":
    root = tk.Tk()
    app = OdinScraperApp(root)
    root.mainloop()
