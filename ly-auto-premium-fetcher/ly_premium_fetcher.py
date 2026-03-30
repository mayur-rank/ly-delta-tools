import sys
import os
import time
import json
import urllib.request
import http.cookiejar
from datetime import datetime, timezone, timedelta
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QComboBox, 
                             QTableWidget, QTableWidgetItem, QHeaderView, 
                             QMessageBox, QFrame)
from PyQt5.QtCore import QTimer, Qt, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QColor

# Add parent directory to path for modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from excel_reader import ExcelReader
    from time_utils import TimeSyncer
except ImportError:
    # Fallback for local testing if needed
    class ExcelReader: 
        def set_config(self, *args): pass
        def append_row(self, *args, **kwargs): pass
    class TimeSyncer:
        def sync(self): return True
        def get_current_time(self): return time.time(), True, "System"

import gzip
import zlib

class TradingDataFetcher:
    """Robust data fetcher using multiple free sources (Upstox JSON & TopStockResearch HTML)."""
    
    # Primary: Upstox JSON API
    UPSTOX_API = "https://service.upstox.com/option-analytics-tool/open/v1/strategy-chains?assetKey=NSE_INDEX%7CNifty+50&strategyChainType=PC_CHAIN"
    
    # Fallback: TopStockResearch (Moneycontrol often uses this)
    FALLBACK_URL = "https://www.topstockresearch.com/rt/ViewOptionChain"

    def __init__(self):
        self.opener = urllib.request.build_opener()

    def _decompress(self, body, encoding):
        if encoding == 'gzip':
            try: return gzip.decompress(body)
            except: pass
        elif encoding == 'deflate':
            try: return zlib.decompress(body, -zlib.MAX_WBITS)
            except: pass
        return body

    def get_upstox_data(self):
        """Fetches data from Upstox JSON API."""
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            "Accept": "application/json",
            "Referer": "https://upstox.com/option-chain/nifty/",
            "Origin": "https://upstox.com"
        }
        try:
            req = urllib.request.Request(self.UPSTOX_API, headers=headers)
            with self.opener.open(req, timeout=10) as response:
                body = self._decompress(response.read(), response.info().get('Content-Encoding'))
                data = json.loads(body.decode('utf-8'))
            
            p_data = data.get('data', {})
            spot = p_data.get('underlyingLTP') or p_data.get('underlyingDetails', {}).get('ltp', 0)
            chain = p_data.get('strategyChain', [])
            expiry = p_data.get('selectedExpiryDate', "Unknown")
            
            if not spot or not chain: return None
            
            atm_strike = int(round(spot / 50) * 50)
            ce, pe = 0, 0
            for item in chain:
                if int(item.get('strikePrice', 0)) == atm_strike:
                    ce = item.get('call', {}).get('ltp', 0)
                    pe = item.get('put', {}).get('ltp', 0)
                    break
            return spot, atm_strike, expiry, ce, pe
        except Exception as e:
            print(f"Upstox API Error: {e}")
            return None

    def get_topstock_data(self):
        """Fallback: Fetches data from TopStockResearch HTML with precise indexing."""
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        }
        try:
            req = urllib.request.Request(self.FALLBACK_URL, headers=headers)
            with self.opener.open(req, timeout=10) as response:
                html = response.read().decode('utf-8', errors='ignore')
            
            # 1. Extract Spot Price (e.g., 'Spot Close : 23428.35' or 'NIFTY : 23428.35')
            import re
            spot_match = re.search(r"(?:Spot Close|NIFTY)\s*[:\s]+([\d\.,]+)", html, re.IGNORECASE)
            if not spot_match:
                # Last resort: look for any large number near 'Nifty'
                spot_match = re.search(r"Nifty[^<]*>([\d\.,]+)", html, re.IGNORECASE)
                
            if not spot_match: return None
            spot = float(spot_match.group(1).replace(',', ''))
            
            # 2. Extract Expiry (often in a dropdown or header)
            expiry_match = re.search(r"Expiry\s*:\s*([^<]+)", html, re.IGNORECASE)
            expiry = expiry_match.group(1).strip() if expiry_match else "Market"
            
            atm_strike = int(round(spot / 50) * 50)
            
            # 3. Parse Table Rows
            # The table is likely in id="results" or just the largest table on page
            if 'id="results"' in html:
                table_html = html.split('id="results"')[1].split('</table>')[0]
            else:
                table_html = html # Fallback to full HTML if ID missing
                
            rows = table_html.split('<tr')[1:] # Skip first potential header or split artifact
            
            ce_ltp, pe_ltp = 0, 0
            found = False
            
            for row in rows:
                # Split columns by <td>
                cols = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
                if len(cols) >= 7:
                    # Strip HTML tags from column content
                    clean_cols = [re.sub(r'<[^>]+>', '', c).strip() for c in cols]
                    try:
                        # Index 0: Strike Price
                        strike_val = float(clean_cols[0].replace(',', ''))
                        if int(strike_val) == atm_strike:
                            # Index 5: Call LTP, Index 6: Put LTP
                            ce_ltp = float(clean_cols[5].replace(',', '')) if clean_cols[5] else 0
                            pe_ltp = float(clean_cols[6].replace(',', '')) if clean_cols[6] else 0
                            found = True
                            break
                    except: continue
            
            if found:
                return spot, atm_strike, expiry, ce_ltp, pe_ltp
            return spot, atm_strike, expiry, 0, 0
        except Exception as e:
            print(f"TopStock Fallback Error: {e}")
            return None

    def get_data(self):
        """Tries primary then fallback."""
        data = self.get_upstox_data()
        if data: return data
        
        print("Switching to TopStock Fallback...")
        data = self.get_topstock_data()
        if data: return data
        
        return None, None, "Fail", 0, 0

class FetchWorker(QThread):
    """Thread to handle network calls without freezing UI."""
    data_fetched = pyqtSignal(dict)
    
    def __init__(self, fetcher):
        super().__init__()
        self.fetcher = fetcher

    def run(self):
        spot, strike, expiry, ce, pe = self.fetcher.get_data()
        
        results = {
            "spot": spot,
            "strike": strike,
            "expiry": expiry,
            "ce": ce,
            "pe": pe,
            "total": (ce + pe) if ce and pe else 0
        }
        self.data_fetched.emit(results)

class LYPremiumFetcher(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("LY Auto Premium Fetcher")
        self.resize(650, 700)
        
        # State
        self.fetching = False
        self.last_premium = None
        self.fetcher = TradingDataFetcher()
        self.syncer = TimeSyncer()
        self.excel_reader = ExcelReader()
        
        # Log directory handling
        log_dir = os.path.dirname(os.path.abspath(__file__))
        self.log_path = os.path.join(log_dir, "Premium_Logs.xlsx")
        self.excel_reader.set_config(self.log_path, "DailyLogs", "", "", "")
        
        self.timer = QTimer()
        self.timer.timeout.connect(self.check_time)
        
        self.init_ui()
        
        # Sync time on startup
        self.syncer.sync()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # --- Top Dashboard ---
        dash_frame = QFrame()
        dash_frame.setFrameShape(QFrame.StyledPanel)
        dash_frame.setLineWidth(2)
        dash_frame.setStyleSheet("background-color: #f8f9fa;")
        dash_layout = QVBoxLayout(dash_frame)
        
        # Spot & Strike Row
        row1 = QHBoxLayout()
        self.lbl_spot = QLabel("Nifty Spot: --")
        self.lbl_spot.setFont(QFont("Segoe UI", 12, QFont.Bold))
        row1.addWidget(self.lbl_spot)
        
        self.lbl_strike = QLabel("ATM Strike: --")
        self.lbl_strike.setFont(QFont("Segoe UI", 12))
        row1.addWidget(self.lbl_strike)
        dash_layout.addLayout(row1)

        # Individual Prices Row
        row_prices = QHBoxLayout()
        self.lbl_ce = QLabel("CE LTP: --")
        self.lbl_ce.setFont(QFont("Segoe UI", 11))
        row_prices.addWidget(self.lbl_ce)
        
        self.lbl_pe = QLabel("PE LTP: --")
        self.lbl_pe.setFont(QFont("Segoe UI", 11))
        row_prices.addWidget(self.lbl_pe)
        dash_layout.addLayout(row_prices)
        
        # Premium & Melt Row
        row2 = QHBoxLayout()
        self.lbl_premium = QLabel("Total Premium: --")
        self.lbl_premium.setFont(QFont("Segoe UI", 24, QFont.Bold))
        self.lbl_premium.setStyleSheet("color: #007bff;")
        row2.addWidget(self.lbl_premium)
        
        self.lbl_melt = QLabel("Melt/Spike: 0.00")
        self.lbl_melt.setFont(QFont("Segoe UI", 16, QFont.Bold))
        row2.addWidget(self.lbl_melt)
        dash_layout.addLayout(row2)
        
        layout.addWidget(dash_frame)
        
        # --- Controls ---
        ctrl_layout = QHBoxLayout()
        ctrl_layout.addWidget(QLabel("Interval:"))
        self.combo_interval = QComboBox()
        self.combo_interval.addItems(["30 seconds", "1 minute"])
        ctrl_layout.addWidget(self.combo_interval)
        
        self.btn_start = QPushButton("START FETCHING")
        self.btn_start.setFixedHeight(50)
        self.btn_start.setStyleSheet("""
            QPushButton {
                background-color: #28a745; 
                color: white; 
                font-size: 14px; 
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #218838; }
        """)
        self.btn_start.clicked.connect(self.toggle_fetching)
        ctrl_layout.addWidget(self.btn_start)
        
        layout.addLayout(ctrl_layout)
        
        # --- History Table ---
        layout.addWidget(QLabel("Fetch History (Last 50):"))
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["Time", "Spot", "Strike", "CE LTP", "PE LTP", "Total", "Diff"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table)
        
        # --- Status Bar ---
        self.status_lbl = QLabel("Status: Idle")
        self.status_lbl.setStyleSheet("color: #6c757d;")
        layout.addWidget(self.status_lbl)

    def toggle_fetching(self):
        if not self.fetching:
            self.fetching = True
            self.btn_start.setText("STOP FETCHING")
            self.btn_start.setStyleSheet("""
                QPushButton {
                    background-color: #dc3545; 
                    color: white; 
                    font-size: 14px; 
                    font-weight: bold;
                    border-radius: 5px;
                }
                QPushButton:hover { background-color: #c82333; }
            """)
            self.status_lbl.setText("Status: Running (Waiting for aligned time...)")
            self.timer.start(100) # Fast check for alignment
        else:
            self.fetching = False
            self.btn_start.setText("START FETCHING")
            self.btn_start.setStyleSheet("""
                QPushButton {
                    background-color: #28a745; 
                    color: white; 
                    font-size: 14px; 
                    font-weight: bold;
                    border-radius: 5px;
                }
                QPushButton:hover { background-color: #218838; }
            """)
            self.status_lbl.setText("Status: Stopped")
            self.timer.stop()

    def check_time(self):
        if not self.fetching: return
        
        ts, is_sync, _ = self.syncer.get_current_time()
        # IST Date
        dt = datetime.fromtimestamp(ts, tz=timezone(timedelta(hours=5, minutes=30)))
        
        # Check market hours (9:15 to 15:30)
        market_start = dt.replace(hour=9, minute=15, second=0, microsecond=0)
        market_end = dt.replace(hour=15, minute=30, second=0, microsecond=0)
        
        if not (market_start <= dt <= market_end):
            self.status_lbl.setText(f"Market Closed. Time: {dt.strftime('%H:%M:%S')}")
            return

        interval_str = self.combo_interval.currentText()
        interval = 60 if "1 minute" in interval_str else 30
        
        # Precise bucket check
        bucket_id = int(ts // interval)
        if not hasattr(self, 'last_bucket'): 
            # First tick: seed the bucket but don't fetch yet to ensure alignment
            self.last_bucket = bucket_id
            return
        
        if bucket_id > self.last_bucket:
            self.last_bucket = bucket_id
            self.trigger_fetch(dt)

    def trigger_fetch(self, dt):
        self.status_lbl.setText(f"Fetching at {dt.strftime('%H:%M:%S')}...")
        self.worker = FetchWorker(self.fetcher)
        self.worker.data_fetched.connect(lambda data: self.handle_fetched_data(data, dt))
        self.worker.start()

    def handle_fetched_data(self, data, dt):
        if data['spot'] is None:
            self.status_lbl.setText("Error: Could not reach NSE API")
            return

        # Update Labels
        self.lbl_spot.setText(f"Nifty Spot: {data['spot']:.2f}")
        self.lbl_strike.setText(f"ATM Strike: {data['strike']} ({data['expiry']})")
        self.lbl_ce.setText(f"CE LTP: {data['ce']:.2f}")
        self.lbl_pe.setText(f"PE LTP: {data['pe']:.2f}")
        self.lbl_premium.setText(f"Total Premium: {data['total']:.2f}")
        
        # Calculate Diff
        diff = 0
        if self.last_premium is not None:
            diff = data['total'] - self.last_premium
        
        self.last_premium = data['total']
        
        # Color logic: Red for melt (diff < 0), Green/Blue for spike
        color = "#dc3545" if diff < 0 else "#007bff"
        self.lbl_melt.setText(f"{'Melt' if diff < 0 else 'Spike'}: {diff:+.2f}")
        self.lbl_melt.setStyleSheet(f"color: {color};")
        
        # Add to Table
        time_str = dt.strftime('%H:%M:%S')
        self.table.insertRow(0)
        self.table.setItem(0, 0, QTableWidgetItem(time_str))
        self.table.setItem(0, 1, QTableWidgetItem(f"{data['spot']:.2f}"))
        self.table.setItem(0, 2, QTableWidgetItem(str(data['strike'])))
        self.table.setItem(0, 3, QTableWidgetItem(f"{data['ce']:.2f}"))
        self.table.setItem(0, 4, QTableWidgetItem(f"{data['pe']:.2f}"))
        self.table.setItem(0, 5, QTableWidgetItem(f"{data['total']:.2f}"))
        
        diff_item = QTableWidgetItem(f"{diff:+.2f}")
        diff_item.setForeground(QColor(color))
        self.table.setItem(0, 6, diff_item)
        
        if self.table.rowCount() > 50:
            self.table.removeRow(50)
            
        # Log to Excel
        self.log_to_excel(dt, data, diff)
        self.status_lbl.setText(f"Last fetched: {time_str}")

    def log_to_excel(self, dt, data, diff):
        sheet_name = dt.strftime('%Y-%m-%d')
        metadata = [f"Date: {sheet_name}", f"Day: {dt.strftime('%A')}", "Type: INTERNET_SYNC"]
        headers = ["Time", "Spot Price", "ATM Strike", "CE LTP", "PE LTP", "Total Premium", "Melt/Spike"]
        row = [dt.strftime('%H:%M:%S'), data['spot'], data['strike'], data['ce'], data['pe'], data['total'], diff]
        
        try:
            self.excel_reader.append_row(sheet_name, row, headers, metadata)
        except Exception as e:
            print(f"Excel Logging Error: {e}")

if __name__ == "__main__":
    # Ensure fonts look better on high DPI
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    
    app = QApplication(sys.argv)
    window = LYPremiumFetcher()
    window.show()
    sys.exit(app.exec_())
