import sys
from PyQt5.QtWidgets import QWidget, QLabel, QVBoxLayout, QHBoxLayout, QApplication
from PyQt5.QtCore import Qt, QTimer, QTime
from datetime import datetime, timezone, timedelta
import threading
from time_utils import TimeSyncer

class TimeOverlay(QWidget):
    def __init__(self):
        super().__init__()
        self.syncer = TimeSyncer()
        self.scale = 1.0
        self.x_pos = -1
        self.y_pos = -1
        
        # New Alert Properties
        self.alert_enabled = False
        self.alert_seconds = 5
        self.initUI()
        
        self.background_sync() # Initial sync
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(33) # ~30 FPS for ultra-smooth transition

        # Periodic sync timer (every 5 minutes)
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.background_sync)
        self.sync_timer.start(2 * 60 * 1000) # Sync every 2 minutes to correct hardware clock drift

    def initUI(self):
        # Frameless, Always on Top, Click-through (WindowTransparentForInput), and Tool (no taskbar icon)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool | Qt.WindowTransparentForInput)
        self.setAttribute(Qt.WA_TranslucentBackground)

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        self.time_label = QLabel()
        layout.addWidget(self.time_label)
        self.setLayout(layout)
        
        self.apply_style()
        self.apply_geometry()
        self.update_time()

    def set_alert_config(self, enabled, seconds):
        self.alert_enabled = enabled
        self.alert_seconds = seconds

    def apply_style(self, color=None, is_alert=False):
        if color is None:
            # Re-read color from current state if not provided
            _, is_synced, source = self.syncer.get_current_time()
            if not is_synced or source == "System":
                color = "#AAAAAA" # Gray (Not Synced)
            elif "HTTP" in source:
                color = "#00FFFF" # Cyan (High-Precision HTTP Fallback)
            else:
                color = "white"   # White (NTP - Best)

        font_size = int(15 * self.scale)
        padding = int(5 * self.scale)
        
        bg_color = "rgba(0,0,0,180)"
        if is_alert:
            # Flashing Red
            ms = QTime.currentTime().msec()
            if ms < 500:
                bg_color = "rgba(180,0,0,220)" # Alert state (Red)
        
        style = f"color: {color}; font-size: {font_size}px; font-weight: bold; background-color: {bg_color}; padding: {padding}px; border-radius: 5px;"
        self.time_label.setStyleSheet(style)

    def apply_geometry(self):
        screen = QApplication.primaryScreen().geometry()
        width = int(110 * self.scale)
        height = int(35 * self.scale)
        qx = self.x_pos if self.x_pos != -1 else (screen.width() - width - 20)
        qy = self.y_pos if self.y_pos != -1 else (screen.height() - height - 60)
        self.setGeometry(qx, qy, width, height)

    def update_scale(self, scale):
        self.scale = scale
        self.apply_style()
        self.apply_geometry()

    def move_to(self, x, y):
        self.x_pos = x
        self.y_pos = y
        self.apply_geometry()

    def background_sync(self):
        # Run sync in a background thread to avoid freezing the UI
        thread = threading.Thread(target=self.syncer.sync, daemon=True)
        thread.start()

    def update_time(self):
        timestamp, is_synced, source = self.syncer.get_current_time()
        
        ist_tz = timezone(timedelta(hours=5, minutes=30))
        dt = datetime.fromtimestamp(timestamp, tz=ist_tz)
        
        current_time_str = dt.strftime('%I:%M:%S %p').lstrip('0')
        self.time_label.setText(current_time_str)
        
        # Diagnostics Tooltip
        rtt_ms = int(self.syncer.last_rtt * 1000)
        manual_ms = int(self.syncer.manual_offset * 1000)
        tooltip = f"Source: {source}\nRTT: {rtt_ms}ms\nManual Sync: {manual_ms}ms\nAccuracy: ±{rtt_ms//2}ms"
        self.time_label.setToolTip(tooltip)

        # Update color only (don't force full restyle if color didn't change)
        if not is_synced or source == "System":
            color = "#AAAAAA"
        elif "HTTP" in source:
            color = "#00FFFF"
        else:
            color = "white"
        
        # --- NEW ALERT LOGIC ---
        is_alert = False
        if self.alert_enabled:
            # Calculate seconds remaining in minute
            # Use synchronized timestamp
            ist_tz = timezone(timedelta(hours=5, minutes=30))
            dt = datetime.fromtimestamp(timestamp, tz=ist_tz)
            seconds_remaining = 60 - dt.second
            
            if seconds_remaining <= self.alert_seconds and seconds_remaining > 0:
                is_alert = True
        
        self.apply_style(color, is_alert)


class PremiumOverlay(QWidget):
    def __init__(self, x=None, y=None, label_prefix="", scale=1.0):
        super().__init__()
        self.x_pos = x
        self.y_pos = y
        self.label_prefix = label_prefix
        self.scale = scale
        self.initUI()
    
    def initUI(self):
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool | Qt.WindowTransparentForInput)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(int(2 * self.scale))

        # Title Label (BSC/NSC)
        if self.label_prefix:
            self.title_label = QLabel(self.label_prefix)
            self.update_title_style()
            self.main_layout.addWidget(self.title_label)

        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(int(2 * self.scale))
        
        self.cell1_label = QLabel("N/A")
        self.cell2_label = QLabel("N/A")
        self.cell3_label = QLabel("N/A")

        self.update_content_style()

        content_layout.addWidget(self.cell1_label)
        content_layout.addWidget(self.cell2_label)
        content_layout.addWidget(self.cell3_label)

        self.main_layout.addLayout(content_layout)
        self.setLayout(self.main_layout)

        self.apply_geometry()

    def update_title_style(self):
        if not hasattr(self, 'title_label'): return
        font_size = int(10 * self.scale)
        padding = int(2 * self.scale)
        style = f"color: #00FF00; font-size: {font_size}px; font-weight: bold; background-color: rgba(0,0,0,180); padding: {padding}px; border-top-left-radius: 5px; border-top-right-radius: 5px;"
        self.title_label.setStyleSheet(style)

    def update_content_style(self):
        font_size = int(15 * self.scale)
        padding = int(5 * self.scale)
        style = f"color: white; font-size: {font_size}px; font-weight: bold; background-color: rgba(0,0,0,180); padding: {padding}px; border-bottom-left-radius: 5px; border-bottom-right-radius: 5px;"
        self.cell1_label.setStyleSheet(style)
        self.cell2_label.setStyleSheet(style)
        self.cell3_label.setStyleSheet(style)

    def apply_geometry(self):
        screen = QApplication.primaryScreen().geometry()
        width = int(250 * self.scale)
        height = int(50 * self.scale)
        qx = self.x_pos if self.x_pos is not None and self.x_pos != -1 else (screen.width() - width - 40)
        qy = self.y_pos if self.y_pos is not None and self.y_pos != -1 else (screen.height() - height - 100)
        self.setGeometry(qx, qy, width, height)

    def update_scale(self, scale):
        self.scale = scale
        self.update_title_style()
        self.update_content_style()
        self.main_layout.setSpacing(int(2 * self.scale))
        self.apply_geometry()

    def move_to(self, x, y):
        self.x_pos = x
        self.y_pos = y
        self.apply_geometry()

    def format_value(self, val):
        if val is None or val == "":
            return "N/A"
        try:
            # If it's a number, format it
            f_val = float(val)
            # If it's an integer-like value (e.g. 130.0), show as integer
            if f_val == int(f_val):
                return str(int(f_val))
            # Otherwise, show with 1 decimal place (e.g. 130.9)
            return f"{f_val:.1f}"
        except (ValueError, TypeError):
            # If not a number, return as string
            return str(val)

    def update_data(self, c1, c2, c3):
        self.cell1_label.setText(self.format_value(c1))
        self.cell2_label.setText(self.format_value(c2))
        self.cell3_label.setText(self.format_value(c3))
