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
        self.initUI()
        
        self.background_sync() # Initial sync
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)

        # Periodic sync timer (every 5 minutes)
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.background_sync)
        self.sync_timer.start(5 * 60 * 1000)

    def initUI(self):
        # Frameless, Always on Top, Click-through (WindowTransparentForInput), and Tool (no taskbar icon)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool | Qt.WindowTransparentForInput)
        self.setAttribute(Qt.WA_TranslucentBackground)

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        self.time_label = QLabel()
        self.time_label.setStyleSheet("color: white; font-size: 15px; font-weight: bold; background-color: rgba(0,0,0,180); padding: 5px; border-radius: 5px;")
        layout.addWidget(self.time_label)
        self.setLayout(layout)
        self.update_time()

        # Place the overlay at the bottom right (approximate taskbar clock location)
        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(screen.width() - 155, screen.height() - 90, 80, 30)

    def background_sync(self):
        # Run sync in a background thread to avoid freezing the UI
        thread = threading.Thread(target=self.syncer.sync, daemon=True)
        thread.start()

    def update_time(self):
        timestamp, is_synced, source = self.syncer.get_current_time()
        
        # Indian Standard Time (IST) offset is UTC + 5:30
        ist_tz = timezone(timedelta(hours=5, minutes=30))
        dt = datetime.fromtimestamp(timestamp, tz=ist_tz)
        
        # Force format to hh:mm:ss AM/PM
        current_time_str = dt.strftime('%I:%M:%S %p').lstrip('0')
        self.time_label.setText(current_time_str)
        
        # Slightly change appearance if NOT synced (e.g. gray out or red dot)
        if not is_synced:
            self.time_label.setStyleSheet("color: #AAAAAA; font-size: 15px; font-weight: bold; background-color: rgba(0,0,0,180); padding: 5px; border-radius: 5px;")
        else:
            self.time_label.setStyleSheet("color: white; font-size: 15px; font-weight: bold; background-color: rgba(0,0,0,180); padding: 5px; border-radius: 5px;")


class PremiumOverlay(QWidget):
    def __init__(self, x=None, y=None, label_prefix=""):
        super().__init__()
        self.x_pos = x
        self.y_pos = y
        self.label_prefix = label_prefix
        self.initUI()
    
    def initUI(self):
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool | Qt.WindowTransparentForInput)
        self.setAttribute(Qt.WA_TranslucentBackground)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(2)

        # Title Label (BSC/NSC)
        if self.label_prefix:
            self.title_label = QLabel(self.label_prefix)
            self.title_label.setStyleSheet("color: #00FF00; font-size: 10px; font-weight: bold; background-color: rgba(0,0,0,180); padding: 2px; border-top-left-radius: 5px; border-top-right-radius: 5px;")
            main_layout.addWidget(self.title_label)

        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        
        self.cell1_label = QLabel("N/A")
        self.cell2_label = QLabel("N/A")
        self.cell3_label = QLabel("N/A")

        # Semi-transparent black background so white text is readable (User preferred 15px)
        # style = "color: white; font-size: 15px; font-weight: bold; background-color: rgba(0,0,0,150); padding: 5px; border-radius: 5px;"
        style = "color: white; font-size: 15px; font-weight: bold; background-color: rgba(0,0,0,180); padding: 5px; border-bottom-left-radius: 5px; border-bottom-right-radius: 5px;"
        
        self.cell1_label.setStyleSheet(style)
        self.cell2_label.setStyleSheet(style)
        self.cell3_label.setStyleSheet(style)

        content_layout.addWidget(self.cell1_label)
        content_layout.addWidget(self.cell2_label)
        content_layout.addWidget(self.cell3_label)

        main_layout.addLayout(content_layout)
        self.setLayout(main_layout)

        # Place at specified position or default
        screen = QApplication.primaryScreen().geometry()
        qx = self.x_pos if self.x_pos is not None else (screen.width() - 290)
        qy = self.y_pos if self.y_pos is not None else (screen.height() - 150)
        
        self.setGeometry(qx, qy, 250, 50)

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
