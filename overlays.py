import sys
from PyQt5.QtWidgets import QWidget, QLabel, QVBoxLayout, QHBoxLayout, QApplication
from PyQt5.QtCore import Qt, QTimer, QTime

class TimeOverlay(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)

    def initUI(self):
        # Frameless, Always on Top, Click-through (WindowTransparentForInput), and Tool (no taskbar icon)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool | Qt.WindowTransparentForInput)
        self.setAttribute(Qt.WA_TranslucentBackground)

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        self.time_label = QLabel()
        self.time_label.setStyleSheet("color: red; font-size: 15px; font-weight: bold; background-color: rgba(0,0,0,150); padding: 5px; border-radius: 5px;")
        layout.addWidget(self.time_label)
        self.setLayout(layout)
        self.update_time()

        # Place the overlay at the bottom right (approximate taskbar clock location)
        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(screen.width() - 155, screen.height() - 90, 80, 30)

    def update_time(self):
        current_time = QTime.currentTime().toString('hh:mm:ss AP')
        self.time_label.setText(current_time)


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
            self.title_label.setStyleSheet("color: #00FF00; font-size: 10px; font-weight: bold; background-color: rgba(0,0,0,150); padding: 2px; border-top-left-radius: 5px; border-top-right-radius: 5px;")
            main_layout.addWidget(self.title_label)

        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        
        self.cell1_label = QLabel("N/A")
        self.cell2_label = QLabel("N/A")
        self.cell3_label = QLabel("N/A")

        # Semi-transparent black background so white text is readable (User preferred 15px)
        # style = "color: white; font-size: 15px; font-weight: bold; background-color: rgba(0,0,0,150); padding: 5px; border-radius: 5px;"
        style = "color: white; font-size: 15px; font-weight: bold; background-color: rgba(0,0,0,150); padding: 5px; border-bottom-left-radius: 5px; border-bottom-right-radius: 5px;"
        
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

    def update_data(self, c1, c2, c3):
        self.cell1_label.setText(str(c1))
        self.cell2_label.setText(str(c2))
        self.cell3_label.setText(str(c3))
