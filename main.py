import sys
import os
import json
from PyQt5.QtWidgets import (QApplication, QSystemTrayIcon, QMenu, QAction, 
                             QDialog, QLineEdit, QPushButton, QFormLayout, QMessageBox, QFileDialog, QHBoxLayout, QTabWidget, QWidget, QVBoxLayout)
from PyQt5.QtGui import QIcon, QPixmap, QColor
from PyQt5.QtCore import Qt, QTimer
from overlays import TimeOverlay, PremiumOverlay
from excel_reader import ExcelReader

class ConfigManager:
    def __init__(self, filename="settings.json"):
        # Store in the same directory as the executable/script
        self.filename = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), filename)

    def load(self):
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r') as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save(self, data):
        try:
            with open(self.filename, 'w') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

class ExcelConfigWidget(QWidget):
    def __init__(self, excel_reader, title="Excel Settings", parent=None):
        super().__init__(parent)
        self.excel_reader = excel_reader
        
        layout = QFormLayout()

        # Workbook Name with Browse button
        self.wb_input = QLineEdit()
        self.wb_input.setText(self.excel_reader.wb_name or "")
        
        wb_layout = QHBoxLayout()
        wb_layout.addWidget(self.wb_input)
        
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_file)
        wb_layout.addWidget(browse_btn)
        
        layout.addRow("Workbook Path/Name:", wb_layout)

        self.sheet_input = QLineEdit()
        self.sheet_input.setText(self.excel_reader.sheet_name or "Sheet1")
        layout.addRow("Sheet Name:", self.sheet_input)

        self.cell1_input = QLineEdit()
        self.cell1_input.setText(self.excel_reader.cells[0] if self.excel_reader.cells else "A1")
        layout.addRow("Cell 1 (e.g. A1):", self.cell1_input)

        self.cell2_input = QLineEdit()
        self.cell2_input.setText(self.excel_reader.cells[1] if self.excel_reader.cells else "B1")
        layout.addRow("Cell 2 (e.g. B1):", self.cell2_input)

        self.cell3_input = QLineEdit()
        self.cell3_input.setText(self.excel_reader.cells[2] if self.excel_reader.cells else "C1")
        layout.addRow("Cell 3 (e.g. C1):", self.cell3_input)

        self.setLayout(layout)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls *.xlsm)")
        if file_path:
            self.wb_input.setText(file_path)

    def get_data(self):
        return (
            self.wb_input.text(),
            self.sheet_input.text(),
            self.cell1_input.text(),
            self.cell2_input.text(),
            self.cell3_input.text()
        )

class SettingsDialog(QDialog):
    def __init__(self, reader_bsc, reader_nsc, parent=None):
        super().__init__(parent)
        self.reader_bsc = reader_bsc
        self.reader_nsc = reader_nsc
        
        self.setWindowTitle("Excel Configuration (BSC & NSC)")
        self.setFixedSize(450, 350)

        main_layout = QVBoxLayout()
        
        self.tabs = QTabWidget()
        self.bsc_widget = ExcelConfigWidget(self.reader_bsc, "BSC Settings")
        self.nsc_widget = ExcelConfigWidget(self.reader_nsc, "NSC Settings")
        
        self.tabs.addTab(self.bsc_widget, "BSC Settings")
        self.tabs.addTab(self.nsc_widget, "NSC Settings")
        
        main_layout.addWidget(self.tabs)

        save_btn = QPushButton("Save All Settings")
        save_btn.clicked.connect(self.save)
        main_layout.addWidget(save_btn)

        self.setLayout(main_layout)

    def save(self):
        # Save BSC
        bsc_data = self.bsc_widget.get_data()
        self.reader_bsc.set_config(*bsc_data)
        
        # Save NSC
        nsc_data = self.nsc_widget.get_data()
        self.reader_nsc.set_config(*nsc_data)
        
        self.accept()

class OdinOverlayApp:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.app.setQuitOnLastWindowClosed(False)
        
        self.config_manager = ConfigManager()
        self.settings = self.config_manager.load()

        # Time Overlay
        self.time_overlay = TimeOverlay()
        
        # Screen geometry for positioning
        screen = QApplication.primaryScreen().geometry()
        
        # BSC Overlay (Top position)
        self.premium_overlay_bsc = PremiumOverlay(
            x=screen.width() - 290, 
            y=screen.height() - 200, 
            label_prefix="BSC"
        )
        self.excel_reader_bsc = ExcelReader()
        
        # NSC Overlay (Bottom position, above clock)
        self.premium_overlay_nsc = PremiumOverlay(
            x=screen.width() - 290, 
            y=screen.height() - 145, 
            label_prefix="NSC"
        )
        self.excel_reader_nsc = ExcelReader()

        # Apply Loaded Settings
        self.apply_settings()

        self.excel_timer = QTimer()
        self.excel_timer.timeout.connect(self.update_excel_data)
        
        self.setup_tray()

    def apply_settings(self):
        bsc = self.settings.get("bsc", {})
        if bsc:
            self.excel_reader_bsc.set_config(
                bsc.get("wb", ""), bsc.get("sheet", ""), 
                bsc.get("c1", "A1"), bsc.get("c2", "B1"), bsc.get("c3", "C1")
            )
        
        nsc = self.settings.get("nsc", {})
        if nsc:
            self.excel_reader_nsc.set_config(
                nsc.get("wb", ""), nsc.get("sheet", ""), 
                nsc.get("c1", "A1"), nsc.get("c2", "B1"), nsc.get("c3", "C1")
            )

    def save_settings(self):
        data = {
            "bsc": {
                "wb": self.excel_reader_bsc.wb_name,
                "sheet": self.excel_reader_bsc.sheet_name,
                "c1": self.excel_reader_bsc.cells[0],
                "c2": self.excel_reader_bsc.cells[1],
                "c3": self.excel_reader_bsc.cells[2]
            },
            "nsc": {
                "wb": self.excel_reader_nsc.wb_name,
                "sheet": self.excel_reader_nsc.sheet_name,
                "c1": self.excel_reader_nsc.cells[0],
                "c2": self.excel_reader_nsc.cells[1],
                "c3": self.excel_reader_nsc.cells[2]
            }
        }
        self.config_manager.save(data)

    def create_icon(self):
        pixmap = QPixmap(64, 64)
        pixmap.fill(QColor("transparent"))
        import PyQt5.QtGui as QtGui
        painter = QtGui.QPainter(pixmap)
        painter.setBrush(QColor("blue"))
        painter.drawEllipse(0, 0, 64, 64)
        painter.setPen(QColor("white"))
        font = painter.font()
        font.setPointSize(24)
        painter.setFont(font)
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "O")
        painter.end()
        return QIcon(pixmap)

    def setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self.create_icon(), self.app)
        self.tray_icon.setToolTip("Odin Overlay Tools")

        menu = QMenu()

        # Time Overlay
        self.action_time = QAction("Show Time Overlay", menu, checkable=True)
        self.action_time.triggered.connect(self.toggle_time_overlay)
        menu.addAction(self.action_time)

        menu.addSeparator()

        # BSC NSC Options
        self.action_bsc = QAction("Show BSC Premium", menu, checkable=True)
        self.action_bsc.triggered.connect(self.toggle_premium_overlays)
        menu.addAction(self.action_bsc)

        self.action_nsc = QAction("Show NSC Premium", menu, checkable=True)
        self.action_nsc.triggered.connect(self.toggle_premium_overlays)
        menu.addAction(self.action_nsc)

        # Config
        action_config = QAction("Configure BSC/NSC Excel...", menu)
        action_config.triggered.connect(self.open_settings)
        menu.addAction(action_config)

        menu.addSeparator()

        action_exit = QAction("Exit", menu)
        action_exit.triggered.connect(self.exit_app)
        menu.addAction(action_exit)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()

    def toggle_time_overlay(self, checked):
        if checked:
            self.time_overlay.show()
        else:
            self.time_overlay.hide()

    def toggle_premium_overlays(self):
        show_bsc = self.action_bsc.isChecked()
        show_nsc = self.action_nsc.isChecked()

        if show_bsc:
            self.premium_overlay_bsc.show()
        else:
            self.premium_overlay_bsc.hide()

        if show_nsc:
            self.premium_overlay_nsc.show()
        else:
            self.premium_overlay_nsc.hide()

        if show_bsc or show_nsc:
            if not self.excel_timer.isActive():
                self.excel_timer.start(100)
        else:
            self.excel_timer.stop()

    def open_settings(self):
        dialog = SettingsDialog(self.excel_reader_bsc, self.excel_reader_nsc)
        if dialog.exec_() == QDialog.Accepted:
            self.save_settings()

    def update_excel_data(self):
        if self.action_bsc.isChecked():
            c1, c2, c3 = self.excel_reader_bsc.read_cells()
            self.premium_overlay_bsc.update_data(c1, c2, c3)
        
        if self.action_nsc.isChecked():
            c1, c2, c3 = self.excel_reader_nsc.read_cells()
            self.premium_overlay_nsc.update_data(c1, c2, c3)

    def exit_app(self):
        self.tray_icon.hide()
        self.app.quit()

    def run(self):
        sys.exit(self.app.exec_())

if __name__ == "__main__":
    app = OdinOverlayApp()
    app.run()
