import sys
import os
import json
from PyQt5.QtWidgets import (QApplication, QSystemTrayIcon, QMenu, QAction, 
                             QDialog, QLineEdit, QPushButton, QFormLayout, QMessageBox, QFileDialog, QHBoxLayout, QTabWidget, QWidget, QVBoxLayout,
                             QCheckBox, QComboBox, QLabel, QDoubleSpinBox, QSpinBox, QSlider)
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
        layout.addRow("Cell 1 (Premium):", self.cell1_input)

        self.cell2_input = QLineEdit()
        self.cell2_input.setText(self.excel_reader.cells[1] if self.excel_reader.cells else "B1")
        layout.addRow("Cell 2:", self.cell2_input)

        self.cell3_input = QLineEdit()
        self.cell3_input.setText(self.excel_reader.cells[2] if self.excel_reader.cells else "C1")
        layout.addRow("Cell 3:", self.cell3_input)

        # Logging Settings
        layout.addRow("---", None)
        self.log_enabled = QCheckBox("Enable Premium Logging")
        self.log_enabled.setChecked(False) # Default to False, will be updated by parent
        layout.addRow(self.log_enabled)

        self.log_interval = QComboBox()
        self.log_interval.addItems(["1 second", "5 seconds", "30 seconds", "1 minute", "5 minutes"])
        self.log_interval.setCurrentText("1 minute")
        layout.addRow("Logging Interval:", self.log_interval)

        self.log_cell_source = QComboBox()
        self.log_cell_source.addItems(["Cell 1", "Cell 2", "Cell 3"])
        layout.addRow("Logging Source Cell:", self.log_cell_source)

        self.setLayout(layout)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls *.xlsm)")
        if file_path:
            self.wb_input.setText(file_path)

    def get_data(self):
        return {
            "wb": self.wb_input.text(),
            "sheet": self.sheet_input.text(),
            "c1": self.cell1_input.text(),
            "c2": self.cell2_input.text(),
            "c3": self.cell3_input.text(),
            "log_enabled": self.log_enabled.isChecked(),
            "log_interval": self.log_interval.currentText(),
            "log_source": self.log_cell_source.currentIndex()
        }

class SettingsDialog(QDialog):
    def __init__(self, reader_bsc, reader_nsc, settings_bsc, settings_nsc, settings_general, parent=None):
        super().__init__(parent)
        self.reader_bsc = reader_bsc
        self.reader_nsc = reader_nsc
        
        self.setWindowTitle("Global Configuration")
        self.setFixedSize(450, 520) 

        main_layout = QVBoxLayout()
        
        self.tabs = QTabWidget()
        
        # General Settings Tab
        self.general_widget = QWidget()
        gen_layout = QFormLayout(self.general_widget)
        self.time_adj_input = QLineEdit()
        self.time_adj_input.setPlaceholderText("e.g. 500 or -200")
        self.time_adj_input.setText(str(settings_general.get("time_adjustment", 0)))
        gen_layout.addRow("Time Adjustment (ms):", self.time_adj_input)
        gen_layout.addRow(QLabel("Positive moves clock forward, negative moves it back."))

        gen_layout.addRow("---", None)
        self.alert_enabled = QCheckBox("Enable Minute Alert")
        self.alert_enabled.setChecked(settings_general.get("minute_alert_enabled", False))
        gen_layout.addRow(self.alert_enabled)

        self.alert_seconds = QSpinBox()
        self.alert_seconds.setRange(1, 30)
        self.alert_seconds.setSuffix(" seconds")
        self.alert_seconds.setValue(settings_general.get("minute_alert_seconds", 5))
        gen_layout.addRow("Alert Start (before minute):", self.alert_seconds)

        self.tabs.addTab(self.general_widget, "General")

        # BSC/NSC Tabs
        self.bsc_widget = ExcelConfigWidget(self.reader_bsc, "BSC Settings")
        self.nsc_widget = ExcelConfigWidget(self.reader_nsc, "NSC Settings")
        
        self.apply_initial_settings(self.bsc_widget, settings_bsc)
        self.apply_initial_settings(self.nsc_widget, settings_nsc)

        self.tabs.addTab(self.bsc_widget, "BSC Settings")
        self.tabs.addTab(self.nsc_widget, "NSC Settings")
        
        main_layout.addWidget(self.tabs)

        save_btn = QPushButton("Save All Settings")
        save_btn.clicked.connect(self.save)
        main_layout.addWidget(save_btn)

        self.setLayout(main_layout)

    def apply_initial_settings(self, widget, settings):
        if not settings: return
        widget.log_enabled.setChecked(settings.get("log_enabled", False))
        widget.log_interval.setCurrentText(settings.get("log_interval", "1 minute"))
        widget.log_cell_source.setCurrentIndex(settings.get("log_source", 0))

    def save(self):
        # Save General
        try:
            self.general_data = {
                "time_adjustment": int(self.time_adj_input.text() or 0),
                "minute_alert_enabled": self.alert_enabled.isChecked(),
                "minute_alert_seconds": self.alert_seconds.value()
            }
        except:
            self.general_data = {
                "time_adjustment": 0,
                "minute_alert_enabled": False,
                "minute_alert_seconds": 5
            }

        # Save BSC
        self.bsc_data = self.bsc_widget.get_data()
        self.reader_bsc.set_config(
            self.bsc_data['wb'], self.bsc_data['sheet'], 
            self.bsc_data['c1'], self.bsc_data['c2'], self.bsc_data['c3']
        )
        
        # Save NSC
        self.nsc_data = self.nsc_widget.get_data()
        self.reader_nsc.set_config(
            self.nsc_data['wb'], self.nsc_data['sheet'], 
            self.nsc_data['c1'], self.nsc_data['c2'], self.nsc_data['c3']
        )
        
        self.accept()

class OverlaySettingsDialog(QDialog):
    def __init__(self, settings, app_instance, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.app = app_instance
        self.setWindowTitle("Global Overlay Adjuster")
        self.setFixedSize(400, 300)

        layout = QVBoxLayout()
        
        # Scale Slider
        layout.addWidget(QLabel("<b>Global Scale (Size)</b>"))
        self.scale_slider = QSlider(Qt.Horizontal)
        self.scale_slider.setRange(5, 40) # 0.5x to 4.0x
        initial_scale = settings.get("general", {}).get("global_scale", 1.0)
        self.scale_slider.setValue(int(initial_scale * 10))
        layout.addWidget(self.scale_slider)
        self.scale_label = QLabel(f"Current: {initial_scale:.1f}x")
        layout.addWidget(self.scale_label)

        layout.addSpacing(20)

        # Position Slider
        layout.addWidget(QLabel("<b>Global Position (Up/Down)</b>"))
        self.pos_slider = QSlider(Qt.Horizontal)
        self.pos_slider.setRange(-200, 800) # Offset from bottom
        initial_pos = settings.get("general", {}).get("global_y_offset", 0)
        self.pos_slider.setValue(initial_pos)
        layout.addWidget(self.pos_slider)
        self.pos_label = QLabel(f"Vertical Offset: {initial_pos}px")
        layout.addWidget(self.pos_label)

        layout.addSpacing(20)

        # Right Margin Slider
        layout.addWidget(QLabel("<b>Global Position (Right Margin)</b>"))
        self.margin_slider = QSlider(Qt.Horizontal)
        self.margin_slider.setRange(0, 500) # Right margin from 0 to 500px
        initial_margin = settings.get("general", {}).get("global_right_margin", 30)
        self.margin_slider.setValue(initial_margin)
        layout.addWidget(self.margin_slider)
        self.margin_label = QLabel(f"Right Margin: {initial_margin}px")
        layout.addWidget(self.margin_label)

        layout.addStretch()
        
        # Interactive Connections
        self.scale_slider.valueChanged.connect(self.on_visuals_changed)
        self.pos_slider.valueChanged.connect(self.on_visuals_changed)
        self.margin_slider.valueChanged.connect(self.on_visuals_changed)

        save_btn = QPushButton("Save & Finish")
        save_btn.clicked.connect(self.accept)
        layout.addWidget(save_btn)

        self.setLayout(layout)

    def on_visuals_changed(self):
        scale = self.scale_slider.value() / 10.0
        pos = self.pos_slider.value()
        margin = self.margin_slider.value()
        
        self.scale_label.setText(f"Current: {scale:.1f}x")
        self.pos_label.setText(f"Vertical Offset: {pos}px")
        self.margin_label.setText(f"Right Margin: {margin}px")
        
        # Real-time reflection
        self.app.update_all_overlays(scale, pos, margin)
        
        # Temporary store in settings object for persistence on save
        if "general" not in self.settings: self.settings["general"] = {}
        self.settings["general"]["global_scale"] = scale
        self.settings["general"]["global_y_offset"] = pos
        self.settings["general"]["global_right_margin"] = margin
        self.app.save_settings()

class OdinOverlayApp:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.app.setQuitOnLastWindowClosed(False)
        
        self.config_manager = ConfigManager()
        self.settings = self.config_manager.load()

        # Time Overlay
        self.time_overlay = TimeOverlay()
        self.syncer = self.time_overlay.syncer
        
        # Screen geometry for positioning
        screen = QApplication.primaryScreen().geometry()
        
        # BSC Overlay (Top position)
        self.premium_overlay_bsc = PremiumOverlay(
            label_prefix="BSC"
        )
        self.excel_reader_bsc = ExcelReader()
        
        # NSC Overlay (Bottom position, above clock)
        self.premium_overlay_nsc = PremiumOverlay(
            label_prefix="NSC"
        )
        self.excel_reader_nsc = ExcelReader()

        # Logging State
        self.log_state = {
            "BSC": {"last_time": 0, "last_val": None},
            "NSC": {"last_time": 0, "last_val": None}
        }

        # Apply Loaded Settings
        self.apply_settings()

        self.excel_timer = QTimer()
        self.excel_timer.timeout.connect(self.update_excel_data)
        
        self.setup_tray()

    def apply_settings(self):
        # Apply Time Adjustment
        general = self.settings.get("general", {})
        self.syncer.set_manual_adjustment(general.get("time_adjustment", 0))
        
        # Apply Alert Settings
        self.time_overlay.set_alert_config(
            general.get("minute_alert_enabled", False),
            general.get("minute_alert_seconds", 5)
        )

        # Global Visuals
        g_scale = general.get("global_scale", 1.0)
        g_y_off = general.get("global_y_offset", 0)
        g_margin = general.get("global_right_margin", 30)

        # Apply to all overlays
        self.update_all_overlays(g_scale, g_y_off, g_margin)

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

    def update_all_overlays(self, scale, y_offset, right_margin):
        # Apply scale to all
        self.time_overlay.update_scale(scale)
        self.premium_overlay_bsc.update_scale(scale)
        self.premium_overlay_nsc.update_scale(scale)

        # Standardized geometry for alignment
        screen = QApplication.primaryScreen().geometry()
        
        # All overlays should have the same right margin (scaled)
        scaled_margin = int(right_margin * scale)
        SPACING = int(8 * scale) # Consistent gap between blocks
        BOTTOM_GAP = int(60 * scale) # Base gap from bottom taskbar area
        
        # 1. Update Time Overlay (Lowest)
        tw, th = self.time_overlay.width(), self.time_overlay.height()
        tx = screen.width() - tw - scaled_margin
        ty = screen.height() - th - BOTTOM_GAP - y_offset
        self.time_overlay.move_to(tx, ty)

        # 2. Update NSC Overlay (Above Time)
        pw1, ph1 = self.premium_overlay_nsc.width(), self.premium_overlay_nsc.height()
        nx = screen.width() - pw1 - scaled_margin
        ny = ty - ph1 - SPACING
        self.premium_overlay_nsc.move_to(nx, ny)

        # 3. Update BSC Overlay (Above NSC)
        pw2, ph2 = self.premium_overlay_bsc.width(), self.premium_overlay_bsc.height()
        bx = screen.width() - pw2 - scaled_margin
        by = ny - ph2 - SPACING
        self.premium_overlay_bsc.move_to(bx, by)

    def save_settings(self):
        data = {
            "bsc": {
                "wb": self.excel_reader_bsc.wb_name,
                "sheet": self.excel_reader_bsc.sheet_name,
                "c1": self.excel_reader_bsc.cells[0],
                "c2": self.excel_reader_bsc.cells[1],
                "c3": self.excel_reader_bsc.cells[2],
                "log_enabled": self.settings.get("bsc", {}).get("log_enabled", False),
                "log_interval": self.settings.get("bsc", {}).get("log_interval", "1 minute"),
                "log_source": self.settings.get("bsc", {}).get("log_source", 0),
                "scale": self.settings.get("bsc", {}).get("scale", 1.0),
                "x": self.settings.get("bsc", {}).get("x", -1),
                "y": self.settings.get("bsc", {}).get("y", -1)
            },
            "nsc": {
                "wb": self.excel_reader_nsc.wb_name,
                "sheet": self.excel_reader_nsc.sheet_name,
                "c1": self.excel_reader_nsc.cells[0],
                "c2": self.excel_reader_nsc.cells[1],
                "c3": self.excel_reader_nsc.cells[2],
                "log_enabled": self.settings.get("nsc", {}).get("log_enabled", False),
                "log_interval": self.settings.get("nsc", {}).get("log_interval", "1 minute"),
                "log_source": self.settings.get("nsc", {}).get("log_source", 0),
                "scale": self.settings.get("nsc", {}).get("scale", 1.0),
                "x": self.settings.get("nsc", {}).get("x", -1),
                "y": self.settings.get("nsc", {}).get("y", -1)
            },
            "general": self.settings.get("general", {})
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

        action_overlay = QAction("Overlay Visual Settings...", menu)
        action_overlay.triggered.connect(self.open_overlay_settings)
        menu.addAction(action_overlay)

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

        # Start timer if overlay is visible OR logging is enabled for either
        log_bsc = self.settings.get("bsc", {}).get("log_enabled", False)
        log_nsc = self.settings.get("nsc", {}).get("log_enabled", False)

        if show_bsc or show_nsc or log_bsc or log_nsc:
            if not self.excel_timer.isActive():
                self.excel_timer.start(100)
        else:
            self.excel_timer.stop()

    def open_settings(self):
        dialog = SettingsDialog(
            self.excel_reader_bsc, self.excel_reader_nsc,
            self.settings.get("bsc", {}), self.settings.get("nsc", {}),
            self.settings.get("general", {})
        )
        if dialog.exec_() == QDialog.Accepted:
            # Update local settings object with new data from dialog
            if "bsc" not in self.settings: self.settings["bsc"] = {}
            if "nsc" not in self.settings: self.settings["nsc"] = {}
            if "general" not in self.settings: self.settings["general"] = {}
            
            self.settings["bsc"].update(dialog.bsc_data)
            self.settings["nsc"].update(dialog.nsc_data)
            self.settings["general"].update(dialog.general_data)
            
            self.save_settings()
            self.apply_settings() # Apply new adjustment instantly
            # Restart timer check in case logging was enabled/disabled
            self.toggle_premium_overlays()

    def open_overlay_settings(self):
        dialog = OverlaySettingsDialog(self.settings, self)
        dialog.exec_()
        # Settings are auto-saved during slider movement in this version

    def update_excel_data(self):
        # Update BSC
        vals_bsc = None
        if self.action_bsc.isChecked() or self.settings.get("bsc", {}).get("log_enabled"):
            vals_bsc = self.excel_reader_bsc.read_cells()
            if self.action_bsc.isChecked():
                self.premium_overlay_bsc.update_data(*vals_bsc)
            self.process_logging("BSC", vals_bsc, self.excel_reader_bsc)
        
        # Update NSC
        vals_nsc = None
        if self.action_nsc.isChecked() or self.settings.get("nsc", {}).get("log_enabled"):
            vals_nsc = self.excel_reader_nsc.read_cells()
            if self.action_nsc.isChecked():
                self.premium_overlay_nsc.update_data(*vals_nsc)
            self.process_logging("NSC", vals_nsc, self.excel_reader_nsc)

    def process_logging(self, type_key, vals, reader):
        config = self.settings.get(type_key.lower(), {})
        if not config.get("log_enabled"):
            return

        # Time handling
        current_ts, _, _ = self.syncer.get_current_time()
        from datetime import datetime, timezone, timedelta
        ist_tz = timezone(timedelta(hours=5, minutes=30))
        dt = datetime.fromtimestamp(current_ts, tz=ist_tz)

        # 1. Market Hours Check (9:15 AM - 3:30 PM)
        market_start = dt.replace(hour=9, minute=15, second=0, microsecond=0)
        market_end = dt.replace(hour=15, minute=30, second=0, microsecond=0)
        
        if not (market_start <= dt <= market_end):
            return

        # 2. Aligned Interval Ticking (Quantized)
        interval_str = config.get("log_interval", "1 minute")
        interval_map = {
            "1 second": 1,
            "5 seconds": 5,
            "30 seconds": 30,
            "1 minute": 60,
            "5 minutes": 300
        }
        interval_sec = interval_map.get(interval_str, 60)
        
        # Bucket is current_ts floor-divided by interval
        bucket_id = int(current_ts // interval_sec)
        state = self.log_state[type_key]
        
        if "last_bucket" not in state: state["last_bucket"] = 0

        if bucket_id > state["last_bucket"]:
            # Time to log
            source_idx = config.get("log_source", 0)
            try:
                # Ensure we have data
                if not vals or len(vals) <= source_idx: return
                current_val_raw = vals[source_idx]
                if current_val_raw == "" or current_val_raw is None: return
                current_val = float(current_val_raw)
            except (ValueError, TypeError):
                return # Skip if data is not valid number

            prev_val = state["last_val"]
            diff = 0
            if prev_val is not None:
                diff = current_val - prev_val
            
            # Update state
            state["last_bucket"] = bucket_id
            state["last_val"] = current_val

            # Prepare row data
            sheet_name = dt.strftime('%Y-%m-%d')
            time_str = dt.strftime('%I:%M:%S %p')
            weekday = dt.strftime('%A')
            
            # Row match req: table want datetime, premium, diffrence
            row_data = [time_str, current_val, diff]
            # Header match req: metadata Row 1, table header Row 3
            metadata = [f"Date: {sheet_name}", f"Day: {weekday}", f"Type: {type_key}"]
            table_header = ["DateTime", "Premium", "Difference"]
            
            reader.append_row(sheet_name, row_data, table_header, metadata)

    def exit_app(self):
        self.tray_icon.hide()
        self.app.quit()

    def run(self):
        sys.exit(self.app.exec_())

if __name__ == "__main__":
    app = OdinOverlayApp()
    app.run()
