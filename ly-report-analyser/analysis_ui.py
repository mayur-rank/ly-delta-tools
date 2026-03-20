import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QProgressBar, QMessageBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor, QFont
from pattern_analyzer import PremiumPatternAnalyzer

class AnalysisThread(QThread):
    finished = pyqtSignal(object, float)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def run(self):
        try:
            analyzer = PremiumPatternAnalyzer(self.folder_path)
            self.progress.emit("Scanning files and loading data...")
            if not analyzer.load_data():
                self.error.emit("No valid Excel data found in the selected folder.")
                return
            
            self.progress.emit("Mining data for patterns...")
            patterns, threshold = analyzer.find_patterns()
            
            if patterns is None or patterns.empty:
                self.error.emit("No recurring patterns found in the data.")
                return
                
            self.finished.emit(patterns, threshold)
        except Exception as e:
            self.error.emit(str(e))

class AnalysisDashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Odin Premium Pattern Analyzer")
        self.setMinimumSize(900, 600)
        self.initUI()
        self.apply_styles()

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Top Bar: Folder Selection
        top_bar = QHBoxLayout()
        self.folder_label = QLabel("Select folder containing your Excel logs...")
        self.folder_label.setStyleSheet("color: #AAAAAA; font-style: italic;")
        
        browse_btn = QPushButton("Browse Folder")
        browse_btn.clicked.connect(self.browse_folder)
        
        self.run_btn = QPushButton("Run Analysis")
        self.run_btn.setEnabled(False)
        self.run_btn.clicked.connect(self.start_analysis)
        
        top_bar.addWidget(self.folder_label, 1)
        top_bar.addWidget(browse_btn)
        top_bar.addWidget(self.run_btn)
        layout.addLayout(top_bar)

        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        # Pattern Table
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["Time", "Occurrences", "Avg Drop", "Max Drop", "Avg Premium", "Days Spotted"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table)

        # Bottom Bar: Summary & Export
        bottom_bar = QHBoxLayout()
        self.threshold_lbl = QLabel("Threshold: N/A")
        export_btn = QPushButton("Export Master Report")
        export_btn.clicked.connect(self.export_report)
        
        bottom_bar.addWidget(self.threshold_lbl)
        bottom_bar.addStretch()
        bottom_bar.addWidget(export_btn)
        layout.addLayout(bottom_bar)

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #1E1E2E;
                color: #CDD6F4;
                font-family: 'Segoe UI', Arial;
            }
            QPushButton {
                background-color: #45475A;
                border: 1px solid #585B70;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #585B70;
            }
            QPushButton:enabled {
                color: #89B4FA;
            }
            QPushButton:disabled {
                color: #6C7086;
            }
            QTableWidget {
                background-color: #181825;
                gridline-color: #313244;
                border: 1px solid #313244;
                border-radius: 8px;
            }
            QHeaderView::section {
                background-color: #313244;
                padding: 5px;
                border: 1px solid #181825;
                font-weight: bold;
            }
            QLabel {
                font-size: 14px;
            }
        """)

    def browse_folder(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Logs Folder")
        if dir_path:
            self.selected_folder = dir_path
            self.folder_label.setText(dir_path)
            self.run_btn.setEnabled(True)

    def start_analysis(self):
        self.run_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0) # Pulsing
        self.table.setRowCount(0)
        
        self.thread = AnalysisThread(self.selected_folder)
        self.thread.progress.connect(self.status_label.setText)
        self.thread.finished.connect(self.display_results)
        self.thread.error.connect(self.handle_error)
        self.thread.start()

    def display_results(self, patterns, threshold):
        self.progress_bar.setVisible(False)
        self.run_btn.setEnabled(True)
        self.status_label.setText("Analysis Complete!")
        self.threshold_lbl.setText(f"Outlier Threshold: {threshold:.4f}")
        self.patterns = patterns
        
        self.table.setRowCount(len(patterns))
        for i, (idx, row) in enumerate(patterns.iterrows()):
            self.table.setItem(i, 0, QTableWidgetItem(str(row['Time'])))
            self.table.setItem(i, 1, QTableWidgetItem(str(row['Occurrences'])))
            self.table.setItem(i, 2, QTableWidgetItem(f"{row['Avg_Drop']:.2f}"))
            self.table.setItem(i, 3, QTableWidgetItem(f"{row['Max_Drop']:.2f}"))
            self.table.setItem(i, 4, QTableWidgetItem(f"{row['Avg_Premium']:.2f}"))
            self.table.setItem(i, 5, QTableWidgetItem(str(row['Weekdays'])))
            
            # Color coding for "Strong" patterns
            if int(row['Occurrences']) > 2:
                for col in range(6):
                    item = self.table.item(i, col)
                    if item: item.setForeground(QColor("#A6E3A1")) # Green for frequency

    def handle_error(self, msg):
        self.progress_bar.setVisible(False)
        self.run_btn.setEnabled(True)
        QMessageBox.critical(self, "Analysis Error", msg)

    def export_report(self):
        if hasattr(self, 'selected_folder'):
            analyzer = PremiumPatternAnalyzer(self.selected_folder)
            analyzer.load_data() # Re-load or cache
            analyzer.generate_report()
            QMessageBox.information(self, "Export Success", "Pattern_Report.xlsx has been generated in the project folder.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AnalysisDashboard()
    window.show()
    sys.exit(app.exec_())
