'''
Author: Timothy Pactwa
Version: 6/3/2026
'''

import sys
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTableWidget, QTableWidgetItem, QFileDialog,
    QHeaderView, QSplitter, QMessageBox,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

RED = QColor("#ffcccc")
GREEN = QColor("#ccffcc")
ORANGE = QColor("#ffd9a0")

MANUFACTURER_MAX_LEN = 12

# UI inherits from Qt6's Main Window
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cleaned Additions")
        self.resize(1400, 700)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        top_bar = QHBoxLayout()

        # Open Button
        open_btn = QPushButton("Open CSV")
        open_btn.clicked.connect(self.load_csv)

        # Save Button
        self.save_btn = QPushButton("Save Cleaned CSV")
        self.save_btn.clicked.connect(self.save_csv)
        self.save_btn.setEnabled(False)

        # add buttons to top bar then add the top bar to the UI
        top_bar.addWidget(open_btn)
        top_bar.addWidget(self.save_btn)
        top_bar.addStretch()
        layout.addLayout(top_bar)

        # split screen for side by side view
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self._maximized_panel = None

        # Original CSV display
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_header = QHBoxLayout()
        left_header.addWidget(QLabel("Original"))
        left_header.addStretch()
        self.expand_left_btn = QPushButton("Max")
        self.expand_left_btn.setToolTip("Maximize Original")
        self.expand_left_btn.clicked.connect(lambda: self._toggle_maximize(0))
        left_header.addWidget(self.expand_left_btn)
        left_layout.addLayout(left_header)
        self.original_table = QTableWidget()
        self.original_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        left_layout.addWidget(self.original_table)

        # Cleaned CSV display
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_header = QHBoxLayout()
        right_header.addWidget(QLabel("Cleaned"))
        right_header.addStretch()
        self.expand_right_btn = QPushButton("Max")
        self.expand_right_btn.setToolTip("Maximize Cleaned")
        self.expand_right_btn.clicked.connect(lambda: self._toggle_maximize(1))
        right_header.addWidget(self.expand_right_btn)
        right_layout.addLayout(right_header)
        self.cleaned_table = QTableWidget()
        right_layout.addWidget(self.cleaned_table)

        self.splitter.addWidget(left_widget)
        self.splitter.addWidget(right_widget)
        layout.addWidget(self.splitter)

    # collapse one side to zero width; calling again restores the split
    def _toggle_maximize(self, panel):
        total = sum(self.splitter.sizes())
        if self._maximized_panel == panel:
            self.splitter.setSizes([total // 2, total // 2])
            self._maximized_panel = None
            self.expand_left_btn.setText("Max")
            self.expand_right_btn.setText("Max")
        else:
            self.splitter.setSizes([0, total] if panel == 1 else [total, 0])
            self._maximized_panel = panel
            self.expand_left_btn.setText("Restore" if panel == 0 else "Max")
            self.expand_right_btn.setText("Restore" if panel == 1 else "Max")

    # load CSV, run cleaning, then fill both tables
    def load_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open CSV", "", "CSV Files (*.csv)")
        if not path:
            return

        self.original_df = pd.read_csv(path)
        self.cleaned_df = self.original_df.copy()
        self._normalize(self.cleaned_df)
        self._populate_tables()
        self.save_btn.setEnabled(True)

    # main logic for data cleaning
    def _normalize(self, df):
        # Yes/No normalization
        responses_dict = {
            "YES": "Y", "Yes": "Y", "yes": "Y", "y": "Y",
            "NO": "N",  "No": "N",  "no": "N",  "n": "N",
        }
        if "Customer Owned" in df.columns:
            df["Customer Owned"] = df["Customer Owned"].replace(responses_dict)

        # dict hits get normalized; misses are left as-is and flagged orange in the table
        manufacturers_dict = {
            "QUALITY STEEL": "QUALITY STL",
            "AMERICAN WELDING AND TANK": "AWT",
        }
        def normalize_manufacturer(manufacturer):
            if pd.isna(manufacturer):
                return "UNKNOWN"
            if len(str(manufacturer)) <= MANUFACTURER_MAX_LEN:
                return manufacturer
            if manufacturer in manufacturers_dict:
                return manufacturers_dict[manufacturer]
            return manufacturer
        if "Manufacturer" in df.columns:
            df["Manufacturer"] = df["Manufacturer"].apply(normalize_manufacturer)

        # fill empty dates with default
        default_year = 1900
        if "ManufacturerDate" in df.columns:
            df["ManufacturerDate"] = df["ManufacturerDate"].fillna(default_year)

    # pandas reads integer columns with any NaN as float64, so 1 becomes 1.0 — convert back
    def _cell_str(self, df, r, c):
        val = df.iloc[r, c]
        if pd.isna(val):
            return ""
        if isinstance(val, float) and val.is_integer():
            return str(int(val))
        return str(val)

    def _populate_tables(self):
        orig = self.original_df
        cleaned = self.cleaned_df
        rows, cols = orig.shape
        headers = list(orig.columns)

        # size both tables and stretch columns to fill available width
        for table in (self.original_table, self.cleaned_table):
            table.setRowCount(rows)
            table.setColumnCount(cols)
            table.setHorizontalHeaderLabels(headers)
            table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        mfr_col = headers.index("Manufacturer") if "Manufacturer" in headers else -1
        equipment_type_col = headers.index("Equipment Type") if "Equipment Type" in headers else -1

        for r in range(rows):
            for c in range(cols):
                orig_val = self._cell_str(orig, r, c)
                clean_val = self._cell_str(cleaned, r, c)
                changed = orig_val != clean_val

                # long manufacturer not found in lookup dict — flag orange on both sides
                is_mfr_miss = (c == mfr_col and not changed and len(orig_val) > MANUFACTURER_MAX_LEN)
                # equipment type must be 1 — flag red in cleaned output
                is_invalid_equipment_type = (c == equipment_type_col and clean_val != "1")

                # original side: read-only; red if a change was made, orange if manufacturer needs review
                # strip the editable flag via bitwise AND with its complement
                orig_item = QTableWidgetItem(orig_val)
                orig_item.setFlags(orig_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if changed:
                    orig_item.setBackground(RED)
                elif is_mfr_miss:
                    orig_item.setBackground(ORANGE)
                self.original_table.setItem(r, c, orig_item)

                # cleaned side: green if normalized, orange if manufacturer needs review
                clean_item = QTableWidgetItem(clean_val)
                if changed:
                    clean_item.setBackground(GREEN)
                elif is_mfr_miss:
                    clean_item.setBackground(ORANGE)
                self.cleaned_table.setItem(r, c, clean_item)

                # equipment type check overrides other colors
                if is_invalid_equipment_type:
                    clean_item.setBackground(RED)

    # read back from the live table so any manual edits in the UI are captured
    def save_csv(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Cleaned CSV", "cleaned_additions.csv", "CSV Files (*.csv)"
        )
        if not path:
            return

        rows = self.cleaned_table.rowCount()
        cols = self.cleaned_table.columnCount()
        headers = [self.cleaned_table.horizontalHeaderItem(c).text() for c in range(cols)]
        data = [
            [self.cleaned_table.item(r, c).text() if self.cleaned_table.item(r, c) else "" for c in range(cols)]
            for r in range(rows)
        ]
        pd.DataFrame(data, columns=headers).to_csv(path, index=False)
        QMessageBox.information(self, "Saved", f"Saved to {path}")


LIGHT_THEME = """
    QMainWindow, QWidget {
        background-color: #f5f5f5;
        color: #000000;
    }
    QTableWidget {
        background-color: #ffffff;
        color: #000000;
        gridline-color: #cccccc;
        border: 1px solid #aaaaaa;
    }
    QHeaderView::section {
        background-color: #e8e8e8;
        color: #000000;
        border: 1px solid #aaaaaa;
        padding: 4px;
    }
    QTableCornerButton::section {
        background-color: #e8e8e8;
        border: 1px solid #aaaaaa;
    }
    QPushButton {
        background-color: #e0e0e0;
        color: #000000;
        border: 1px solid #aaaaaa;
        padding: 4px 10px;
        border-radius: 3px;
    }
    QPushButton:hover { background-color: #d0d0d0; }
    QPushButton:disabled { color: #999999; }
    QLabel { color: #000000; }
    QSplitter::handle { background-color: #cccccc; }
"""

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(LIGHT_THEME)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
