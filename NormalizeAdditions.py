'''
Clean a propane-tank "additions" file and review it side by side.

Open an Excel workbook (preferred) or a CSV, normalize every column against the
rules the downstream system expects, and save the result as a CSV. The window
shows the original and the cleaned data next to each other so the operator can
see what changed and fix anything the rules could not resolve.

Excel is the better input: a CSV only carries the text a cell *displays*, so a
date shown as "Dec-95" has already lost the real 12/1/1995 behind it. Reading
the workbook hands us the underlying value instead of a formatting guess.

Cell colors
    RED     required field is empty or invalid — must be fixed
    ORANGE  cleaned, but a human should confirm it
    GREEN   cleaned side: the value was changed and passed validation
    BLUE    original side: this cell was changed by cleaning

Layout: the cleaning and validation rules are plain functions over strings with
no Qt in sight, so they can be exercised headless. `MainWindow` only displays
what `normalize_dataframe` decided.

Author: Timothy Pactwa
Version: 7/31/2026
'''

import re
import sys
from datetime import date, datetime
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTableWidget, QTableWidgetItem, QFileDialog,
    QHeaderView, QSplitter, QMessageBox,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

RED = QColor("#ffcccc")      # hard error: required field is empty or invalid
GREEN = QColor("#ccffcc")    # cleaned side: value was normalized and passed validation
ORANGE = QColor("#ffd9a0")   # warning: needs a human to review
BLUE = QColor("#d0e4ff")     # original side: cell was changed by cleaning

MANUFACTURER_MAX_LEN = 12
SIZE_GALLONS_MIN = 57
SIZE_GALLONS_MAX = 5000

# Excel counts days from here (serial 0); the offset absorbs its 1900 leap bug
EXCEL_EPOCH = date(1899, 12, 30)

# Batches are stamped in Eastern wall-clock time so the name matches the clock
# in the office no matter what zone the machine is set to. "America/New_York"
# rather than a fixed -5 because it has to be EDT in summer and EST in winter.
# Falls back to the machine clock if the zone database isn't available.
try:
    from zoneinfo import ZoneInfo
    EASTERN = ZoneInfo("America/New_York")
except Exception:  # no tzdata on this machine
    EASTERN = None

# Excel writes month names into CSVs whenever the cell is formatted that way
MONTHS = {
    "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04", "MAY": "05", "JUN": "06",
    "JUL": "07", "AUG": "08", "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12",
}

# long manufacturer names -> approved short codes (<=12 chars); misses stay as-is and flag orange
MANUFACTURERS = {
    "QUALITY STEEL": "QUALITY STL",
    "AMERICAN WELDING AND TANK": "AWT",
    "NATIONAL BUTANE": "NAT BUTANE",
    "AMERICAN": "AWT",
}


# ---------------------------------------------------------------------------
# Cleaning / validation logic (kept free of Qt so it can be tested headless)
# ---------------------------------------------------------------------------

def _is_blank(v):
    """True for None, NaN, an empty cell, or a string of only whitespace."""
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return str(v).strip() == ""


def _is_numeric(v):
    """True if the value reads as a number once commas are ignored."""
    if _is_blank(v):
        return False
    try:
        float(str(v).replace(",", "").strip())
        return True
    except (ValueError, TypeError):
        return False


def _as_text(v):
    """Every cleaner's starting point: a trimmed string, blank for empty."""
    return "" if _is_blank(v) else str(v).strip()


# ---------------------------------------------------------------------------
# Per-column cleaners. Each takes one cell and returns the cleaned string.
# ---------------------------------------------------------------------------

def _normalize_serial(v):
    # kept verbatim as text so values like "12E45" never become 1.2e+46
    return _as_text(v)


def _normalize_customer_owned(v):
    # anything starting with y/n answers the question; the rest flags later
    s = _as_text(v)
    if s.lower().startswith("y"):
        return "Y"
    if s.lower().startswith("n"):
        return "N"
    return s


def _normalize_manufacturer(v):
    if _is_blank(v):
        return "UNKNOWN"
    m = _as_text(v).upper()
    # "?", "???", "N/A", "N.A." are all somebody writing "we don't know"
    if "?" in m or re.sub(r"[^A-Z]", "", m) == "NA":
        return "UNKNOWN"
    # the shortcut table wins at any length, so short aliases like "AMERICAN"
    # map too; a long name with no shortcut stays long and flags orange later
    return MANUFACTURERS.get(m, m)


def _normalize_type(v):
    # only ASME / DOT allowed; any mix of letters resolves on the first A or D
    s = _as_text(v)
    letters = re.sub(r"[^A-Za-z]", "", s).upper()
    if letters[:1] == "A":
        return "ASME"
    if letters[:1] == "D":
        return "DOT"
    return s  # unrecognized -> flagged red later


def _strip_commas(v):
    # commas and stray spaces are formatting, not data: "1,234" -> "1234"
    return re.sub(r"[,\s]", "", _as_text(v))


def _normalize_size(v):
    # no commas, no decimals: "1,000.0" -> "1000". The fraction is dropped, not
    # deleted, so the magnitude stays right ("100.5" -> "100", never "1005").
    # Range is checked in the flag pass.
    s = _strip_commas(v)
    frac = re.match(r"^(\d+)\.\d+$", s)
    return frac.group(1) if frac else s


# Which cleaner runs on which column. Add a column by adding a line here.
# ManufacturerDate is missing on purpose: it needs the row's Type, so it is
# handled separately in normalize_dataframe.
COLUMN_CLEANERS = {
    "Serial": _normalize_serial,
    "CSC": _strip_commas,
    "Serv loc": _strip_commas,
    "Customer Owned": _normalize_customer_owned,
    "Manufacturer": _normalize_manufacturer,
    "Type": _normalize_type,
    "SizeGallons": _normalize_size,
    "SizePounds": _normalize_size,
}

# Stamped onto every row regardless of what the file said, because these mark
# the batch rather than describe the tank. "Activity date" is stamped with
# today's date separately.
CONSTANT_COLUMNS = {
    "Activity": "1",
    "Source": "Miscellaneous",
}


# ---------------------------------------------------------------------------
# ManufacturerDate gets its own section: a date reaches us in a dozen shapes,
# and a round trip through a spreadsheet mangles several of them.
# ---------------------------------------------------------------------------

def _date_default(tank_type):
    # nothing usable in the cell — stamp the placeholder the system expects
    if tank_type == "ASME":
        return "1900"
    if tank_type == "DOT":
        return "12 60"
    return ""


def _expand_year(y):
    # two-digit years pivot at 50: "05" -> 2005, "95" -> 1995
    if len(y) == 4:
        return y
    if len(y) <= 2:
        y = y.zfill(2)
        return "20" + y if int(y) < 50 else "19" + y
    return None  # three digits is not a year we can trust


def _undo_number_formatting(s):
    """Strip the artifacts Excel and pandas leave on a date before parsing it.

    None of these are real date syntax — they are what a year looks like after a
    round trip through a spreadsheet cell or a float column.
    """
    s = s.lstrip("'")  # Excel's "treat this as text" marker
    # pandas/Excel may hand us "2005.0" — collapse to a plain integer year. Only
    # values that actually carry a decimal, so "05" keeps its leading zero
    if "." in s or "e" in s.lower():
        try:
            n = float(s)
            if n == int(n):
                s = str(int(n))
        except ValueError:
            pass
    # a fraction on something too long to be a month ("2005.5") is noise rather
    # than a "MM.YY" separator — drop it so no year carries a decimal
    s = re.sub(r"^(\d{3,})\.\d+$", r"\1", s)
    # Excel stores a typed "12.60" as the number 12.6, so a lone fractional
    # digit is a dropped trailing zero, not a one-digit year
    return re.sub(r"^(\d{1,2})\.(\d)$", r"\g<1>.\g<2>0", s)


def _parse_month_year(s):
    """Pull ("MM", "YYYY", day) out of a date string, or None if unreadable.

    Excel writes the cell's *displayed* text into a CSV, never the underlying
    date, so the same tank can arrive as "Dec-95", "12/1/1995", "1995-12-01" or
    "12 95". Parse the pieces rather than guessing at one fixed layout. The day
    is not part of the output but is returned so the caller can spot Excel's
    date-serial mangling; it is None when the cell carried no day.
    """
    tokens = [t for t in re.split(r"[\s\-/.,]+", s.strip()) if t]
    named_month = None
    nums = []
    for t in tokens:
        if t.isdigit():
            nums.append(t)
        elif named_month is None and t.isalpha() and t[:3].upper() in MONTHS:
            named_month = MONTHS[t[:3].upper()]
        else:
            return None  # a word we don't recognize -> not a date

    # "Dec-95", "1-Dec-95", "Dec 1, 1995": the last number is always the year
    if named_month is not None:
        if not nums:
            return None
        day = int(nums[0]) if len(nums) > 1 else None
        year4 = _expand_year(nums[-1])
        return None if year4 is None else (named_month, year4, day)

    day = None
    if len(nums) == 3:
        if len(nums[0]) == 4:
            month, year, day = nums[1], nums[0], int(nums[2])   # ISO 1995-12-01
        elif int(nums[0]) <= 12:
            month, year, day = nums[0], nums[2], int(nums[1])   # US 12/1/1995
        else:
            month, year, day = nums[1], nums[2], int(nums[0])   # 25/12/1995 D/M/Y
    elif len(nums) == 2:
        month, year = nums                          # "12 95", "06/12", "3.05"
    elif len(nums) == 1 and len(nums[0]) in (2, 4):
        month, year = "01", nums[0]                 # year only
    else:
        return None

    if len(month) > 2 or not 1 <= int(month) <= 12:
        return None
    year4 = _expand_year(year)
    return None if year4 is None else (month.zfill(2), year4, day)


def _recover_serial_year(month, year4, day):
    """Undo Excel eating a typed 4-digit year, or None if that isn't what this is.

    Typing "2020" into a date-formatted cell makes Excel read it as day-serial
    2020 and store 1905-07-12. Every year from 1900 to 2100 lands somewhere in
    1905 and nothing else does, so a 1905 date converted back to its serial is
    the year the user meant. Same trick as Format Cells -> Number in Excel.
    """
    if year4 != "1905" or day is None:
        return None
    try:
        serial = (date(1905, int(month), day) - EXCEL_EPOCH).days
    except ValueError:  # impossible day for the month
        return None
    return str(serial) if 1900 <= serial <= 2100 else None


def _normalize_date(v, tank_type):
    """Clean one date -> (value, review).

    Usable input becomes a 4-digit year for ASME and "MM YY" for DOT. Anything
    else takes the type default. `review` asks the flag pass for an orange: the
    cell was filled but unreadable, or it was a year Excel had mangled into a
    date serial and we put it back.
    """
    if _is_blank(v):
        return _date_default(tank_type), False
    parsed = _parse_month_year(_undo_number_formatting(_as_text(v)))
    if parsed is None:
        return _date_default(tank_type), True  # "UNKNOWN", "N/A", junk
    month, year4, day = parsed
    # a typed year Excel turned into a date serial: put the year back, but flag
    # it so someone confirms the repair rather than trusting it blindly
    review = False
    recovered = _recover_serial_year(month, year4, day)
    if recovered is not None:
        month, year4, review = "01", recovered, True
    if tank_type == "DOT":
        return f"{month} {year4[2:]}", review  # month is zero-padded: "03 05"
    return year4, review  # ASME and fallback: 4-digit year only


# ---------------------------------------------------------------------------
# Loading a file
# ---------------------------------------------------------------------------

def _cell_to_text(v):
    """Flatten one Excel cell to the text the cleaning rules expect.

    Reading the workbook instead of a CSV export is the whole point: a date cell
    arrives as a real datetime, so "Dec-95" and "12/1/1995" are the same value
    here and the display format never has to be guessed.
    """
    if v is None:
        return ""
    try:
        if pd.isna(v):  # empty cell, NaT
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, date):  # covers datetime and pandas Timestamp
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))  # Excel hands back 500.0 for a plain 500
    return str(v).strip()


def batch_filename(now=None):
    """Default name for a saved batch: AMS_Batch_MMDDYYYY_HHMM.csv.

    Stamped at save time, in Eastern time, on a 24-hour clock — so a batch saved
    at 1:45pm on July 31 2026 becomes AMS_Batch_07312026_1345.csv.
    """
    return (now or datetime.now(EASTERN)).strftime("AMS_Batch_%m%d%Y_%H%M.csv")


def read_table(path):
    """Load a CSV or Excel workbook into an all-text DataFrame."""
    if path.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xls")):
        # keep_default_na=False so a literal "N/A" survives as text and still
        # trips the flag, matching how the CSV path sees it
        df = pd.read_excel(path, sheet_name=0, dtype=object, keep_default_na=False)
        for c in df.columns:
            df[c] = df[c].apply(_cell_to_text)
        return df
    # read everything as text so nothing (Serial, integers) gets float-coerced
    return pd.read_csv(path, dtype=str, keep_default_na=False)


# ---------------------------------------------------------------------------
# The whole-file pass: clean every column, then judge what came out
# ---------------------------------------------------------------------------

def _bad_size_gallons(v):
    # must be a whole number of gallons inside the tank sizes we stock
    s = str(v)
    return not (s.isdigit() and SIZE_GALLONS_MIN <= int(s) <= SIZE_GALLONS_MAX)


def _not_a_number(v):
    return _is_blank(v) or not _is_numeric(v)


def _flag_rows(flags, column, values, is_bad, color):
    """Color every row of `column` whose value fails `is_bad`.

    `values` is None when the file has no such column, in which case there is
    nothing to judge and no flag is raised.
    """
    if values is None:
        return
    for i, v in enumerate(values):
        if is_bad(v):
            flags[(i, column)] = color


def normalize_dataframe(df):
    """Clean `df` in place and return {(row_index, column_name): 'RED'|'ORANGE'}.

    Columns the file does not have are skipped, never invented — a missing
    column raises no flags. Cleaning runs first so validation always judges the
    cleaned value, not what the operator originally typed.
    """
    cols = df.columns

    for column, clean in COLUMN_CLEANERS.items():
        if column in cols:
            df[column] = df[column].apply(clean)

    # the date needs its row's Type, so it can't go in the table above; it also
    # reports which rows want review (unreadable, or recovered from a serial)
    date_review = [False] * len(df)
    if "ManufacturerDate" in cols:
        types = df["Type"] if "Type" in cols else [""] * len(df)
        cleaned = [_normalize_date(d, t) for d, t in zip(df["ManufacturerDate"], types)]
        df["ManufacturerDate"] = [value for value, _ in cleaned]
        date_review = [review for _, review in cleaned]

    # batch markers for new additions: same value stamped on every row
    for column, marker in CONSTANT_COLUMNS.items():
        if column in cols:
            df[column] = marker
    if "Activity date" in cols:
        df["Activity date"] = pd.Timestamp.today().strftime("%m/%d/%Y")

    # ---- validation pass over the cleaned values ----
    flags = {}

    def column(name):
        return df[name].tolist() if name in cols else None

    typ = column("Type")
    mfr = column("Manufacturer")
    sizep = column("SizePounds")

    _flag_rows(flags, "CSC", column("CSC"), _not_a_number, "RED")
    _flag_rows(flags, "Serv loc", column("Serv loc"), _not_a_number, "ORANGE")
    _flag_rows(flags, "Equipment Type", column("Equipment Type"),
               lambda v: str(v).strip() != "1", "RED")
    _flag_rows(flags, "Type", typ, lambda v: v not in ("ASME", "DOT"), "RED")
    _flag_rows(flags, "SizeGallons", column("SizeGallons"), _bad_size_gallons, "RED")
    # anything still over the length cap had no shortcut in MANUFACTURERS
    _flag_rows(flags, "Manufacturer", mfr,
               lambda v: len(str(v)) > MANUFACTURER_MAX_LEN, "ORANGE")
    # unreadable and stamped with the type default, or a year recovered from
    # Excel's date-serial mangling — either way, worth a human look
    _flag_rows(flags, "ManufacturerDate", date_review, bool, "ORANGE")

    # SizePounds is the one rule that reads other columns, so it gets its own
    # loop: pounds are required for DOT tanks and for anything Worthington made
    if sizep is not None:
        for i, v in enumerate(sizep):
            required = (
                (typ is not None and typ[i] == "DOT")
                or (mfr is not None and str(mfr[i]).strip().upper() == "WORTHINGTON")
            )
            if required and _is_blank(v):
                flags[(i, "SizePounds")] = "ORANGE"

    return flags


# ---------------------------------------------------------------------------
# The window. Displays what the rules above decided; holds no rules of its own.
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cleaned Additions")
        self.resize(1400, 700)

        self.cell_flags = {}

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        top_bar = QHBoxLayout()

        # Open Button — Excel is the preferred input (real dates, no format guessing)
        open_btn = QPushButton("Open Excel / CSV")
        open_btn.clicked.connect(self.load_file)

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

    # load the additions file, run cleaning, then fill both tables
    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Additions File", "",
            "Excel or CSV (*.xlsx *.xlsm *.xls *.csv);;Excel Workbook (*.xlsx *.xlsm *.xls);;CSV Files (*.csv)",
        )
        if not path:
            return

        try:
            self.original_df = read_table(path)
        except Exception as exc:  # unreadable file, wrong sheet, locked by Excel
            QMessageBox.critical(self, "Could not open file", f"{path}\n\n{exc}")
            return

        self.cleaned_df = self.original_df.copy()
        self._normalize(self.cleaned_df)
        self._populate_tables()
        self.save_btn.setEnabled(True)

    # main logic for data cleaning — stores per-cell validation flags for display
    def _normalize(self, df):
        self.cell_flags = normalize_dataframe(df)

    def _cell_str(self, df, r, c):
        """One cell as display text. Deliberately not `_as_text`: the original
        side must show exactly what was in the file, untrimmed, or the
        changed-by-cleaning comparison would miss whitespace-only edits."""
        val = df.iloc[r, c]
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ""
        return str(val)

    @staticmethod
    def _make_item(text, flag, changed, changed_color, editable):
        """Build one table cell. A validation flag always outranks the
        was-changed tint, so a problem is never hidden by a green cell."""
        item = QTableWidgetItem(text)
        if not editable:
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        if flag == "RED":
            item.setBackground(RED)
        elif flag == "ORANGE":
            item.setBackground(ORANGE)
        elif changed:
            item.setBackground(changed_color)
        return item

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

        for r in range(rows):
            for c in range(cols):
                col = headers[c]
                orig_val = self._cell_str(orig, r, c)
                clean_val = self._cell_str(cleaned, r, c)
                changed = orig_val != clean_val
                flag = self.cell_flags.get((r, col))

                # original side is read-only and tints changed cells blue; the
                # cleaned side is editable and tints them green
                self.original_table.setItem(r, c, self._make_item(
                    orig_val, flag, changed, BLUE, editable=False))
                self.cleaned_table.setItem(r, c, self._make_item(
                    clean_val, flag, changed, GREEN, editable=True))

    # read back from the live table so any manual edits in the UI are captured
    def save_csv(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Cleaned CSV", batch_filename(), "CSV Files (*.csv)"
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
