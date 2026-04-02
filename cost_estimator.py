# cost_estimator.py
"""
Cost Estimator tab — replicates the OC Lev 4-ASPCostReport Excel workflow.

Sections:
  1. Consulting Effort  — hours × hourly rate per labour category
  2. Third Party        — free-form description / cost rows
  3. Hardware / Software / Materials — qty × unit cost rows
  4. Travel & Expenses  — airfare / hotel / food / car
  5. Summary            — rollup → risk → margin → resale → profit
"""

import math
import json
import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QScrollArea, QFrame, QSizePolicy, QDoubleSpinBox, QAbstractItemView,
    QFileDialog, QMessageBox,
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QColor, QFont, QDoubleValidator, QIntValidator

# ── Shared colour constants ──────────────────────────────────────────────────
_RED      = "#920d2e"
_RED_DARK = "#7a0b27"
_TEXT     = "#1a0509"
_SUBTEXT  = "#3a3a5c"
_BG       = "#ffffff"
_BG_CARD  = "#ffffff"
_BG_HDR   = "#f9f0f2"
_BORDER   = "#d6c0c5"
_BLUE_BG  = "#e8f0fe"
_BLUE_FG  = "#1a3a6e"
_GRAY_BG  = "#f4f6fa"

_FIELD_STYLE = (
    "QLineEdit { color:#1a0509; background:#ffffff;"
    "  border:1px solid #d6c0c5; border-radius:4px;"
    "  padding:4px 8px; font-size:12px; }"
    "QLineEdit:focus { border-color:#920d2e; }"
)
_CALC_STYLE = (
    "QLineEdit { color:#1a3a6e; background:#e8f0fe;"
    "  border:1px solid #b8cce4; border-radius:4px;"
    "  padding:4px 8px; font-size:12px; font-weight:600; }"
)
_BTN_STYLE = (
    "QPushButton { background:#f4f6fa; color:#3a3a5c;"
    "  border:1px solid #dde1e7; border-radius:4px;"
    "  padding:4px 12px; font-size:11px; }"
    "QPushButton:hover { background:#e8ecf5; border-color:#b0b8d0; }"
    "QPushButton:pressed { background:#d8def0; }"
)
_DEL_BTN_STYLE = (
    "QPushButton { background:transparent; color:#e05252;"
    "  border:1px solid #e05252; border-radius:4px;"
    "  padding:4px 12px; font-size:11px; }"
    "QPushButton:hover { background:#fdf2f2; }"
)
_TABLE_STYLE = (
    "QTableWidget { background:#ffffff; border:1px solid #d6c0c5;"
    "  border-radius:4px; gridline-color:#e8dde0; font-size:11px; }"
    "QHeaderView::section { background:#f4f6fa; color:#3a3a5c;"
    "  font-size:11px; font-weight:600; border:none;"
    "  border-bottom:1px solid #d6c0c5; padding:4px 6px; }"
    "QTableWidget::item { padding:3px 6px; }"
    "QTableWidget::item:selected { background:#f5d0da; color:#920d2e; }"
)


# ── Helper: read a numeric QLineEdit ────────────────────────────────────────
def _num(widget: QLineEdit, default=0.0) -> float:
    try:
        return float(widget.text().replace(",", "").replace("$", "").strip() or default)
    except ValueError:
        return default


def _fmt_dollar(v: float) -> str:
    return f"${math.ceil(v):,}"


def _fmt_plain(v: float) -> str:
    if v == int(v):
        return f"{int(v):,}"
    return f"{v:,.2f}"


# ── Card widget (collapsible) ────────────────────────────────────────────────
class _Card(QFrame):
    """Collapsible section card matching the app's visual language."""

    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setStyleSheet(
            f"QFrame {{ background:{_BG_CARD}; border:1px solid {_BORDER};"
            "  border-radius:8px; }"
        )
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Header bar
        hdr = QWidget()
        hdr.setStyleSheet(
            f"QWidget {{ background:{_BG_HDR}; border-radius:7px 7px 0 0;"
            f"  border-bottom:1px solid {_BORDER}; }}"
        )
        hdr.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        hdr.setFixedHeight(40)
        hdr_row = QHBoxLayout(hdr)
        hdr_row.setContentsMargins(12, 0, 8, 0)
        hdr_row.setSpacing(8)

        self._toggle_btn = QPushButton("▼")
        self._toggle_btn.setFixedSize(22, 22)
        self._toggle_btn.setStyleSheet(
            f"QPushButton {{ background:transparent; border:none;"
            f"  color:{_RED}; font-size:12px; font-weight:700; padding:0; }}"
        )
        self._toggle_btn.clicked.connect(self._toggle)
        hdr_row.addWidget(self._toggle_btn)

        self._title_lbl = QLabel(title)
        self._title_lbl.setStyleSheet(
            f"QLabel {{ font-size:13px; font-weight:700; color:{_RED};"
            "  background:transparent; border:none; padding:0; }"
        )
        hdr_row.addWidget(self._title_lbl)
        hdr_row.addStretch()

        # slot for subtotal badge in header
        self._subtotal_lbl = QLabel("")
        self._subtotal_lbl.setStyleSheet(
            f"QLabel {{ font-size:12px; font-weight:700; color:{_RED};"
            "  background:transparent; border:none; padding:0 4px; }"
        )
        hdr_row.addWidget(self._subtotal_lbl)

        root.addWidget(hdr)

        # Body
        self._body = QWidget()
        self._body.setStyleSheet(
            "QWidget { background:transparent; border:none; }"
        )
        self._body_layout = QVBoxLayout(self._body)
        self._body_layout.setContentsMargins(14, 12, 14, 14)
        self._body_layout.setSpacing(8)
        root.addWidget(self._body)

        self._collapsed = False

    def _toggle(self):
        self._collapsed = not self._collapsed
        self._body.setVisible(not self._collapsed)
        self._toggle_btn.setText("▶" if self._collapsed else "▼")

    def set_subtotal(self, text: str):
        self._subtotal_lbl.setText(text)


# ── Numeric-only QLineEdit ───────────────────────────────────────────────────
class _NumEdit(QLineEdit):
    """QLineEdit that accepts decimal numbers and emits on change."""

    def __init__(self, placeholder="0", decimals=2, parent=None):
        super().__init__(parent)
        self.setPlaceholderText(placeholder)
        self.setStyleSheet(_FIELD_STYLE)
        self.setAlignment(Qt.AlignmentFlag.AlignRight)
        self._decimals = decimals

    def value(self) -> float:
        return _num(self)


# ── Calculated (read-only) display field ────────────────────────────────────
class _CalcEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setStyleSheet(_CALC_STYLE)
        self.setAlignment(Qt.AlignmentFlag.AlignRight)


# ══════════════════════════════════════════════════════════════════════════════
# Section 1 — Consulting Effort
# ══════════════════════════════════════════════════════════════════════════════
class _ConsultingCard(_Card):
    _ROWS = ["On-Site Time (hrs)", "Travel Time (hrs)", "Other (hrs)"]

    def __init__(self, on_change, parent=None):
        super().__init__("1.  Consulting Effort", parent)
        self._on_change = on_change

        # Hourly rate row
        rate_row = QHBoxLayout()
        rate_row.setSpacing(8)
        rate_row.addWidget(QLabel("Hourly Rate  ($):"))
        self._rate = _NumEdit("47.00")
        self._rate.setFixedWidth(120)
        self._rate.textChanged.connect(self._recalc)
        rate_row.addWidget(self._rate)
        rate_row.addStretch()
        self._body_layout.addLayout(rate_row)

        # Table: Labour category | Hours | Rate | Cost
        self._table = QTableWidget(len(self._ROWS), 4)
        self._table.setStyleSheet(_TABLE_STYLE)
        self._table.setHorizontalHeaderLabels(["Category", "Hours", "Rate ($/hr)", "Cost"])
        self._table.verticalHeader().setVisible(False)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        hh = self._table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for c in (1, 2, 3):
            hh.setSectionResizeMode(c, QHeaderView.ResizeMode.Fixed)
        self._table.setColumnWidth(1, 90)
        self._table.setColumnWidth(2, 110)
        self._table.setColumnWidth(3, 110)
        self._table.setFixedHeight(32 * len(self._ROWS) + 34)

        self._hour_edits = []
        self._cost_edits = []
        for r, lbl in enumerate(self._ROWS):
            # Category label
            cat = QTableWidgetItem(lbl)
            cat.setFlags(Qt.ItemFlag.ItemIsEnabled)
            cat.setForeground(QColor(_SUBTEXT))
            self._table.setItem(r, 0, cat)
            # Hours
            hrs = _NumEdit("0")
            hrs.textChanged.connect(self._recalc)
            self._table.setCellWidget(r, 1, hrs)
            self._hour_edits.append(hrs)
            # Rate (calculated)
            rate_cell = _CalcEdit()
            self._table.setCellWidget(r, 2, rate_cell)
            # Cost (calculated)
            cost_cell = _CalcEdit()
            self._table.setCellWidget(r, 3, cost_cell)
            self._cost_edits.append(cost_cell)

        self._body_layout.addWidget(self._table)

        # Subtotal row
        sub_row = QHBoxLayout()
        sub_row.addStretch()
        sub_row.addWidget(QLabel("Subtotal:"))
        self._subtotal_edit = _CalcEdit()
        self._subtotal_edit.setFixedWidth(130)
        sub_row.addWidget(self._subtotal_edit)
        self._body_layout.addLayout(sub_row)

        self._recalc()

    def _recalc(self):
        rate = _num(self._rate)
        total = 0.0
        for r in range(len(self._ROWS)):
            hrs = self._hour_edits[r].value()
            cost = hrs * rate
            total += cost
            # Update rate display
            rate_wid = self._table.cellWidget(r, 2)
            if isinstance(rate_wid, _CalcEdit):
                rate_wid.setText(f"${rate:,.2f}")
            self._cost_edits[r].setText(_fmt_dollar(cost))
        self._subtotal_edit.setText(_fmt_dollar(total))
        self.set_subtotal(_fmt_dollar(total))
        self._on_change()

    def subtotal(self) -> float:
        total = 0.0
        rate = _num(self._rate)
        for ed in self._hour_edits:
            total += ed.value() * rate
        return total

    def get_data(self) -> dict:
        return {
            "rate": self._rate.text(),
            "rows": [ed.text() for ed in self._hour_edits],
        }

    def restore_data(self, d: dict):
        self._rate.setText(d.get("rate", ""))
        for i, val in enumerate(d.get("rows", [])):
            if i < len(self._hour_edits):
                self._hour_edits[i].setText(val)


# ══════════════════════════════════════════════════════════════════════════════
# Section 2 — Third Party
# ══════════════════════════════════════════════════════════════════════════════
class _ThirdPartyCard(_Card):
    def __init__(self, on_change, parent=None):
        super().__init__("2.  Third Party", parent)
        self._on_change = on_change
        self._rows: list[tuple[QLineEdit, _NumEdit, _CalcEdit]] = []

        # Table
        self._table = QTableWidget(0, 3)
        self._table.setStyleSheet(_TABLE_STYLE)
        self._table.setHorizontalHeaderLabels(["Description", "Qty", "Cost Total"])
        self._table.verticalHeader().setVisible(False)
        hh = self._table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self._table.setColumnWidth(1, 80)
        self._table.setColumnWidth(2, 130)
        self._body_layout.addWidget(self._table)

        # Add / Remove buttons
        btn_row = QHBoxLayout()
        add_btn = QPushButton("＋  Add Row")
        add_btn.setStyleSheet(_BTN_STYLE)
        add_btn.clicked.connect(self.add_row)
        del_btn = QPushButton("－  Remove Selected")
        del_btn.setStyleSheet(_DEL_BTN_STYLE)
        del_btn.clicked.connect(self._remove_row)
        btn_row.addWidget(add_btn); btn_row.addWidget(del_btn); btn_row.addStretch()
        self._body_layout.addLayout(btn_row)

        # Subtotal
        sub_row = QHBoxLayout()
        sub_row.addStretch()
        sub_row.addWidget(QLabel("Subtotal:"))
        self._subtotal_edit = _CalcEdit()
        self._subtotal_edit.setFixedWidth(130)
        sub_row.addWidget(self._subtotal_edit)
        self._body_layout.addLayout(sub_row)

        self.add_row()

    def add_row(self):
        r = self._table.rowCount()
        self._table.insertRow(r)
        self._table.setRowHeight(r, 30)

        desc = QLineEdit()
        desc.setStyleSheet(_FIELD_STYLE)
        self._table.setCellWidget(r, 0, desc)

        qty = _NumEdit("1", decimals=0)
        qty.setFixedWidth(70)
        qty.textChanged.connect(self._recalc)
        self._table.setCellWidget(r, 1, qty)

        cost = _NumEdit("0.00")
        cost.textChanged.connect(self._recalc)
        self._table.setCellWidget(r, 2, cost)

        self._rows.append((desc, qty, cost))
        self._update_height()
        self._recalc()

    def _remove_row(self):
        rows = sorted({i.row() for i in self._table.selectedIndexes()}, reverse=True)
        for r in rows:
            self._table.removeRow(r)
            if r < len(self._rows):
                self._rows.pop(r)
        self._update_height()
        self._recalc()

    def _update_height(self):
        n = max(1, self._table.rowCount())
        self._table.setFixedHeight(30 * n + 34)

    def _recalc(self):
        total = 0.0
        for _, qty, cost in self._rows:
            total += qty.value() * cost.value()
        self._subtotal_edit.setText(_fmt_dollar(total))
        self.set_subtotal(_fmt_dollar(total))
        self._on_change()

    def subtotal(self) -> float:
        return sum(q.value() * c.value() for _, q, c in self._rows)

    def get_data(self) -> dict:
        return {"rows": [(d.text(), q.text(), c.text()) for d, q, c in self._rows]}

    def restore_data(self, d: dict):
        while self._table.rowCount():
            self._table.removeRow(0)
        self._rows.clear()
        for desc, qty, cost in d.get("rows", []):
            self.add_row()
            self._rows[-1][0].setText(desc)
            self._rows[-1][1].setText(qty)
            self._rows[-1][2].setText(cost)


# ══════════════════════════════════════════════════════════════════════════════
# Section 3 — Hardware / Software / Materials
# ══════════════════════════════════════════════════════════════════════════════
class _MaterialsCard(_Card):
    def __init__(self, on_change, parent=None):
        super().__init__("3.  Hardware / Software / Materials", parent)
        self._on_change = on_change
        self._rows: list[tuple[QLineEdit, _NumEdit, _NumEdit, _CalcEdit]] = []

        self._table = QTableWidget(0, 4)
        self._table.setStyleSheet(_TABLE_STYLE)
        self._table.setHorizontalHeaderLabels(["Description", "Qty", "Cost Per", "Cost Total"])
        self._table.verticalHeader().setVisible(False)
        hh = self._table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for c in (1, 2, 3):
            hh.setSectionResizeMode(c, QHeaderView.ResizeMode.Fixed)
        self._table.setColumnWidth(1, 70)
        self._table.setColumnWidth(2, 110)
        self._table.setColumnWidth(3, 120)
        self._body_layout.addWidget(self._table)

        btn_row = QHBoxLayout()
        add_btn = QPushButton("＋  Add Row"); add_btn.setStyleSheet(_BTN_STYLE)
        add_btn.clicked.connect(self.add_row)
        del_btn = QPushButton("－  Remove Selected"); del_btn.setStyleSheet(_DEL_BTN_STYLE)
        del_btn.clicked.connect(self._remove_row)
        btn_row.addWidget(add_btn); btn_row.addWidget(del_btn); btn_row.addStretch()
        self._body_layout.addLayout(btn_row)

        sub_row = QHBoxLayout()
        sub_row.addStretch()
        sub_row.addWidget(QLabel("Subtotal:"))
        self._subtotal_edit = _CalcEdit()
        self._subtotal_edit.setFixedWidth(130)
        sub_row.addWidget(self._subtotal_edit)
        self._body_layout.addLayout(sub_row)

        self.add_row()

    def add_row(self):
        r = self._table.rowCount()
        self._table.insertRow(r)
        self._table.setRowHeight(r, 30)

        desc = QLineEdit(); desc.setStyleSheet(_FIELD_STYLE)
        self._table.setCellWidget(r, 0, desc)

        qty = _NumEdit("0", decimals=0)
        qty.textChanged.connect(self._recalc)
        self._table.setCellWidget(r, 1, qty)

        cost_per = _NumEdit("0.00")
        cost_per.textChanged.connect(self._recalc)
        self._table.setCellWidget(r, 2, cost_per)

        total = _CalcEdit()
        self._table.setCellWidget(r, 3, total)

        self._rows.append((desc, qty, cost_per, total))
        self._update_height()
        self._recalc()

    def _remove_row(self):
        rows = sorted({i.row() for i in self._table.selectedIndexes()}, reverse=True)
        for r in rows:
            self._table.removeRow(r)
            if r < len(self._rows):
                self._rows.pop(r)
        self._update_height()
        self._recalc()

    def _update_height(self):
        n = max(1, self._table.rowCount())
        self._table.setFixedHeight(30 * n + 34)

    def _recalc(self):
        total = 0.0
        for _, qty, cost_per, cost_total in self._rows:
            v = qty.value() * cost_per.value()
            total += v
            cost_total.setText(_fmt_dollar(v))
        self._subtotal_edit.setText(_fmt_dollar(total))
        self.set_subtotal(_fmt_dollar(total))
        self._on_change()

    def subtotal(self) -> float:
        return sum(q.value() * c.value() for _, q, c, _ in self._rows)

    def get_data(self) -> dict:
        return {"rows": [(d.text(), q.text(), c.text()) for d, q, c, _ in self._rows]}

    def restore_data(self, d: dict):
        while self._table.rowCount():
            self._table.removeRow(0)
        self._rows.clear()
        for desc, qty, cost in d.get("rows", []):
            self.add_row()
            self._rows[-1][0].setText(desc)
            self._rows[-1][1].setText(qty)
            self._rows[-1][2].setText(cost)


# ══════════════════════════════════════════════════════════════════════════════
# Section 4 — Travel & Expenses
# ══════════════════════════════════════════════════════════════════════════════
class _TravelCard(_Card):
    _LABELS   = ["Airfare",   "Hotel (nights)", "Food (meals)", "Car (miles)"]
    _QTY_HINT = ["0",         "0",              "0",            "0"          ]
    _RATE_HINT= ["0.00",      "0.00",           "0.00",         "0.670"      ]
    _QTY_LBL  = ["Qty",       "Nights",         "Meals",        "Miles"      ]
    _RATE_LBL = ["Cost Each", "$/Night",        "$/Meal",       "$/Mile"     ]

    def __init__(self, on_change, parent=None):
        super().__init__("4.  Travel & Expenses", parent)
        self._on_change = on_change

        self._table = QTableWidget(len(self._LABELS), 4)
        self._table.setStyleSheet(_TABLE_STYLE)
        self._table.setHorizontalHeaderLabels(["Category", "Qty", "Rate", "Cost"])
        self._table.verticalHeader().setVisible(False)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        hh = self._table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for c in (1, 2, 3):
            hh.setSectionResizeMode(c, QHeaderView.ResizeMode.Fixed)
        self._table.setColumnWidth(1, 90)
        self._table.setColumnWidth(2, 110)
        self._table.setColumnWidth(3, 110)
        self._table.setFixedHeight(32 * len(self._LABELS) + 34)

        self._qty_edits  = []
        self._rate_edits = []
        self._cost_edits = []
        for r, lbl in enumerate(self._LABELS):
            cat = QTableWidgetItem(lbl)
            cat.setFlags(Qt.ItemFlag.ItemIsEnabled)
            cat.setForeground(QColor(_SUBTEXT))
            self._table.setItem(r, 0, cat)

            qty = _NumEdit(self._QTY_HINT[r])
            qty.setToolTip(self._QTY_LBL[r])
            qty.textChanged.connect(self._recalc)
            self._table.setCellWidget(r, 1, qty)
            self._qty_edits.append(qty)

            rate = _NumEdit(self._RATE_HINT[r])
            rate.setToolTip(self._RATE_LBL[r])
            rate.textChanged.connect(self._recalc)
            self._table.setCellWidget(r, 2, rate)
            self._rate_edits.append(rate)

            cost = _CalcEdit()
            self._table.setCellWidget(r, 3, cost)
            self._cost_edits.append(cost)

        self._body_layout.addWidget(self._table)

        sub_row = QHBoxLayout()
        sub_row.addStretch()
        sub_row.addWidget(QLabel("Subtotal:"))
        self._subtotal_edit = _CalcEdit()
        self._subtotal_edit.setFixedWidth(130)
        sub_row.addWidget(self._subtotal_edit)
        self._body_layout.addLayout(sub_row)

        self._recalc()

    def _recalc(self):
        total = 0.0
        for r in range(len(self._LABELS)):
            v = self._qty_edits[r].value() * self._rate_edits[r].value()
            total += v
            self._cost_edits[r].setText(_fmt_dollar(v))
        self._subtotal_edit.setText(_fmt_dollar(total))
        self.set_subtotal(_fmt_dollar(total))
        self._on_change()

    def subtotal(self) -> float:
        return sum(q.value() * r.value()
                   for q, r in zip(self._qty_edits, self._rate_edits))

    def get_data(self) -> dict:
        return {
            "qty":  [e.text() for e in self._qty_edits],
            "rate": [e.text() for e in self._rate_edits],
        }

    def restore_data(self, d: dict):
        for i, v in enumerate(d.get("qty", [])):
            if i < len(self._qty_edits): self._qty_edits[i].setText(v)
        for i, v in enumerate(d.get("rate", [])):
            if i < len(self._rate_edits): self._rate_edits[i].setText(v)


# ══════════════════════════════════════════════════════════════════════════════
# Section 5 — Summary
# ══════════════════════════════════════════════════════════════════════════════
class _SummaryCard(QFrame):
    """Always-visible summary / rollup panel at the bottom."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(
            f"QFrame {{ background:{_BG_CARD}; border:2px solid {_RED};"
            "  border-radius:8px; }"
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 16)
        layout.setSpacing(8)

        title = QLabel("Cost Summary")
        title.setStyleSheet(
            f"QLabel {{ font-size:14px; font-weight:800; color:{_RED};"
            "  border:none; padding:0; }"
        )
        layout.addWidget(title)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"border:none; border-top:1px solid {_BORDER};")
        layout.addWidget(sep)

        grid = QWidget()
        grid.setStyleSheet("QWidget { background:transparent; border:none; }")
        gl = QVBoxLayout(grid)
        gl.setContentsMargins(0, 0, 0, 0)
        gl.setSpacing(6)

        def _row(label: str, is_input=False, hint="0.00", bold=False, accent=False):
            row = QHBoxLayout()
            lbl = QLabel(label)
            lbl.setFixedWidth(240)
            col = _RED if accent else _SUBTEXT
            w = 700 if bold else 400
            lbl.setStyleSheet(
                f"QLabel {{ font-size:12px; font-weight:{w}; color:{col};"
                "  border:none; padding:0; }"
            )
            row.addWidget(lbl)
            if is_input:
                ed = _NumEdit(hint)
                ed.setFixedWidth(130)
                val = ed
            else:
                ed = _CalcEdit()
                ed.setFixedWidth(130)
                val = ed
            row.addWidget(val)
            row.addStretch()
            gl.addLayout(row)
            return val

        self._cost_before_risk = _row("Cost Before Risk:", bold=True)
        self._risk_pct         = _row("Risk / Insurance  (%):", is_input=True, hint="0.00")
        self._risk_cost        = _row("  Risk Cost:")
        self._estimated_cost   = _row("Estimated Cost:", bold=True)
        self._margin_pct       = _row("Margin  (%):", is_input=True, hint="32.5")
        self._resale           = _row("Recommended Resale:", bold=True, accent=True)
        self._profit           = _row("Profit:", bold=True, accent=True)

        layout.addWidget(grid)

    def update(self, consulting: float, third_party: float,
               materials: float, travel: float):
        cost_before = consulting + third_party + materials + travel
        risk_pct    = _num(self._risk_pct)
        risk_cost   = cost_before * risk_pct / 100.0
        estimated   = cost_before + risk_cost
        margin_pct  = _num(self._margin_pct)
        if 0 < margin_pct < 100:
            resale = estimated / (1.0 - margin_pct / 100.0)
        else:
            resale = estimated
        profit = resale - estimated

        self._cost_before_risk.setText(_fmt_dollar(cost_before))
        self._risk_cost.setText(_fmt_dollar(risk_cost))
        self._estimated_cost.setText(_fmt_dollar(estimated))
        self._resale.setText(_fmt_dollar(resale))
        self._profit.setText(_fmt_dollar(profit))

    def get_data(self) -> dict:
        return {
            "risk_pct":   self._risk_pct.text(),
            "margin_pct": self._margin_pct.text(),
        }

    def restore_data(self, d: dict):
        self._risk_pct.setText(d.get("risk_pct", ""))
        self._margin_pct.setText(d.get("margin_pct", ""))


# ══════════════════════════════════════════════════════════════════════════════
# Top-level Cost Estimator widget
# ══════════════════════════════════════════════════════════════════════════════
class CostEstimatorWidget(QWidget):
    """
    Drop-in widget for the Cost Estimator tab.
    Mirrors the Excel OC Lev 4-ASPCostReport structure.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Scroll area for the cards
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet(
            "QScrollArea { background:#f0f2f7; border:none; }"
            "QScrollBar:vertical { background:#f0f2f7; width:8px; margin:0; }"
            "QScrollBar::handle:vertical { background:#c8cedd; border-radius:4px; min-height:20px; }"
            "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height:0; }"
        )

        content = QWidget()
        content.setStyleSheet("QWidget { background:#f0f2f7; }")
        cl = QVBoxLayout(content)
        cl.setContentsMargins(20, 16, 20, 16)
        cl.setSpacing(12)

        # Project info strip
        info_card = _Card("Project Info", parent=content)
        info_row = QHBoxLayout()
        info_row.setSpacing(12)
        for lbl_txt, attr, hint in [
            ("Client:",   "_client",  "Client name…"),
            ("Project:",  "_proj",    "Project name…"),
            ("Date:",     "_date",    "MM/DD/YYYY"),
        ]:
            lbl = QLabel(lbl_txt)
            lbl.setStyleSheet("QLabel { border:none; }")
            info_row.addWidget(lbl)
            ed = QLineEdit()
            ed.setPlaceholderText(hint)
            ed.setStyleSheet(_FIELD_STYLE)
            setattr(self, attr, ed)
            info_row.addWidget(ed)
        info_card._body_layout.addLayout(info_row)
        cl.addWidget(info_card)

        # Cost sections
        self._consulting = _ConsultingCard(self._recalculate, parent=content)
        self._third_party = _ThirdPartyCard(self._recalculate, parent=content)
        self._materials   = _MaterialsCard(self._recalculate, parent=content)
        self._travel      = _TravelCard(self._recalculate, parent=content)
        cl.addWidget(self._consulting)
        cl.addWidget(self._third_party)
        cl.addWidget(self._materials)
        cl.addWidget(self._travel)

        # Summary (always visible, not collapsible)
        self._summary = _SummaryCard(parent=content)
        # Wire summary inputs to recalculate
        self._summary._risk_pct.textChanged.connect(self._recalculate)
        self._summary._margin_pct.textChanged.connect(self._recalculate)
        cl.addWidget(self._summary)
        cl.addStretch()

        scroll.setWidget(content)
        root.addWidget(scroll)

        # ── Bottom action bar ──────────────────────────────────────────────
        bar = QWidget()
        bar.setStyleSheet(
            "QWidget { background:#f9f0f2; border-top:1px solid #d6c0c5; }"
        )
        bar.setFixedHeight(52)
        bar_row = QHBoxLayout(bar)
        bar_row.setContentsMargins(16, 8, 16, 8)
        bar_row.setSpacing(10)

        import_btn = QPushButton("⬆  Import MCMXC…")
        import_btn.setStyleSheet(
            "QPushButton { background:#ffffff; color:#920d2e;"
            "  border:1px solid #d6c0c5; border-radius:5px;"
            "  padding:7px 18px; font-size:12px; font-weight:600; }"
            "QPushButton:hover { background:#fdf0f3; border-color:#920d2e; }"
        )
        import_btn.clicked.connect(self.import_cost_sheet)
        bar_row.addWidget(import_btn)
        bar_row.addStretch()

        gen_btn = QPushButton("⬇  Generate Cost Sheet")
        gen_btn.setStyleSheet(
            "QPushButton { background:#920d2e; color:#ffffff;"
            "  border:none; border-radius:5px;"
            "  padding:7px 22px; font-size:13px; font-weight:700; }"
            "QPushButton:hover { background:#7a0b27; }"
            "QPushButton:pressed { background:#600820; }"
        )
        gen_btn.clicked.connect(self.generate_cost_sheet)
        bar_row.addWidget(gen_btn)
        root.addWidget(bar)

        # Initial calculation
        self._recalculate()

    def _recalculate(self):
        if not hasattr(self, '_summary'):
            return
        self._summary.update(
            self._consulting.subtotal(),
            self._third_party.subtotal(),
            self._materials.subtotal(),
            self._travel.subtotal(),
        )

    # ── Save / Load ────────────────────────────────────────────────────────
    def get_data(self) -> dict:
        return {
            "client":      getattr(self, "_client", None) and self._client.text() or "",
            "project":     self._proj.text(),
            "date":        self._date.text(),
            "consulting":  self._consulting.get_data(),
            "third_party": self._third_party.get_data(),
            "materials":   self._materials.get_data(),
            "travel":      self._travel.get_data(),
            "summary":     self._summary.get_data(),
        }

    def restore_data(self, d: dict):
        self._client.setText(d.get("client", ""))
        self._proj.setText(d.get("project", ""))
        self._date.setText(d.get("date", ""))
        self._consulting.restore_data(d.get("consulting", {}))
        self._third_party.restore_data(d.get("third_party", {}))
        self._materials.restore_data(d.get("materials", {}))
        self._travel.restore_data(d.get("travel", {}))
        self._summary.restore_data(d.get("summary", {}))
        self._recalculate()

    # ── Generate / Import ──────────────────────────────────────────────────
    def generate_cost_sheet(self):
        """Save .mcmxc + Excel breakdown side-by-side."""
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        client  = self._client.text().strip().replace(" ", "_") or "CostSheet"
        import re as _re
        safe = _re.sub(r'[\\/:*?"<>|]', "-", client)
        default = os.path.join(desktop, f"{safe}.mcmxc")
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Cost Sheet", default,
            "MCMX Cost Sheet (*.mcmxc)")
        if not path:
            return
        # ── .mcmxc (JSON) ─────────────────────────────────────────────────
        data = self.get_data()
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump({"_type": "mcmxc", "_version": 1, **data}, f,
                          indent=2, ensure_ascii=False)
        except Exception as e:
            QMessageBox.critical(self, "Save Failed", str(e)); return
        # ── Excel ──────────────────────────────────────────────────────────
        xlsx_path = os.path.splitext(path)[0] + ".xlsx"
        try:
            _export_excel(data, xlsx_path)
        except Exception as e:
            QMessageBox.warning(self, "Excel Export Failed", str(e))
            return
        QMessageBox.information(
            self, "Cost Sheet Generated",
            f"✔  Files saved:\n{path}\n{xlsx_path}")

    def import_cost_sheet(self):
        """Load a previously saved .mcmxc file."""
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Cost Sheet", desktop,
            "MCMX Cost Sheet (*.mcmxc)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            QMessageBox.critical(self, "Import Failed", str(e)); return
        self.restore_data(data)
        QMessageBox.information(self, "Imported", "✔  Cost sheet restored.")


# ══════════════════════════════════════════════════════════════════════════════
# Excel export
# ══════════════════════════════════════════════════════════════════════════════
def _export_excel(data: dict, path: str):
    """
    Write a formatted Excel workbook from the cost estimator data dict.
    Requires openpyxl (already a transitive dependency via python-docx).
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "Cost Breakdown"

    # ── Colour palette ────────────────────────────────────────────────────
    RED     = "920D2E"
    RED_LT  = "F5D0DA"
    GRAY    = "F4F6FA"
    BLUE_BG = "E8F0FE"
    BLUE_FG = "1A3A6E"
    WHITE   = "FFFFFF"
    DARK    = "1A0509"
    BDR_C   = "D6C0C5"

    thin = Side(style="thin", color=BDR_C)
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def _font(bold=False, color=DARK, size=11, italic=False):
        return Font(name="Aptos", bold=bold, color=color, size=size, italic=italic)

    def _align(h="left", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    # ── Column widths ─────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16

    row = 1

    # ── Title bar ─────────────────────────────────────────────────────────
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row, 1, "MCMX Cost Estimator")
    c.font = _font(bold=True, color=WHITE, size=16)
    c.fill = _fill(RED)
    c.alignment = _align("center")
    ws.row_dimensions[row].height = 28
    row += 1

    # ── Project info ──────────────────────────────────────────────────────
    for lbl, key in [("Client", "client"), ("Project", "project"), ("Date", "date")]:
        ws.cell(row, 1, lbl).font = _font(bold=True)
        ws.merge_cells(f"B{row}:D{row}")
        ws.cell(row, 2, data.get(key, "")).font = _font()
        ws.row_dimensions[row].height = 18
        row += 1
    row += 1  # blank

    # ── Helpers ───────────────────────────────────────────────────────────
    def _section_hdr(title):
        nonlocal row
        ws.merge_cells(f"A{row}:D{row}")
        c = ws.cell(row, 1, title)
        c.font = _font(bold=True, color=WHITE, size=12)
        c.fill = _fill(RED)
        c.alignment = _align("left")
        ws.row_dimensions[row].height = 22
        row += 1

    def _col_hdrs(*labels):
        nonlocal row
        for ci, lbl in enumerate(labels, 1):
            c = ws.cell(row, ci, lbl)
            c.font = _font(bold=True)
            c.fill = _fill(GRAY)
            c.alignment = _align("center")
            c.border = bdr
        ws.row_dimensions[row].height = 18
        row += 1

    def _data_row(*vals, shade=False):
        nonlocal row
        bg = GRAY if shade else WHITE
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row, ci, val)
            c.font = _font()
            c.fill = _fill(bg)
            c.alignment = _align("right" if ci > 1 else "left")
            c.border = bdr
            if isinstance(val, (int, float)) and ci > 1:
                c.number_format = "#,##0"
        ws.row_dimensions[row].height = 17
        row += 1

    def _subtotal_row(val):
        nonlocal row
        ws.merge_cells(f"A{row}:C{row}")
        lc = ws.cell(row, 1, "Subtotal")
        lc.font = _font(bold=True)
        lc.fill = _fill(RED_LT)
        lc.alignment = _align("right")
        lc.border = bdr
        vc = ws.cell(row, 4, val)
        vc.font = _font(bold=True, color=BLUE_FG)
        vc.fill = _fill(BLUE_BG)
        vc.alignment = _align("right")
        vc.border = bdr
        vc.number_format = "#,##0"
        ws.row_dimensions[row].height = 17
        row += 2  # blank after section

    def _summary_row(label, value, accent=False, pct_str=None):
        nonlocal row
        ws.merge_cells(f"A{row}:C{row}")
        lc = ws.cell(row, 1, label)
        lc.font = _font(bold=accent, color=RED if accent else DARK)
        lc.alignment = _align("right")
        lc.border = bdr
        if pct_str is not None:
            vc = ws.cell(row, 4, pct_str)
            vc.number_format = "@"
        else:
            vc = ws.cell(row, 4, value)
            vc.number_format = "#,##0"
        vc.font = _font(bold=accent, color=RED if accent else BLUE_FG)
        vc.fill = _fill(RED_LT if accent else BLUE_BG)
        vc.alignment = _align("right")
        vc.border = bdr
        ws.row_dimensions[row].height = 18
        row += 1

    # ─────────────────────────────────────────────────────────────────────
    # 1. Consulting Effort
    # ─────────────────────────────────────────────────────────────────────
    cons = data.get("consulting", {})
    rate = _num_s(cons.get("rate", "0"))
    hours_list = cons.get("rows", [])
    cat_labels = ["On-Site Time (hrs)", "Travel Time (hrs)", "Other (hrs)"]

    _section_hdr("1.  Consulting Effort")
    _col_hdrs("Category", "Hours", "Rate ($/hr)", "Cost")
    cons_sub = 0.0
    for i, lbl in enumerate(cat_labels):
        hrs = _num_s(hours_list[i] if i < len(hours_list) else "0")
        cost = math.ceil(hrs * rate)
        cons_sub += cost
        _data_row(lbl, hrs, rate, cost, shade=(i % 2 == 1))
    _subtotal_row(math.ceil(cons_sub))

    # ─────────────────────────────────────────────────────────────────────
    # 2. Third Party
    # ─────────────────────────────────────────────────────────────────────
    tp = data.get("third_party", {})
    _section_hdr("2.  Third Party")
    _col_hdrs("Description", "Qty", "", "Cost Total")
    tp_total = 0.0
    for i, row_vals in enumerate(tp.get("rows", [])):
        desc, qty, cost = (row_vals + ["", "0", "0"])[:3]
        q = _num_s(qty); c = _num_s(cost)
        v = math.ceil(q * c)
        tp_total += v
        _data_row(desc, q, "", v, shade=(i % 2 == 1))
    _subtotal_row(math.ceil(tp_total))

    # ─────────────────────────────────────────────────────────────────────
    # 3. Hardware / Software / Materials
    # ─────────────────────────────────────────────────────────────────────
    mat = data.get("materials", {})
    _section_hdr("3.  Hardware / Software / Materials")
    _col_hdrs("Description", "Qty", "Cost Per", "Cost Total")
    mat_total = 0.0
    for i, row_vals in enumerate(mat.get("rows", [])):
        desc, qty, cost = (row_vals + ["", "0", "0"])[:3]
        q = _num_s(qty); c = _num_s(cost)
        v = math.ceil(q * c)
        mat_total += v
        _data_row(desc, q, c, v, shade=(i % 2 == 1))
    _subtotal_row(math.ceil(mat_total))

    # ─────────────────────────────────────────────────────────────────────
    # 4. Travel & Expenses
    # ─────────────────────────────────────────────────────────────────────
    trav = data.get("travel", {})
    qty_list  = trav.get("qty",  [])
    rate_list = trav.get("rate", [])
    t_labels  = ["Airfare", "Hotel (nights)", "Food (meals)", "Car (miles)"]
    _section_hdr("4.  Travel & Expenses")
    _col_hdrs("Category", "Qty", "Rate", "Cost")
    trav_total = 0.0
    for i, lbl in enumerate(t_labels):
        q = _num_s(qty_list[i]  if i < len(qty_list)  else "0")
        r = _num_s(rate_list[i] if i < len(rate_list) else "0")
        v = math.ceil(q * r)
        trav_total += v
        _data_row(lbl, q, r, v, shade=(i % 2 == 1))
    _subtotal_row(math.ceil(trav_total))

    # ─────────────────────────────────────────────────────────────────────
    # 5. Summary
    # ─────────────────────────────────────────────────────────────────────
    summ = data.get("summary", {})
    risk_pct   = _num_s(summ.get("risk_pct",   "0"))
    margin_pct = _num_s(summ.get("margin_pct", "0"))
    cost_before = math.ceil(cons_sub) + math.ceil(tp_total) + math.ceil(mat_total) + math.ceil(trav_total)
    risk_cost   = math.ceil(cost_before * risk_pct / 100)
    estimated   = cost_before + risk_cost
    resale      = math.ceil(estimated / (1 - margin_pct / 100)) if 0 < margin_pct < 100 else estimated
    profit      = resale - estimated

    _section_hdr("5.  Cost Summary")
    _summary_row("Cost Before Risk",                    cost_before)
    _summary_row(f"Risk / Insurance ({risk_pct:.1f}%)", risk_cost)
    _summary_row("Estimated Cost",                      estimated)
    _summary_row(f"Margin ({margin_pct:.1f}%)",         None, pct_str=f"{margin_pct:.1f}%")
    _summary_row("Recommended Resale",                  resale,  accent=True)
    _summary_row("Profit",                              profit,  accent=True)

    wb.save(path)


def _num_s(s) -> float:
    """Parse a numeric string from saved cost data."""
    try:
        return float(str(s).replace(",", "").replace("$", "").strip() or 0)
    except (ValueError, TypeError):
        return 0.0
