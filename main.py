import sys
import json
import os as _os
import os
from datetime import date


def _resource_path(relative):
    """Return absolute path to a bundled resource, works for PyInstaller and dev."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QMessageBox,
    QWidget, QVBoxLayout, QDoubleSpinBox, QLabel, QHBoxLayout, QPushButton,
    QTextEdit, QSizePolicy, QTabWidget,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QIcon
from PyQt6.uic import loadUi

from form_widgets import DynamicFormWidget, DragDropLabel
from doc_generator import generate_doc
from cost_estimator import CostEstimatorWidget


APP_STYLE = """
    /* ── Base ── */
    QMainWindow, QWidget#centralwidget {
        background: #ffffff;
    }
    QLabel { color: #1a0509; font-family: 'Aptos'; font-size: 12px; }

    /* ── Inputs ── */
    QLineEdit, QTextEdit {
        border: 1px solid #d6c0c5; border-radius: 5px;
        padding: 6px 10px; font-family: 'Aptos'; font-size: 12px;
        background: #ffffff; color: #1a0509;
    }
    QLineEdit:focus, QTextEdit:focus { border-color: #920d2e; }

    /* ── Add Section / Add BOM buttons ── */
    QPushButton#button_add_table, QPushButton#button_add_paragraph {
        background: #ffffff; color: #1a0509;
        border: 1px solid #d6c0c5; border-radius: 5px;
        padding: 7px 16px; font-size: 12px;
    }
    QPushButton#button_add_table:hover,
    QPushButton#button_add_paragraph:hover {
        background: #fdf0f3; border-color: #920d2e;
    }

    /* ── Generate button ── */
    QPushButton#button_generate_doc {
        background: #920d2e; color: #ffffff;
        border: none; border-radius: 5px;
        padding: 8px 22px; font-size: 13px; font-weight: 700;
    }
    QPushButton#button_generate_doc:hover  { background: #7a0b27; }
    QPushButton#button_generate_doc:pressed { background: #600820; }

    /* ── Picture buttons ── */
    QPushButton#button_upload_picture {
        background: #ffffff; color: #1a0509;
        border: 1px solid #d6c0c5; border-radius: 5px;
        padding: 6px 14px; font-size: 11px;
    }
    QPushButton#button_upload_picture:hover { background: #fdf0f3; border-color: #920d2e; }
    QPushButton#button_clear_picture {
        background: #ffffff; color: #920d2e;
        border: 1px solid #d6c0c5; border-radius: 5px;
        padding: 6px 14px; font-size: 11px;
    }
    QPushButton#button_clear_picture:hover { background: #fdf0f3; }

    /* ── Collapse button ── */
    QPushButton#button_collapse_top {
        background: #ffffff; border: 1px solid #d6c0c5;
        border-radius: 5px; padding: 6px 14px;
        font-size: 12px; font-weight: 600; color: #920d2e;
        text-align: left;
    }
    QPushButton#button_collapse_top:hover { background: #fdf0f3; }

    /* ── Project info panel ── */
    QWidget#widget_top_panel {
        background: #ffffff;
        border: 1px solid #d6c0c5;
        border-radius: 8px;
    }

    /* ── Scroll ── */
    QScrollArea { border: none; background: transparent; }
    QScrollArea > QWidget > QWidget { background: transparent; }
    QScrollBar:vertical {
        width: 8px; background: #faf6f7; border-radius: 4px;
    }
    QScrollBar::handle:vertical {
        background: #d6c0c5; border-radius: 4px; min-height: 30px;
    }
    QScrollBar::handle:vertical:hover { background: #920d2e; }
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }

    /* ── Spinboxes ── */
    QDoubleSpinBox, QSpinBox {
        border: 1px solid #d6c0c5; border-radius: 4px;
        padding: 4px 8px; font-size: 12px;
        background: #ffffff; color: #1a0509;
    }
    QDoubleSpinBox:focus, QSpinBox:focus { border-color: #920d2e; }

    /* ── Checkboxes ── */
    QCheckBox { color: #1a0509; font-family: 'Aptos'; font-size: 12px; }
"""



def _format_date_ordinal(d: date) -> str:
    day = d.day
    suffix = ("th" if 11 <= day <= 13 else
              {1:"st",2:"nd",3:"rd"}.get(day % 10, "th"))
    return d.strftime(f"%B {day}{suffix}, %Y")



_MSG_STYLE = """
    QMessageBox {
        background: #ffffff;
        color: #1a1a2e;
    }
    QMessageBox QLabel {
        color: #1a1a2e;
        background: transparent;
        font-size: 12px;
    }
    QMessageBox QPushButton {
        background: #f4f6fa;
        color: #1a1a2e;
        border: 1px solid #c8cedd;
        border-radius: 4px;
        padding: 6px 18px;
        min-width: 72px;
        font-size: 12px;
    }
    QMessageBox QPushButton:hover   { background: #e0e6f0; }
    QMessageBox QPushButton:pressed { background: #d0d8ef; }
"""

def _msg(parent, kind, title, text):
    """Styled QMessageBox — prevents dark theme bleed-through."""
    box = QMessageBox(parent)
    box.setWindowTitle(title)
    box.setText(text)
    box.setStyleSheet(_MSG_STYLE)
    if kind == "info":
        box.setIcon(QMessageBox.Icon.Information)
    elif kind == "warn":
        box.setIcon(QMessageBox.Icon.Warning)
    elif kind == "crit":
        box.setIcon(QMessageBox.Icon.Critical)
    box.exec()


class PlainPasteTextEdit(QTextEdit):
    """
    QTextEdit that always pastes as plain text and enforces
    Aptos 11pt black regardless of clipboard source formatting.
    """
    _FONT   = "Aptos"
    _SIZE   = 11.0
    _COLOR  = "#1a1a2e"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._apply_default_fmt()

    def _apply_default_fmt(self):
        from PyQt6.QtGui import QTextCharFormat, QFont, QColor, QTextCursor
        fmt = QTextCharFormat()
        fmt.setFontFamily(self._FONT)
        fmt.setFontPointSize(self._SIZE)
        fmt.setForeground(QColor(self._COLOR))
        fmt.setFontWeight(QFont.Weight.Normal)
        c = self.textCursor()
        c.select(QTextCursor.SelectionType.Document)
        c.setCharFormat(fmt)
        c.clearSelection()
        self.setTextCursor(c)
        self.setCurrentCharFormat(fmt)

    def insertFromMimeData(self, source):
        """Paste plain text only — strips all colour/font/size from clipboard."""
        from PyQt6.QtCore import QMimeData
        plain = QMimeData()
        plain.setText(source.text())
        super().insertFromMimeData(plain)
        self._apply_default_fmt()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        ui_path = _resource_path("mainwindow.ui")
        loadUi(ui_path, self)
        self.setStyleSheet(APP_STYLE)

        # ── App version label (bottom-right) ──────────────────────────
        _ver_lbl = QLabel("V1.1")
        _ver_lbl.setStyleSheet(
            "QLabel { color:#b0a0a4; font-size:10px; padding:0 6px 2px 0; }")
        self.statusBar().addPermanentWidget(_ver_lbl)
        self.statusBar().setStyleSheet(
            "QStatusBar { background:#f9f4f5; border-top:1px solid #e8dde0; }")

        # ── Menu bar ──────────────────────────────────────────────────
        from PyQt6.QtWidgets import QMenuBar
        from PyQt6.QtGui import QAction
        menubar = self.menuBar()
        menubar.setStyleSheet(
            "QMenuBar { background:#920d2e; color:#ffffff; font-size:12px;"
            "  font-weight:600; padding:2px 4px; }"
            "QMenuBar::item { background:transparent; padding:4px 12px; }"
            "QMenuBar::item:selected { background:#7a0b27; border-radius:3px; }"
            "QMenu { background:#ffffff; color:#1a0509; border:1px solid #d6c0c5;"
            "  font-size:12px; }"
            "QMenu::item { padding:6px 20px; }"
            "QMenu::item:selected { background:#f5d0da; color:#920d2e; }")
        file_menu = menubar.addMenu("File")
        act_import = QAction("Import MCMXQ…", self)
        act_import.setShortcut("Ctrl+O")
        act_import.triggered.connect(self.load_project)
        file_menu.addAction(act_import)
        act_import_c = QAction("Import MCMXC…", self)
        act_import_c.triggered.connect(
            lambda: self._cost_estimator.import_cost_sheet())
        file_menu.addAction(act_import_c)
        file_menu.addSeparator()
        act_save = QAction("Save MCMXQ…", self)
        act_save.setShortcut("Ctrl+S")
        act_save.triggered.connect(self.save_project)
        file_menu.addAction(act_save)
        file_menu.addSeparator()
        act_quit = QAction("Quit", self)
        act_quit.setShortcut("Ctrl+Q")
        act_quit.triggered.connect(self.close)
        file_menu.addAction(act_quit)

        # Set margins and spacing in code (contentsMargins not supported by uic)
        self.centralwidget.layout().setContentsMargins(20, 16, 20, 16)
        self.centralwidget.layout().setSpacing(10)
        if self.widget_top_panel.layout():
            self.widget_top_panel.layout().setContentsMargins(16, 14, 16, 14)

        # ── Header bar with logo ───────────────────────────────────────
        import os as _os2
        _logo_path = _resource_path('logo.png')
        _header_bar = QWidget()
        _header_bar.setStyleSheet('background:transparent;')
        _hbl = QHBoxLayout(_header_bar)
        _hbl.setContentsMargins(0, 0, 0, 0); _hbl.setSpacing(12)
        if _os2.path.exists(_logo_path):
            from PyQt6.QtGui import QPixmap
            _logo_lbl = QLabel()
            _logo_lbl.setPixmap(
                QPixmap(_logo_path).scaled(
                    48, 48,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation))
            _hbl.addWidget(_logo_lbl)
        _title = self.label_app_title
        _title.setStyleSheet(
            'font-size:20px; font-weight:800; color:#920d2e; padding-bottom:2px;')
        _title.setParent(None)
        _hbl.addWidget(_title)
        _hbl.addStretch()
        self.centralwidget.layout().insertWidget(0, _header_bar)
        self.scrollAreaWidgetContents.layout().setContentsMargins(12, 12, 12, 12)
        self.scrollAreaWidgetContents.layout().setSpacing(12)

        # Replace plain textEdit_contact with PlainPasteTextEdit
        # Uses replaceWidget — same pattern as the picture label replacement
        old_contact = self.textEdit_contact
        self.textEdit_contact = PlainPasteTextEdit()
        self.textEdit_contact.setObjectName('textEdit_contact')
        self.textEdit_contact.setMinimumSize(old_contact.minimumSize())
        self.textEdit_contact.setMaximumSize(old_contact.maximumSize())
        self.textEdit_contact.setPlaceholderText(old_contact.placeholderText())
        self.formLayout_top.replaceWidget(old_contact, self.textEdit_contact)
        old_contact.hide()
        old_contact.deleteLater()
        # Restore tab order: location → contact → picture area
        QWidget.setTabOrder(self.lineEdit_location, self.textEdit_contact)

        # ── Global margin spinner (injected into top panel) ────────────────
        # Global margin — +/- button control matching BOM section style
        self._global_margin_value = 20.0

        # Hidden QDoubleSpinBox kept as compatibility shim for signal wiring
        self._global_margin_spin = QDoubleSpinBox()
        self._global_margin_spin.setRange(0, 99.9)
        self._global_margin_spin.setDecimals(1)
        self._global_margin_spin.setValue(self._global_margin_value)
        self._global_margin_spin.setVisible(False)
        self._global_margin_spin.valueChanged.connect(self._on_global_margin_changed)

        margin_row = QWidget()
        margin_row.setFixedHeight(44)
        margin_row.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        margin_layout = QHBoxLayout(margin_row)
        margin_layout.setContentsMargins(0, 6, 0, 6)
        margin_layout.setSpacing(0)

        margin_lbl = QLabel("Global Margin:")
        margin_lbl.setStyleSheet("font-size:12px; font-weight:700; color:#920d2e;")
        margin_layout.addWidget(margin_lbl)
        margin_layout.addSpacing(10)

        _btn_style = (
            "QPushButton { background:#f4f6fa; color:#3a3a5c;"
            "  border:1px solid #c8cedd; border-radius:4px;"
            "  font-size:16px; font-weight:700; }"
            "QPushButton:hover { background:#e0e4ef; }"
            "QPushButton:pressed { background:#d0d8ef; }"
        )

        self._global_margin_minus = QPushButton("−")
        self._global_margin_minus.setFixedSize(30, 30)
        self._global_margin_minus.setStyleSheet(_btn_style)
        self._global_margin_minus.setToolTip("Decrease global margin by 1%")
        self._global_margin_minus.clicked.connect(self._decrement_global_margin)
        margin_layout.addWidget(self._global_margin_minus)
        margin_layout.addSpacing(8)

        from PyQt6.QtWidgets import QLineEdit as _QLE
        self._global_margin_lbl = _QLE(f"{self._global_margin_value:.1f}")
        self._global_margin_lbl.setFixedWidth(74)
        self._global_margin_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._global_margin_lbl.setStyleSheet(
            "QLineEdit { background:#ffffff; border:1px solid #6a8fd8;"
            "  border-radius:4px; color:#920d2e; font-size:12px;"
            "  font-weight:600; padding:4px 6px; }")
        self._global_margin_lbl.setToolTip(
            "Global margin applied to all BOM tables.\n"
            "Formula: Price = Cost / (1 - Margin%/100)")
        self._global_margin_lbl.editingFinished.connect(self._on_global_margin_typed)
        margin_layout.addWidget(self._global_margin_lbl)
        margin_layout.addSpacing(8)

        self._global_margin_plus = QPushButton("+")
        self._global_margin_plus.setFixedSize(30, 30)
        self._global_margin_plus.setStyleSheet(_btn_style)
        self._global_margin_plus.setToolTip("Increase global margin by 1%")
        self._global_margin_plus.clicked.connect(self._increment_global_margin)
        margin_layout.addWidget(self._global_margin_plus)

        # Margin row lives outside widget_top_panel so it stays visible
        # when Project Info is collapsed
        self._margin_row_widget = margin_row
        # Will be inserted into centralwidget layout after UI loads — see below

        # Insert margin row between collapse button (idx 1) and top panel (idx 2)
        # Insert AFTER widget_top_panel (index 3) so it sits below Project Info
        self.centralwidget.layout().insertWidget(3, self._margin_row_widget)

        # ── Version row ────────────────────────────────────────────────────
        # Uses the same +/− QPushButton style as the Global Margin widget so
        # the arrows are always clearly visible.

        _VER_BTN = (
            "QPushButton { background:#f4f6fa; color:#3a3a5c; border:1px solid #c8cedd;"
            "  border-radius:4px; font-size:16px; font-weight:700; }"
            "QPushButton:hover { background:#e0e4ef; }"
            "QPushButton:pressed { background:#d0d8ef; }")
        _VER_VAL = (
            "QLabel { background:#ffffff; border:1px solid #d6c0c5;"
            "  border-radius:4px; color:#1a0509; font-size:12px;"
            "  font-weight:700; padding:4px 8px; }")

        class _VerSpin:
            """Lightweight +/− spinner with explicit button widgets."""
            def __init__(self_, lo, hi, initial, on_change):
                self_._val = initial
                self_._lo  = lo
                self_._hi  = hi
                self_._cb  = on_change
                self_.widget = QWidget()
                _row = QHBoxLayout(self_.widget)
                _row.setContentsMargins(0, 0, 0, 0)
                _row.setSpacing(3)
                self_._minus = QPushButton("−")
                self_._minus.setFixedSize(28, 28)
                self_._minus.setStyleSheet(_VER_BTN)
                self_._minus.clicked.connect(self_._dec)
                _row.addWidget(self_._minus)
                self_._lbl = QLabel(str(initial))
                self_._lbl.setFixedWidth(36)
                self_._lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
                self_._lbl.setStyleSheet(_VER_VAL)
                _row.addWidget(self_._lbl)
                self_._plus = QPushButton("+")
                self_._plus.setFixedSize(28, 28)
                self_._plus.setStyleSheet(_VER_BTN)
                self_._plus.clicked.connect(self_._inc)
                _row.addWidget(self_._plus)
            def value(self_):       return self_._val
            def setValue(self_, v):
                self_._val = max(self_._lo, min(self_._hi, int(v)))
                self_._lbl.setText(str(self_._val))
            def _inc(self_):        self_.setValue(self_._val + 1); self_._cb()
            def _dec(self_):        self_.setValue(self_._val - 1); self_._cb()

        ver_row = QWidget()
        ver_row.setFixedHeight(44)
        ver_row.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        ver_layout = QHBoxLayout(ver_row)
        ver_layout.setContentsMargins(0, 6, 0, 6)
        ver_layout.setSpacing(0)

        ver_lbl = QLabel("Version:")
        ver_lbl.setStyleSheet("font-size:12px; font-weight:700; color:#920d2e;")
        ver_layout.addWidget(ver_lbl)
        ver_layout.addSpacing(10)

        self._version_major_spin = _VerSpin(1, 99, 1, self._update_version_label)
        ver_layout.addWidget(self._version_major_spin.widget)

        _dot_lbl = QLabel(".")
        _dot_lbl.setStyleSheet("font-size:16px; font-weight:700; color:#1a0509;")
        _dot_lbl.setFixedWidth(10)
        _dot_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ver_layout.addWidget(_dot_lbl)

        self._version_minor_spin = _VerSpin(0, 99, 0, self._update_version_label)
        ver_layout.addWidget(self._version_minor_spin.widget)

        ver_layout.addSpacing(12)
        self._version_badge = QLabel("V1.0")
        self._version_badge.setStyleSheet(
            "QLabel { background:#9E1B32; color:#ffffff; font-size:11px;"
            "  font-weight:700; border-radius:4px; padding:3px 10px; }")
        ver_layout.addWidget(self._version_badge)
        self.centralwidget.layout().insertWidget(4, ver_row)

        # ── Change History panel ────────────────────────────────────────────
        from PyQt6.QtWidgets import QListWidget as _QLW

        self._history_collapsed = True
        _hist_outer = QWidget()
        _hist_vbox = QVBoxLayout(_hist_outer)
        _hist_vbox.setContentsMargins(0, 0, 0, 0)
        _hist_vbox.setSpacing(4)

        self._btn_collapse_hist = QPushButton("▶  Change History")
        self._btn_collapse_hist.setObjectName("button_collapse_hist")
        self._btn_collapse_hist.setStyleSheet(
            "QPushButton { background:#ffffff; border:1px solid #d6c0c5;"
            "  border-radius:5px; padding:6px 14px; font-size:12px;"
            "  font-weight:600; color:#920d2e; text-align:left; }"
            "QPushButton:hover { background:#fdf0f3; }")
        self._btn_collapse_hist.clicked.connect(self._toggle_history_panel)
        _hist_vbox.addWidget(self._btn_collapse_hist)

        self._history_panel = QWidget()
        self._history_panel.setVisible(False)
        _hp_vbox = QVBoxLayout(self._history_panel)
        _hp_vbox.setContentsMargins(0, 2, 0, 2)
        _hp_vbox.setSpacing(4)

        self._version_history = _QLW()
        self._version_history.setMaximumHeight(130)
        self._version_history.setStyleSheet(
            "QListWidget { background:#ffffff; border:1px solid #d6c0c5;"
            "  border-radius:4px; font-size:11px; color:#1a0509; }"
            "QListWidget::item { padding:4px 8px; }"
            "QListWidget::item:selected { background:#f5d0da; color:#920d2e; }")
        _hp_vbox.addWidget(self._version_history)

        _hist_btn_row = QWidget()
        _hbr_layout = QHBoxLayout(_hist_btn_row)
        _hbr_layout.setContentsMargins(0, 0, 0, 0)
        _hbr_layout.setSpacing(8)
        _hist_add_btn = QPushButton("+ Add Entry")
        _hist_add_btn.setStyleSheet(
            "QPushButton { background:#f4f6fa; color:#3a3a5c;"
            "  border:1px solid #dde1e7; border-radius:4px;"
            "  padding:4px 12px; font-size:11px; }"
            "QPushButton:hover { background:#e8ecf5; border-color:#b0b8d0; }")
        _hist_add_btn.clicked.connect(self._add_version_entry)
        _hbr_layout.addWidget(_hist_add_btn)
        _hist_del_btn = QPushButton("Remove")
        _hist_del_btn.setStyleSheet(
            "QPushButton { background:transparent; color:#e05252;"
            "  border:1px solid #e05252; border-radius:4px;"
            "  padding:4px 12px; font-size:11px; }"
            "QPushButton:hover { background:#fdf2f2; }")
        _hist_del_btn.clicked.connect(self._remove_version_entry)
        _hbr_layout.addWidget(_hist_del_btn)
        _hbr_layout.addStretch()
        _hp_vbox.addWidget(_hist_btn_row)

        _hist_vbox.addWidget(self._history_panel)
        self.centralwidget.layout().insertWidget(5, _hist_outer)

        # ── Replace placeholder with DragDropLabel in a grouped frame ────
        self._picture_label = DragDropLabel(parent=self.widget_top_panel)
        _pic_parent_layout = self.label_picture_preview.parent().layout()
        if _pic_parent_layout is not None:
            _pic_idx = _pic_parent_layout.indexOf(self.label_picture_preview)
            _pic_parent_layout.insertWidget(_pic_idx, self._picture_label)
        self.label_picture_preview.hide()
        self.label_picture_preview.deleteLater()
        # Lock the DragDropLabel to its fixed 150×150 size
        self._picture_label.setFixedSize(150, 150)

        # Style the picture VBoxLayout container as a clearly grouped panel
        _pic_container = self._picture_label.parent()
        if _pic_container and _pic_container is not self.widget_top_panel:
            _pic_container.setStyleSheet(
                "QWidget { background:#f9f0f2; border:1px solid #d6c0c5;"
                "  border-radius:6px; padding:4px; }")
            # Override buttons inside so they don't inherit the container border
            self.button_upload_picture.setStyleSheet(
                self.button_upload_picture.styleSheet() +
                "QPushButton { border-radius:4px; }")
            self.button_clear_picture.setStyleSheet(
                self.button_clear_picture.styleSheet() +
                "QPushButton { border-radius:4px; }")
        # Also label the group clearly
        self.label_picture.setText("Customer Logo")
        self.label_picture.setStyleSheet(
            "QLabel { font-size:11px; font-weight:700; color:#920d2e;"
            "  background:transparent; border:none; padding:0; }")

        # ── Scroll area ────────────────────────────────────────────────────
        scroll_widget = self.scrollArea_sections.widget()
        if scroll_widget.layout() is None:
            scroll_widget.setLayout(QVBoxLayout())
        scroll_widget.layout().setAlignment(Qt.AlignmentFlag.AlignTop)

        self.form_widget = DynamicFormWidget(
            scroll_widget.layout(),
            global_margin_ref=self._get_global_margin
        )
        self.form_widget._global_spin = self._global_margin_spin

        # ── Live TOC tree — placed inside widget_top_panel to the right ──
        from PyQt6.QtWidgets import QTreeWidget, QTreeWidgetItem
        self._toc_tree = QTreeWidget()
        self._toc_tree.setHeaderLabel("Table of Contents")
        self._toc_tree.setMinimumWidth(200)
        self._toc_tree.setStyleSheet(
            "QTreeWidget { background:#ffffff; border:1px solid #d6c0c5;"
            "  border-radius:6px; font-size:11px; color:#1a0509; }"
            "QTreeWidget::item { padding:2px 4px; }"
            "QTreeWidget::item:selected { background:#f5d0da; color:#920d2e; }"
            "QHeaderView::section { background:#f9f0f2; color:#920d2e;"
            "  font-weight:700; font-size:11px; border:none;"
            "  border-bottom:1px solid #d6c0c5; padding:4px 8px; }")
        self._toc_tree.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        # Inject TOC tree into the top panel's horizontal layout
        # (right of the picture group)
        _top_hbox = self.widget_top_panel.layout()
        if _top_hbox is not None:
            # Remove the trailing horizontal spacer if present
            for _i in range(_top_hbox.count() - 1, -1, -1):
                _item = _top_hbox.itemAt(_i)
                if _item and _item.spacerItem():
                    _top_hbox.removeItem(_item)
                    break
            _top_hbox.addWidget(self._toc_tree)
        # Hook DynamicFormWidget to refresh TOC when sections change
        self.form_widget._toc_refresh_cb = self._refresh_toc
        # Also poll every 2s as belt-and-suspenders for sub-header updates
        from PyQt6.QtCore import QTimer
        self._toc_timer = QTimer(self)
        self._toc_timer.setInterval(1500)
        self._toc_timer.timeout.connect(self._refresh_toc)
        self._toc_timer.start()

        # ── Buttons ────────────────────────────────────────────────────────
        self.button_add_table.clicked.connect(self._add_table)
        self.button_add_paragraph.clicked.connect(self._add_paragraph)
        self.button_generate_doc.clicked.connect(self.generate_document)

        self.button_upload_picture.clicked.connect(self.upload_picture)
        self.button_clear_picture.clicked.connect(self._picture_label.clear_image)
        self.button_collapse_top.clicked.connect(self._toggle_top_panel)

        self._top_collapsed = False

        # ── Tab wrapper ─────────────────────────────────────────────────────
        # Move everything below the header bar into a "Quote Generator" tab
        # and add a "Cost Estimator" tab alongside it.
        _main_layout = self.centralwidget.layout()

        # Collect all items from index 1 onwards (everything after _header_bar)
        _tab_items = []
        while _main_layout.count() > 1:
            _item = _main_layout.takeAt(1)
            if _item.widget():
                _tab_items.append(_item.widget())
            elif _item.layout():
                _tab_items.append(_item.layout())

        # Quote tab container
        _quote_tab = QWidget()
        _quote_tab.setStyleSheet("QWidget { background:transparent; }")
        _qt_layout = QVBoxLayout(_quote_tab)
        _qt_layout.setContentsMargins(0, 8, 0, 0)
        _qt_layout.setSpacing(10)
        for _w in _tab_items:
            if isinstance(_w, QWidget):
                _qt_layout.addWidget(_w)
            else:
                _qt_layout.addLayout(_w)

        # Cost Estimator tab
        self._cost_estimator = CostEstimatorWidget()

        # Tab widget
        self._tabs = QTabWidget()
        self._tabs.setStyleSheet(
            "QTabWidget::pane { border:none; background:transparent; }"
            "QTabBar::tab {"
            "  background:#f9f0f2; color:#920d2e;"
            "  border:1px solid #d6c0c5; border-bottom:none;"
            "  border-radius:5px 5px 0 0;"
            "  padding:6px 20px; font-size:12px; font-weight:600; }"
            "QTabBar::tab:selected {"
            "  background:#920d2e; color:#ffffff; }"
            "QTabBar::tab:hover:!selected { background:#fdf0f3; }"
        )
        self._tabs.addTab(_quote_tab,              "Quote Generator")
        self._tabs.addTab(self._cost_estimator,    "Cost Estimator")
        _main_layout.addWidget(self._tabs)

    def _get_global_margin(self) -> float:
        return self._global_margin_value

    def _increment_global_margin(self):
        self._set_global_margin(min(99.0, self._global_margin_value + 1.0))

    def _decrement_global_margin(self):
        self._set_global_margin(max(0.0, self._global_margin_value - 1.0))

    def _on_global_margin_typed(self):
        try:
            v = float(self._global_margin_lbl.text().replace('%', '').strip())
        except ValueError:
            self._global_margin_lbl.setText(f"{self._global_margin_value:.1f}")
            return
        self._set_global_margin(max(0.0, min(99.9, v)))

    def _set_global_margin(self, v: float):
        self._global_margin_value = round(v, 1)
        self._global_margin_lbl.setText(f"{self._global_margin_value:.1f}")
        # Sync hidden spin so existing signal chain fires
        self._global_margin_spin.blockSignals(True)
        self._global_margin_spin.setValue(self._global_margin_value)
        self._global_margin_spin.blockSignals(False)
        self._on_global_margin_changed()


    def _on_global_margin_changed(self):
        self.form_widget.set_global_margin_ref(
            self._get_global_margin, spin=self._global_margin_spin)
        # Directly notify each TableSection so Margin % cells update now
        from form_widgets import TableSection
        for section in self.form_widget.sections:
            if isinstance(section, TableSection):
                section._on_global_spin_changed()

    def _toggle_top_panel(self):
        self._top_collapsed = not self._top_collapsed
        self.widget_top_panel.setVisible(not self._top_collapsed)
        self.button_collapse_top.setText(
            "▶  Project Info" if self._top_collapsed else "▼  Project Info")

    # ── Version helpers ────────────────────────────────────────────────────
    def _version_str(self) -> str:
        return f"V{self._version_major_spin.value()}.{self._version_minor_spin.value()}"

    def _update_version_label(self):
        self._version_badge.setText(self._version_str())

    def _toggle_history_panel(self):
        self._history_collapsed = not self._history_collapsed
        self._history_panel.setVisible(not self._history_collapsed)
        self._btn_collapse_hist.setText(
            "▶  Change History" if self._history_collapsed
            else "▼  Change History")

    def _add_version_entry(self):
        from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout,
                                     QDialogButtonBox, QLabel, QLineEdit,
                                     QTextEdit as _QTE)
        v = self._version_str()
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Add Change Entry — {v}")
        dlg.setFixedWidth(420)
        dlg.setStyleSheet(
            "QDialog { background:#ffffff; }"
            "QLabel  { color:#1a0509; font-size:12px; }"
            "QLineEdit { border:1px solid #d6c0c5; border-radius:4px;"
            "  padding:6px 10px; font-size:12px; color:#1a0509; background:#fafbfd; }"
            "QLineEdit:focus { border-color:#920d2e; }"
            "QPushButton { background:#f4f6fa; color:#1a0509;"
            "  border:1px solid #d6c0c5; border-radius:4px;"
            "  padding:6px 18px; min-width:72px; font-size:12px; }"
            "QPushButton:hover { background:#e0e6f0; }"
        )
        vbox = QVBoxLayout(dlg)
        vbox.setSpacing(10)
        vbox.setContentsMargins(16, 16, 16, 16)
        ver_row_dlg = QLabel(f"<b>Version:</b>  {v}")
        vbox.addWidget(ver_row_dlg)
        desc = QLineEdit()
        desc.setPlaceholderText("Describe this change…")
        vbox.addWidget(desc)
        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        vbox.addWidget(btns)
        if dlg.exec() == QDialog.DialogCode.Accepted and desc.text().strip():
            today = date.today().strftime("%m/%d/%Y")
            self._version_history.addItem(
                f"{v}  —  {desc.text().strip()}  ({today})")

    def _remove_version_entry(self):
        row = self._version_history.currentRow()
        if row >= 0:
            self._version_history.takeItem(row)

    def _add_table(self):
        if not self._top_collapsed:
            self._toggle_top_panel()
        self.form_widget.add_section("table")

    def _add_paragraph(self):
        if not self._top_collapsed:
            self._toggle_top_panel()
        self.form_widget.add_section("paragraph")
        # Wire the newest section's editor directly to refresh TOC
        if self.form_widget.sections:
            s = self.form_widget.sections[-1]
            if hasattr(s, 'rich_editor'):
                s.rich_editor.editor.document().contentsChanged.connect(
                    self._refresh_toc)
                s.header.textChanged.connect(self._refresh_toc)

    def upload_picture(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "",
            "Images (*.png *.jpg *.jpeg *.bmp *.gif *.webp)")
        if path:
            self._picture_label.set_image(path)

    def _collect_fields(self):
        return {
            "project_name":      self.lineEdit_project.text(),
            "customer_name":     self.lineEdit_customer.text(),
            "customer_location": self.lineEdit_location.text(),
            "contact_info":      self.textEdit_contact.toPlainText(),
            "proposal_number":   self.lineEdit_proposal.text(),
            "customer_picture":  self._picture_label.image_path or "",
            "version_major":     self._version_major_spin.value(),
            "version_minor":     self._version_minor_spin.value(),
            "version_history":   [self._version_history.item(i).text()
                                  for i in range(self._version_history.count())],
        }

    def save_project(self):
        import re as _re
        desktop = _os.path.join(_os.path.expanduser('~'), 'Desktop')
        proposal = self.lineEdit_proposal.text().strip()
        safe = _re.sub(r'[\\/:*?"<>|]', '-', proposal) if proposal else 'project'
        _date_stamp = date.today().strftime('%Y%m%d')
        _date_part = '' if _re.search(r'\d{8}', safe) else f'-{_date_stamp}'
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Project",
            _os.path.join(desktop, f'{safe}{_date_part} {self._version_str()}.mcmxq'),
            "MCMXQ Project (*.mcmxq)")
        if not path:
            return
        state = self.form_widget.collect_project_state(self._collect_fields())
        state["cost_estimator"] = self._cost_estimator.get_data()
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(state, f, indent=2, ensure_ascii=False)
            _msg(self, "info", "Project Saved", f"✔  MCMXQ saved:\n{path}")
        except Exception as e:
            _msg(self, "crit", "Save Failed", str(e))

    def load_project(self):
        desktop = _os.path.join(_os.path.expanduser('~'), 'Desktop')
        path, _ = QFileDialog.getOpenFileName(
            self, "Import MCMXQ", desktop,
            "MCMXQ Project (*.mcmxq)")
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                state = json.load(f)
        except Exception as e:
            _msg(self, "crit", "Load Failed", str(e))
            return
        fields = state.get("fields", {})
        # Older MCMXQ files saved before version support was added may not have
        # version keys inside "fields". Check the root of the state dict as a
        # fallback so nothing is silently lost.
        def _fget(key, default):
            v = fields.get(key)
            if v is None:
                v = state.get(key)
            return v if v is not None else default

        self.lineEdit_project.setText(_fget("project_name", ""))
        self.lineEdit_customer.setText(_fget("customer_name", ""))
        self.lineEdit_location.setText(_fget("customer_location", ""))
        self.textEdit_contact.setPlainText(_fget("contact_info", ""))
        self.lineEdit_proposal.setText(_fget("proposal_number", ""))
        pic = _fget("customer_picture", "")
        if pic and _os.path.exists(pic):
            self._picture_label.set_image(pic)
        # Restore version
        self._version_major_spin.setValue(int(_fget("version_major", 1)))
        self._version_minor_spin.setValue(int(_fget("version_minor", 0)))
        self._version_history.clear()
        for entry in _fget("version_history", []):
            self._version_history.addItem(entry)
        self._update_version_label()
        self.form_widget.restore_project_state(state)
        if "cost_estimator" in state:
            self._cost_estimator.restore_data(state["cost_estimator"])
        _msg(self, "info", "Project Loaded", "✔  MCMXQ project restored.")

    def _refresh_toc(self):
        """Rebuild the live TOC tree from current sections."""
        from PyQt6.QtWidgets import QTreeWidgetItem
        from PyQt6.QtGui import QFont as _QF, QColor as _QC
        import re as _re
        self._toc_tree.clear()
        _bold_aptos = _QF('Aptos', 11); _bold_aptos.setBold(True)
        _norm_aptos = _QF('Aptos', 10)
        for i, section in enumerate(self.form_widget.sections):
            num = i + 1
            if hasattr(section, 'header'):
                hdr = section.header.text().strip() or f"Section {num}"
                top = QTreeWidgetItem([f"{num}.  {hdr}"])
                top.setForeground(0, _QC('#920d2e'))
                top.setFont(0, _bold_aptos)
                # Sub-headers: scan QTextDocument blocks directly for
                # 16pt bold dark-red text (more reliable than HTML regex)
                doc = section.rich_editor.editor.document()
                subs = []
                block = doc.begin()
                while block.isValid():
                    it = block.begin()
                    while not it.atEnd():
                        frag = it.fragment()
                        if frag.isValid():
                            cf = frag.charFormat()
                            col = cf.foreground().color()
                            # Detect header: dark-red foreground color
                            # #9E1B32 → r=139, g=0, b=0
                            # #9E1B32 = r=158, g=27, b=50
                            is_hdr = (
                                col.red() >= 140 and
                                col.green() >= 20 and col.green() <= 40 and
                                col.blue() >= 40 and col.blue() <= 65 and
                                col.red() > col.blue() + 80)
                            if is_hdr:
                                txt = frag.text().strip()
                                if txt and txt not in subs:
                                    subs.append(txt)
                        it += 1
                    block = block.next()
                for sub_i, sub in enumerate(subs, 1):
                    child = QTreeWidgetItem([f"{num}.{sub_i}  {sub}"])
                    child.setFont(0, _norm_aptos)
                    child.setForeground(0, _QC('#3a3a5c'))
                    top.addChild(child)
                self._toc_tree.addTopLevelItem(top)
                top.setExpanded(True)
            else:
                top = QTreeWidgetItem([f"{num}.  Bill of Material"])
                top.setFont(0, _bold_aptos)
                top.setForeground(0, _QC('#920d2e'))
                self._toc_tree.addTopLevelItem(top)

    def generate_document(self):
        if not self.lineEdit_project.text().strip():
            _msg(self, "warn", "Missing Field", "Please enter a Project Name.")
            return

        import os as _os, re as _re
        desktop = _os.path.join(_os.path.expanduser('~'), 'Desktop')
        proposal = self.lineEdit_proposal.text().strip()
        project  = self.lineEdit_project.text().strip()
        safe = _re.sub(r'[\\/:*?"<>|]', '-', proposal) if proposal else \
               _re.sub(r'[\\/:*?"<>|]', '-', project).replace(' ', '_') or 'output'
        _date_stamp = date.today().strftime('%Y%m%d')
        _date_part = '' if _re.search(r'\d{8}', safe) else f'-{_date_stamp}'
        default_path = _os.path.join(desktop,
                                     f'{safe}{_date_part} {self._version_str()}.docx')
        output_path, _ = QFileDialog.getSaveFileName(
            self, "Save Document", default_path, "Word Documents (*.docx)")
        if not output_path:
            return

        data = self.form_widget.get_form_data()
        data["project_name"]      = self.lineEdit_project.text().strip()
        data["customer_name"]     = self.lineEdit_customer.text().strip()
        data["customer_location"] = self.lineEdit_location.text().strip()
        data["contact_info"]      = self.textEdit_contact.toPlainText().strip()
        data["proposal_number"]   = self.lineEdit_proposal.text().strip()
        data["customer_picture"]  = self._picture_label.image_path
        data["today_date"]        = _format_date_ordinal(date.today())
        data["version"]           = self._version_str()

        try:
            result = generate_doc(data, output_path=output_path)
            # Auto-save MCMXQ alongside the docx
            _mcmxq_path = _os.path.splitext(output_path)[0] + '.mcmxq'
            try:
                _state = self.form_widget.collect_project_state(self._collect_fields())
                _state["cost_estimator"] = self._cost_estimator.get_data()
                with open(_mcmxq_path, 'w', encoding='utf-8') as _mf:
                    json.dump(_state, _mf, indent=2, ensure_ascii=False)
            except Exception as _me:
                print(f'MCMXQ auto-save failed: {_me}')
            _msg(self, "info", "Document Generated",
                 f"✔  Saved successfully:\n{result}\n\nProject file:\n{_mcmxq_path}")
        except Exception as e:
            _msg(self, "crit", "Generation Failed", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Set Aptos 11pt as the app-wide default font
    app.setFont(QFont("Aptos", 11))
    # Force light palette so dialogs and message boxes are always white
    from PyQt6.QtGui import QPalette, QColor as _QC
    light = QPalette()
    light.setColor(QPalette.ColorRole.Window,          _QC('#f0f2f7'))
    light.setColor(QPalette.ColorRole.WindowText,      _QC('#1a1a2e'))
    light.setColor(QPalette.ColorRole.Base,            _QC('#ffffff'))
    light.setColor(QPalette.ColorRole.AlternateBase,   _QC('#f4f6fa'))
    light.setColor(QPalette.ColorRole.Text,            _QC('#1a1a2e'))
    light.setColor(QPalette.ColorRole.BrightText,      _QC('#1a1a2e'))
    light.setColor(QPalette.ColorRole.Button,          _QC('#f4f6fa'))
    light.setColor(QPalette.ColorRole.ButtonText,      _QC('#1a1a2e'))
    light.setColor(QPalette.ColorRole.Highlight,       _QC('#3a5bd9'))
    light.setColor(QPalette.ColorRole.HighlightedText, _QC('#ffffff'))
    light.setColor(QPalette.ColorRole.ToolTipBase,     _QC('#ffffff'))
    light.setColor(QPalette.ColorRole.ToolTipText,     _QC('#1a1a2e'))
    light.setColor(QPalette.ColorRole.PlaceholderText, _QC('#9aa5c8'))
    app.setPalette(light)
    _icon = QIcon(_resource_path('logo.png'))
    app.setWindowIcon(_icon)
    window = MainWindow()
    window.setWindowTitle("Quote Generator")
    window.setWindowIcon(_icon)
    window.show()
    sys.exit(app.exec())
