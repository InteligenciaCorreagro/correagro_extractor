"""
CORREAGRO — Generador de Extractos PDF
Aplicación de escritorio moderna con PySide6
"""
import os
import sys

from core.version import __version__ as APP_VERSION

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QProgressBar, QFrame,
    QScrollArea, QCheckBox, QLineEdit, QSizePolicy, QSpacerItem,
    QStackedWidget, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QToolButton, QAbstractItemView, QGridLayout,
    QGraphicsDropShadowEffect,
)
from PySide6.QtCore import (
    Qt, QThread, Signal, QObject, QPropertyAnimation,
    QEasingCurve, QSize, QTimer,
)
from PySide6.QtGui import (
    QPixmap, QFont, QIcon, QPainter, QColor, QLinearGradient,
    QBrush, QPen, QFontDatabase,
)

# ── Paths ─────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
LOGO_PATH  = os.path.join(ASSETS_DIR, "logo.png")

# ── Paleta ────────────────────────────────────────────────────
DARK_BG    = "#0D1B2A"
CARD_BG    = "#132337"
ACCENT     = "#2E75B6"
ACCENT2    = "#1F4E79"
SUCCESS    = "#27AE60"
ERROR      = "#E74C3C"
WARNING    = "#F39C12"
TEXT_MAIN  = "#ECF0F1"
TEXT_SUB   = "#95A5A6"
BORDER     = "#1E3A5F"

STYLE_MAIN = f"""
QMainWindow, QWidget#root {{
    background: {DARK_BG};
}}
QWidget {{
    color: {TEXT_MAIN};
    font-family: 'Segoe UI', 'Arial', sans-serif;
}}
QScrollBar:vertical {{
    background: {CARD_BG};
    width: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:vertical {{
    background: {ACCENT};
    border-radius: 4px;
    min-height: 30px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
QScrollBar:horizontal {{
    background: {CARD_BG};
    height: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:horizontal {{
    background: {ACCENT};
    border-radius: 4px;
    min-width: 30px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0px;
}}
QTableWidget {{
    background: {CARD_BG};
    border: 1px solid {BORDER};
    border-radius: 8px;
    gridline-color: {BORDER};
    outline: none;
}}
QTableWidget::item {{
    padding: 6px 10px;
    border: none;
}}
QTableWidget::item:selected {{
    background: {ACCENT}44;
    color: {TEXT_MAIN};
}}
QHeaderView::section {{
    background: {ACCENT2};
    color: white;
    font-weight: bold;
    font-size: 11px;
    padding: 8px 10px;
    border: none;
    border-right: 1px solid {BORDER};
}}
QHeaderView::section:last {{
    border-right: none;
}}
QCheckBox {{
    spacing: 8px;
    font-size: 12px;
}}
QCheckBox::indicator {{
    width: 18px;
    height: 18px;
    border-radius: 4px;
    border: 2px solid {ACCENT};
    background: {DARK_BG};
}}
QCheckBox::indicator:checked {{
    background: {ACCENT};
    image: none;
}}
QCheckBox::indicator:checked::after {{
    content: "✓";
}}
QLineEdit {{
    background: {DARK_BG};
    border: 1px solid {BORDER};
    border-radius: 6px;
    padding: 6px 10px;
    color: {TEXT_MAIN};
    font-size: 12px;
}}
QLineEdit:focus {{
    border: 1px solid {ACCENT};
}}
QToolTip {{
    background: {CARD_BG};
    color: {TEXT_MAIN};
    border: 1px solid {ACCENT};
    padding: 4px 8px;
    border-radius: 4px;
}}
"""

BTN_PRIMARY = f"""
QPushButton {{
    background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
        stop:0 {ACCENT}, stop:1 {ACCENT2});
    color: white;
    border: none;
    border-radius: 8px;
    padding: 10px 22px;
    font-weight: bold;
    font-size: 13px;
    min-height: 38px;
}}
QPushButton:hover {{
    background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
        stop:0 #3D8FD1, stop:1 {ACCENT});
}}
QPushButton:pressed {{
    background: {ACCENT2};
}}
QPushButton:disabled {{
    background: #2A3A4A;
    color: {TEXT_SUB};
}}
"""

BTN_SECONDARY = f"""
QPushButton {{
    background: transparent;
    color: {ACCENT};
    border: 2px solid {ACCENT};
    border-radius: 8px;
    padding: 9px 20px;
    font-weight: bold;
    font-size: 12px;
    min-height: 36px;
}}
QPushButton:hover {{
    background: {ACCENT}22;
}}
QPushButton:pressed {{
    background: {ACCENT}44;
}}
QPushButton:disabled {{
    color: {TEXT_SUB};
    border-color: {BORDER};
}}
"""

BTN_SUCCESS = f"""
QPushButton {{
    background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
        stop:0 #2ECC71, stop:1 {SUCCESS});
    color: white;
    border: none;
    border-radius: 8px;
    padding: 10px 22px;
    font-weight: bold;
    font-size: 13px;
    min-height: 38px;
}}
QPushButton:hover {{
    background: #2ECC71;
}}
QPushButton:pressed {{
    background: #1E8449;
}}
QPushButton:disabled {{
    background: #2A3A4A;
    color: {TEXT_SUB};
}}
"""

BTN_SMALL = f"""
QPushButton {{
    background: {ACCENT}22;
    color: {ACCENT};
    border: 1px solid {ACCENT}66;
    border-radius: 5px;
    padding: 4px 10px;
    font-size: 11px;
    min-height: 26px;
}}
QPushButton:hover {{ background: {ACCENT}44; }}
"""


# ═══════════════════════════════════════════════════════════════
# Worker thread
# ═══════════════════════════════════════════════════════════════
class GeneratorWorker(QObject):
    progress   = Signal(int, int, str)   # current, total, message
    finished   = Signal(list, list)      # ok_paths, errors
    error      = Signal(str)

    def __init__(self, informe, cartera, selected_nits_op, selected_nits_fc, selected_nits_ca, output_dir):
        super().__init__()
        self.informe          = informe
        self.cartera          = cartera
        self.selected_nits_op = set(selected_nits_op)
        self.selected_nits_fc = set(selected_nits_fc)
        self.selected_nits_ca = set(selected_nits_ca)
        self.output_dir       = output_dir

    def run(self):
        try:
            from generators.op_vigentes     import generate_op_vigentes_pdf
            from generators.fisicos_compras import generate_fisicos_compras_pdf
            from generators.cartera         import generate_cartera_pdf

            os.makedirs(self.output_dir, exist_ok=True)
            ok_paths, errors = [], []

            tasks = []
            if self.informe:
                for c in self.informe.clientes_op_vigentes:
                    if c.nit in self.selected_nits_op:
                        tasks.append(("op", c))
                for c in self.informe.clientes_fisicos_compras:
                    if c.nit in self.selected_nits_fc:
                        tasks.append(("fc", c))
            if self.cartera:
                for c in self.cartera.clientes:
                    if c.nit in self.selected_nits_ca:
                        tasks.append(("ca", c))

            total = len(tasks)
            for i, (kind, cliente) in enumerate(tasks):
                if kind == "op":
                    label = f"{cliente.nit} OP-VIGENTES.pdf"
                elif kind == "fc":
                    label = f"{cliente.nit} EXTRACTOS.pdf"
                else:
                    label = f"{cliente.nit} CARTERA.pdf"
                self.progress.emit(i, total, f"Generando {label}…")
                try:
                    if kind == "op":
                        path = generate_op_vigentes_pdf(
                            cliente, self.informe.desde, self.informe.hasta,
                            self.output_dir)
                    elif kind == "fc":
                        path = generate_fisicos_compras_pdf(
                            cliente, self.informe.desde, self.informe.hasta,
                            self.output_dir)
                    else:
                        path = generate_cartera_pdf(
                            cliente, self.cartera.corte, self.cartera.cuenta,
                            self.output_dir)
                    ok_paths.append(path)
                except Exception as e:
                    errors.append((label, str(e)))

            self.progress.emit(total, total, "¡Completado!")
            self.finished.emit(ok_paths, errors)
        except Exception as e:
            import traceback
            self.error.emit(f"{e}\n\n{traceback.format_exc()}")


# ═══════════════════════════════════════════════════════════════
# Componentes UI reutilizables
# ═══════════════════════════════════════════════════════════════
class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setStyleSheet(f"""
            QFrame#card {{
                background: {CARD_BG};
                border: 1px solid {BORDER};
                border-radius: 12px;
            }}
        """)
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(18)
        shadow.setOffset(0, 4)
        shadow.setColor(QColor(0, 0, 0, 80))
        self.setGraphicsEffect(shadow)


class SectionTitle(QLabel):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet(f"""
            color: {TEXT_MAIN};
            font-size: 15px;
            font-weight: bold;
            padding: 0 0 6px 0;
        """)


class SubLabel(QLabel):
    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setStyleSheet(f"color: {TEXT_SUB}; font-size: 11px;")


class StatusBadge(QLabel):
    def __init__(self, text, color=ACCENT, parent=None):
        super().__init__(f"  {text}  ", parent)
        self.setStyleSheet(f"""
            background: {color}33;
            color: {color};
            border: 1px solid {color}66;
            border-radius: 10px;
            font-size: 11px;
            font-weight: bold;
            padding: 2px 8px;
        """)
        self.setAlignment(Qt.AlignCenter)


class Divider(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.HLine)
        self.setStyleSheet(f"color: {BORDER}; background: {BORDER}; max-height: 1px;")


class SelectableTable(QTableWidget):
    """Tabla con checkboxes en primera columna."""
    def __init__(self, headers, parent=None):
        super().__init__(parent)
        full_headers = ["", *headers]
        self.setColumnCount(len(full_headers))
        self.setHorizontalHeaderLabels(full_headers)
        self.verticalHeader().setVisible(False)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setAlternatingRowColors(True)
        self.setStyleSheet(self.styleSheet() + f"""
            QTableWidget {{
                alternate-background-color: {DARK_BG};
            }}
        """)
        hh = self.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.Fixed)
        self.setColumnWidth(0, 40)
        for i in range(1, len(full_headers)):
            hh.setSectionResizeMode(i, QHeaderView.Stretch)

    def populate(self, rows: list[dict]):
        """rows = list of {'nit': ..., 'nombre': ..., 'registros': ..., 'total': ...}"""
        self.setRowCount(len(rows))
        for r, row in enumerate(rows):
            # Checkbox
            chk = QCheckBox()
            chk.setChecked(True)
            chk.setProperty("nit", row["nit"])
            w = QWidget()
            l = QHBoxLayout(w)
            l.addWidget(chk)
            l.setAlignment(Qt.AlignCenter)
            l.setContentsMargins(0, 0, 0, 0)
            self.setCellWidget(r, 0, w)
            # Resto de columnas
            for c, key in enumerate(["nit", "nombre", "registros", "total"], 1):
                val = row.get(key, "")
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
                self.setItem(r, c, item)
        self.resizeRowsToContents()

    def get_selected_nits(self) -> list[str]:
        nits = []
        for r in range(self.rowCount()):
            w = self.cellWidget(r, 0)
            if w:
                chk = w.findChild(QCheckBox)
                if chk and chk.isChecked():
                    nits.append(chk.property("nit"))
        return nits

    def select_all(self, checked: bool):
        for r in range(self.rowCount()):
            w = self.cellWidget(r, 0)
            if w:
                chk = w.findChild(QCheckBox)
                if chk:
                    chk.setChecked(checked)


# ═══════════════════════════════════════════════════════════════
# Barra lateral
# ═══════════════════════════════════════════════════════════════
class SidebarButton(QPushButton):
    def __init__(self, icon_txt, label, parent=None):
        super().__init__(parent)
        self.setText(f"  {icon_txt}  {label}")
        self.setCheckable(True)
        self.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {TEXT_SUB};
                border: none;
                border-left: 3px solid transparent;
                border-radius: 0;
                padding: 12px 16px;
                font-size: 13px;
                text-align: left;
                min-height: 44px;
            }}
            QPushButton:hover {{
                background: {ACCENT}18;
                color: {TEXT_MAIN};
            }}
            QPushButton:checked {{
                background: {ACCENT}25;
                color: {ACCENT};
                border-left: 3px solid {ACCENT};
                font-weight: bold;
            }}
        """)


class Sidebar(QWidget):
    page_changed = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(220)
        self.setStyleSheet(f"background: {CARD_BG}; border-right: 1px solid {BORDER};")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # Logo
        logo_frame = QFrame()
        logo_frame.setFixedHeight(90)
        logo_frame.setStyleSheet(f"background: {ACCENT2}; border: none;")
        lf_lay = QHBoxLayout(logo_frame)
        lf_lay.setContentsMargins(14, 10, 14, 10)
        logo_lbl = QLabel()
        if os.path.exists(LOGO_PATH):
            pix = QPixmap(LOGO_PATH).scaled(55, 50, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_lbl.setPixmap(pix)
        title_lbl = QLabel("CORREAGRO\nExtractor")
        title_lbl.setStyleSheet("color: white; font-weight: bold; font-size: 13px; background: transparent;")
        lf_lay.addWidget(logo_lbl)
        lf_lay.addWidget(title_lbl, 1)
        lay.addWidget(logo_frame)

        lay.addSpacing(10)

        self.btns = []
        items = [
            ("📂", "Cargar archivo"),
            ("📄", "OP Vigentes"),
            ("📊", "Físicos Compras"),
            ("💼", "Cartera"),
            ("⚙️",  "Generar PDFs"),
        ]
        for i, (ico, lbl) in enumerate(items):
            btn = SidebarButton(ico, lbl)
            btn.clicked.connect(lambda checked, idx=i: self._on_click(idx))
            lay.addWidget(btn)
            self.btns.append(btn)

        lay.addStretch()

        # Versión
        ver = QLabel(f"{APP_VERSION}  ·  CORREAGRO S.A.")
        ver.setStyleSheet(f"color: {TEXT_SUB}; font-size: 10px; padding: 10px 16px;")
        lay.addWidget(ver)

        self.btns[0].setChecked(True)

    def _on_click(self, idx):
        for i, b in enumerate(self.btns):
            b.setChecked(i == idx)
        self.page_changed.emit(idx)

    def set_page(self, idx):
        self._on_click(idx)

    def enable_pages(self, enabled: bool):
        for b in self.btns[1:]:
            b.setEnabled(enabled)


# ═══════════════════════════════════════════════════════════════
# Página 0: Cargar archivo
# ═══════════════════════════════════════════════════════════════
class PageLoad(QWidget):
    files_loaded = Signal(str, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(32, 32, 32, 32)
        lay.setSpacing(20)

        lay.addWidget(SectionTitle("📂  Cargar archivos Excel"))
        lay.addWidget(SubLabel("Selecciona el archivo de extractos y/o el archivo de cartera."))
        lay.addWidget(Divider())
        lay.addSpacing(10)

        # Drop zone card (Extractos)
        card = Card()
        card.setMinimumHeight(220)
        card_lay = QVBoxLayout(card)
        card_lay.setAlignment(Qt.AlignCenter)
        card_lay.setSpacing(14)

        self.drop_icon = QLabel("📁")
        self.drop_icon.setStyleSheet("font-size: 52px;")
        self.drop_icon.setAlignment(Qt.AlignCenter)

        self.drop_label = QLabel("Arrastra aquí el Excel de EXTRACTOS\no haz clic para seleccionarlo")
        self.drop_label.setStyleSheet(f"color: {TEXT_SUB}; font-size: 14px;")
        self.drop_label.setAlignment(Qt.AlignCenter)

        self.btn_select = QPushButton("  Seleccionar EXTRACTOS…")
        self.btn_select.setStyleSheet(BTN_PRIMARY)
        self.btn_select.setMaximumWidth(240)
        self.btn_select.clicked.connect(self._browse)

        card_lay.addWidget(self.drop_icon)
        card_lay.addWidget(self.drop_label)
        card_lay.addWidget(self.btn_select, alignment=Qt.AlignCenter)
        lay.addWidget(card)

        # Info card (Extractos)
        info_card = Card()
        info_card.setMaximumHeight(90)
        info_lay = QHBoxLayout(info_card)
        info_lay.setContentsMargins(16, 10, 16, 10)

        self.file_icon = QLabel("—")
        self.file_icon.setStyleSheet(f"font-size: 26px; color: {ACCENT};")
        self.file_name = QLabel("EXTRACTOS: sin cargar")
        self.file_name.setStyleSheet(f"color: {TEXT_MAIN}; font-size: 13px;")
        self.file_status = StatusBadge("Sin cargar", WARNING)

        info_lay.addWidget(self.file_icon)
        info_lay.addWidget(self.file_name, 1)
        info_lay.addWidget(self.file_status)
        lay.addWidget(info_card)

        cartera_card = Card()
        cartera_card.setMaximumHeight(90)
        cartera_lay = QHBoxLayout(cartera_card)
        cartera_lay.setContentsMargins(16, 10, 16, 10)

        self.cartera_icon = QLabel("—")
        self.cartera_icon.setStyleSheet(f"font-size: 26px; color: {ACCENT};")
        self.cartera_name = QLabel("CARTERA: sin cargar")
        self.cartera_name.setStyleSheet(f"color: {TEXT_MAIN}; font-size: 13px;")
        self.cartera_status = StatusBadge("Sin cargar", WARNING)
        self.btn_select_cartera = QPushButton("  Seleccionar CARTERA…")
        self.btn_select_cartera.setStyleSheet(BTN_SECONDARY)
        self.btn_select_cartera.clicked.connect(self._browse_cartera)

        cartera_lay.addWidget(self.cartera_icon)
        cartera_lay.addWidget(self.cartera_name, 1)
        cartera_lay.addWidget(self.btn_select_cartera)
        cartera_lay.addWidget(self.cartera_status)
        lay.addWidget(cartera_card)

        # Botón continuar
        self.btn_continue = QPushButton("  Continuar →")
        self.btn_continue.setStyleSheet(BTN_SUCCESS)
        self.btn_continue.setMaximumWidth(200)
        self.btn_continue.setEnabled(False)
        self.btn_continue.clicked.connect(lambda: self.files_loaded.emit(self._path, self._cartera_path))
        lay.addWidget(self.btn_continue, alignment=Qt.AlignRight)
        lay.addStretch()

        self._path = ""
        self._cartera_path = ""
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        for url in e.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".xlsx"):
                self._set_file(path)
                break

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar informe Excel", "",
            "Excel Files (*.xlsx *.xls)")
        if path:
            self._set_file(path)

    def _browse_cartera(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar cartera Excel", "",
            "Excel Files (*.xlsx *.xls)")
        if path:
            self._set_cartera_file(path)

    def _set_file(self, path):
        self._path = path
        self.file_name.setText(os.path.basename(path))
        self.file_icon.setText("📄")
        self.file_status.setText("  Listo  ")
        self.file_status.setStyleSheet(f"""
            background: {SUCCESS}33; color: {SUCCESS};
            border: 1px solid {SUCCESS}66;
            border-radius: 10px; font-size: 11px;
            font-weight: bold; padding: 2px 8px;
        """)
        self.btn_continue.setEnabled(bool(self._path or self._cartera_path))
        self.drop_label.setText(f"✅  {os.path.basename(path)}")

    def _set_cartera_file(self, path):
        self._cartera_path = path
        self.cartera_name.setText(os.path.basename(path))
        self.cartera_icon.setText("📄")
        self.cartera_status.setText("  Listo  ")
        self.cartera_status.setStyleSheet(f"""
            background: {SUCCESS}33; color: {SUCCESS};
            border: 1px solid {SUCCESS}66;
            border-radius: 10px; font-size: 11px;
            font-weight: bold; padding: 2px 8px;
        """)
        self.btn_continue.setEnabled(bool(self._path or self._cartera_path))


# ═══════════════════════════════════════════════════════════════
# Página genérica de tabla de clientes
# ═══════════════════════════════════════════════════════════════
class PageClientTable(QWidget):
    def __init__(self, title: str, subtitle: str, icon: str,
                 table_headers: list, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(32, 32, 32, 32)
        lay.setSpacing(14)

        # Header row
        top = QHBoxLayout()
        lbl = SectionTitle(f"{icon}  {title}")
        self.badge = StatusBadge("0 clientes", ACCENT)
        top.addWidget(lbl)
        top.addStretch()
        top.addWidget(self.badge)
        lay.addLayout(top)
        lay.addWidget(SubLabel(subtitle))
        lay.addWidget(Divider())

        # Barra de búsqueda + botones selección
        bar = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("🔍  Filtrar por nombre o NIT…")
        self.search.setMaximumWidth(320)
        self.search.textChanged.connect(self._filter)

        self.btn_all  = QPushButton("☑  Todos")
        self.btn_all.setStyleSheet(BTN_SMALL)
        self.btn_none = QPushButton("☐  Ninguno")
        self.btn_none.setStyleSheet(BTN_SMALL)
        self.btn_all.clicked.connect(lambda: self.table.select_all(True))
        self.btn_none.clicked.connect(lambda: self.table.select_all(False))

        bar.addWidget(self.search)
        bar.addStretch()
        bar.addWidget(self.btn_all)
        bar.addWidget(self.btn_none)
        lay.addLayout(bar)

        # Tabla
        self.table = SelectableTable(table_headers)
        lay.addWidget(self.table, 1)

        self._rows_data = []

    def load(self, rows: list[dict]):
        self._rows_data = rows
        self.table.populate(rows)
        count = len(rows)
        self.badge.setText(f"  {count} cliente{'s' if count != 1 else ''}  ")

    def _filter(self, text):
        text = text.lower()
        for r in range(self.table.rowCount()):
            nit   = (self.table.item(r, 1) or QTableWidgetItem("")).text().lower()
            name  = (self.table.item(r, 2) or QTableWidgetItem("")).text().lower()
            match = text in nit or text in name
            self.table.setRowHidden(r, not match)

    def get_selected_nits(self) -> list[str]:
        return self.table.get_selected_nits()


# ═══════════════════════════════════════════════════════════════
# Página 3: Generar PDFs
# ═══════════════════════════════════════════════════════════════
class PageGenerate(QWidget):
    def __init__(self, app_ref, parent=None):
        super().__init__(parent)
        self.app_ref = app_ref
        self._worker  = None
        self._thread  = None

        lay = QVBoxLayout(self)
        lay.setContentsMargins(32, 32, 32, 32)
        lay.setSpacing(16)

        lay.addWidget(SectionTitle("⚙️  Generar PDFs"))
        lay.addWidget(SubLabel("Configura la carpeta de salida y genera los PDFs para los clientes seleccionados."))
        lay.addWidget(Divider())

        # Carpeta de salida
        out_card = Card()
        out_lay  = QHBoxLayout(out_card)
        out_lay.setContentsMargins(16, 12, 16, 12)

        out_lay.addWidget(QLabel("📁  Carpeta de salida:"))
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("Selecciona una carpeta…")
        self.out_edit.setText(os.path.join(os.path.expanduser("~"), "Desktop", "Extractos_CORREAGRO"))
        self.btn_browse_out = QPushButton("Examinar…")
        self.btn_browse_out.setStyleSheet(BTN_SMALL)
        self.btn_browse_out.clicked.connect(self._browse_out)
        out_lay.addWidget(self.out_edit, 1)
        out_lay.addWidget(self.btn_browse_out)
        lay.addWidget(out_card)

        # Resumen de selección
        self.summary_card = Card()
        sum_lay = QGridLayout(self.summary_card)
        sum_lay.setContentsMargins(20, 14, 20, 14)
        sum_lay.setHorizontalSpacing(40)
        sum_lay.setVerticalSpacing(8)

        def _stat(row, col, title, attr_name, color=ACCENT):
            lbl_t = QLabel(title)
            lbl_t.setStyleSheet(f"color: {TEXT_SUB}; font-size: 11px;")
            lbl_v = QLabel("0")
            lbl_v.setStyleSheet(f"color: {color}; font-size: 22px; font-weight: bold;")
            sum_lay.addWidget(lbl_t, row*2,   col)
            sum_lay.addWidget(lbl_v, row*2+1, col)
            setattr(self, attr_name, lbl_v)

        _stat(0, 0, "OP Vigentes seleccionados",     "lbl_op_count", ACCENT)
        _stat(0, 1, "Físicos Compras seleccionados", "lbl_fc_count", "#9B59B6")
        _stat(0, 2, "Cartera seleccionados",         "lbl_ca_count", "#F39C12")
        _stat(0, 3, "Total PDFs a generar",          "lbl_total",    SUCCESS)
        lay.addWidget(self.summary_card)

        # Barra de progreso
        prog_card = Card()
        prog_lay = QVBoxLayout(prog_card)
        prog_lay.setContentsMargins(20, 14, 20, 14)
        prog_lay.setSpacing(8)

        self.prog_label = QLabel("Listo para generar")
        self.prog_label.setStyleSheet(f"color: {TEXT_MAIN}; font-size: 12px;")
        self.progress = QProgressBar()
        self.progress.setFixedHeight(10)
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet(f"""
            QProgressBar {{
                background: {DARK_BG};
                border-radius: 5px;
                border: none;
            }}
            QProgressBar::chunk {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 {ACCENT}, stop:1 #27AE60);
                border-radius: 5px;
            }}
        """)
        self.progress.setValue(0)
        prog_lay.addWidget(self.prog_label)
        prog_lay.addWidget(self.progress)
        lay.addWidget(prog_card)

        # Log de resultados
        self.log_scroll = QScrollArea()
        self.log_scroll.setWidgetResizable(True)
        self.log_scroll.setMinimumHeight(140)
        self.log_scroll.setStyleSheet(f"""
            QScrollArea {{ border: 1px solid {BORDER}; border-radius: 8px; background: {DARK_BG}; }}
        """)
        self.log_widget = QWidget()
        self.log_lay    = QVBoxLayout(self.log_widget)
        self.log_lay.setAlignment(Qt.AlignTop)
        self.log_lay.setContentsMargins(12, 10, 12, 10)
        self.log_lay.setSpacing(4)
        self.log_scroll.setWidget(self.log_widget)
        lay.addWidget(self.log_scroll, 1)

        # Botones
        btn_row = QHBoxLayout()
        self.btn_open = QPushButton("  📂  Abrir carpeta")
        self.btn_open.setStyleSheet(BTN_SECONDARY)
        self.btn_open.setEnabled(False)
        self.btn_open.clicked.connect(self._open_folder)

        self.btn_generate = QPushButton("  🚀  Generar PDFs")
        self.btn_generate.setStyleSheet(BTN_SUCCESS)
        self.btn_generate.clicked.connect(self._generate)

        btn_row.addWidget(self.btn_open)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_generate)
        lay.addLayout(btn_row)

    def _browse_out(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de salida")
        if folder:
            self.out_edit.setText(folder)

    def update_summary(self, n_op, n_fc, n_ca):
        self.lbl_op_count.setText(str(n_op))
        self.lbl_fc_count.setText(str(n_fc))
        self.lbl_ca_count.setText(str(n_ca))
        self.lbl_total.setText(str(n_op + n_fc + n_ca))

    def _log(self, msg: str, color=TEXT_MAIN):
        lbl = QLabel(msg)
        lbl.setStyleSheet(f"color: {color}; font-size: 11px; font-family: 'Consolas', monospace;")
        lbl.setWordWrap(True)
        self.log_lay.addWidget(lbl)
        QTimer.singleShot(50, lambda: self.log_scroll.verticalScrollBar().setValue(
            self.log_scroll.verticalScrollBar().maximum()))

    def _clear_log(self):
        while self.log_lay.count():
            item = self.log_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def _generate(self):
        informe = self.app_ref.informe
        cartera = self.app_ref.cartera
        if not informe and not cartera:
            QMessageBox.warning(self, "Sin datos", "Primero carga un archivo Excel.")
            return

        nits_op = self.app_ref.page_op.get_selected_nits() if informe else []
        nits_fc = self.app_ref.page_fc.get_selected_nits() if informe else []
        nits_ca = self.app_ref.page_cartera.get_selected_nits() if cartera else []

        if not nits_op and not nits_fc and not nits_ca:
            QMessageBox.warning(self, "Sin selección",
                                "Selecciona al menos un cliente en OP Vigentes, Físicos Compras o Cartera.")
            return

        out_dir = self.out_edit.text().strip()
        if not out_dir:
            QMessageBox.warning(self, "Sin carpeta", "Especifica una carpeta de salida.")
            return

        self._clear_log()
        self._log(f"📁  Carpeta: {out_dir}", ACCENT)
        total_jobs = len(nits_op) + len(nits_fc) + len(nits_ca)
        self._log(f"▶  Iniciando generación de {total_jobs} PDFs…", TEXT_MAIN)
        self.progress.setValue(0)
        self.progress.setMaximum(total_jobs)
        self.btn_generate.setEnabled(False)
        self.btn_open.setEnabled(False)
        self._out_dir = out_dir

        self._worker = GeneratorWorker(informe, cartera, nits_op, nits_fc, nits_ca, out_dir)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._thread.start()

    def _on_progress(self, current, total, msg):
        self.progress.setValue(current)
        self.prog_label.setText(msg)
        self._log(f"  ✓  {msg}", SUCCESS)

    def _on_finished(self, ok_paths, errors):
        self._thread.quit()
        self._thread.wait()
        self.btn_generate.setEnabled(True)
        self.btn_open.setEnabled(True)
        self.progress.setValue(self.progress.maximum())
        self.prog_label.setText("¡Generación completada!")

        self._log(f"\n✅  {len(ok_paths)} PDF(s) generados correctamente.", SUCCESS)
        if errors:
            for label, err in errors:
                self._log(f"❌  Error en {label}: {err}", ERROR)
        self._log(f"📁  Guardados en: {self._out_dir}", ACCENT)

    def _on_error(self, msg):
        self._thread.quit()
        self.btn_generate.setEnabled(True)
        self._log(f"❌  Error crítico:\n{msg}", ERROR)

    def _open_folder(self):
        path = self.out_edit.text().strip()
        if path and os.path.exists(path):
            import subprocess, platform
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])


# ═══════════════════════════════════════════════════════════════
# Ventana principal
# ═══════════════════════════════════════════════════════════════
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"CORREAGRO — Generador de Extractos PDF ({APP_VERSION})")
        self.setMinimumSize(1050, 680)
        self.resize(1200, 760)
        self.setStyleSheet(STYLE_MAIN)
        if os.path.exists(LOGO_PATH):
            self.setWindowIcon(QIcon(LOGO_PATH))

        self.informe = None
        self.cartera = None

        # Root widget
        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)
        root_lay = QHBoxLayout(root)
        root_lay.setContentsMargins(0, 0, 0, 0)
        root_lay.setSpacing(0)

        # Sidebar
        self.sidebar = Sidebar()
        self.sidebar.page_changed.connect(self._switch_page)
        self.sidebar.enable_pages(False)
        root_lay.addWidget(self.sidebar)

        # Stack
        self.stack = QStackedWidget()
        root_lay.addWidget(self.stack, 1)

        # Pages
        self.page_load = PageLoad()
        self.page_load.files_loaded.connect(self._load_files)

        self.page_op = PageClientTable(
            "OP Vigentes", "Selecciona los clientes para generar sus extractos.",
            "📄",
            ["NIT", "Razón Social", "Registros", "Total Saldo"],
        )
        self.page_fc = PageClientTable(
            "Físicos Compras",
            "Selecciona los clientes para generar sus extractos de físicos.",
            "📊",
            ["NIT", "Razón Social", "Movimientos", "Saldo Final"],
        )
        self.page_cartera = PageClientTable(
            "Cartera",
            "Selecciona los clientes para generar su estado de cuenta.",
            "💼",
            ["NIT", "Razón Social", "Registros", "Total"],
        )
        self.page_gen = PageGenerate(self)

        for p in [self.page_load, self.page_op, self.page_fc, self.page_cartera, self.page_gen]:
            self.stack.addWidget(p)

        # Connect generate page summary updates
        QTimer.singleShot(100, self._connect_summary)
        QTimer.singleShot(1500, self._check_updates)

    def _connect_summary(self):
        # Actualizar resumen al cambiar a página de generar
        pass

    def _check_updates(self):
        if not str(APP_VERSION).startswith("production-"):
            return
        try:
            from core.updater import UpdateWorker
        except Exception:
            return

        self._upd_thread = QThread()
        self._upd_worker = UpdateWorker(APP_VERSION)
        self._upd_worker.moveToThread(self._upd_thread)
        self._upd_thread.started.connect(self._upd_worker.run)
        self._upd_worker.finished.connect(self._on_update_checked)
        self._upd_worker.finished.connect(self._upd_thread.quit)
        self._upd_worker.finished.connect(self._upd_worker.deleteLater)
        self._upd_thread.finished.connect(self._upd_thread.deleteLater)
        self._upd_thread.start()

    def _on_update_checked(self, release, error):
        if error or release is None:
            return
        try:
            from core.updater import DEFAULT_ASSET_NAME, download_file, run_installer
        except Exception:
            return

        tag = getattr(release, "tag", "")
        url = getattr(release, "url", "")
        asset_name = getattr(release, "asset_name", "") or DEFAULT_ASSET_NAME
        if not url:
            return

        res = QMessageBox.question(
            self,
            "Actualización disponible",
            f"Hay una nueva versión disponible ({tag}).\n\n¿Deseas descargarla e instalarla ahora?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if res != QMessageBox.Yes:
            return

        try:
            installer = download_file(url, asset_name)
            run_installer(installer)
            QApplication.quit()
        except Exception as e:
            QMessageBox.warning(self, "No se pudo actualizar", f"No se pudo descargar/instalar:\n\n{e}")

    def _switch_page(self, idx):
        self.stack.setCurrentIndex(idx)
        if idx == 4:
            n_op = len(self.page_op.get_selected_nits()) if self.informe else 0
            n_fc = len(self.page_fc.get_selected_nits()) if self.informe else 0
            n_ca = len(self.page_cartera.get_selected_nits()) if self.cartera else 0
            self.page_gen.update_summary(n_op, n_fc, n_ca)

    def _load_files(self, informe_path: str, cartera_path: str):
        from PySide6.QtWidgets import QProgressDialog
        dlg = QProgressDialog("Cargando archivo…", None, 0, 0, self)
        dlg.setWindowModality(Qt.ApplicationModal)
        dlg.setWindowTitle("Procesando")
        dlg.setMinimumDuration(0)
        dlg.setValue(0)
        dlg.show()
        QApplication.processEvents()

        try:
            sys.path.insert(0, BASE_DIR)
            if informe_path:
                from core.reader import parse_informe
                self.informe = parse_informe(informe_path)
                self._populate_tables()

            if cartera_path:
                from core.reader import parse_cartera
                self.cartera = parse_cartera(cartera_path)
                self._populate_cartera_table()

            self.sidebar.enable_pages(True)
            if self.informe:
                self.sidebar.set_page(1)
            elif self.cartera:
                self.sidebar.set_page(3)
            else:
                self.sidebar.set_page(0)

            self.sidebar.btns[1].setEnabled(bool(self.informe))
            self.sidebar.btns[2].setEnabled(bool(self.informe))
            self.sidebar.btns[3].setEnabled(bool(self.cartera))
            self.sidebar.btns[4].setEnabled(bool(self.informe or self.cartera))
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "Error al cargar",
                                 f"No se pudo leer el archivo:\n\n{e}\n\n{traceback.format_exc()}")
        finally:
            dlg.close()

    def _populate_tables(self):
        from generators.pdf_utils import fmt_currency

        # OP Vigentes
        rows_op = []
        for c in self.informe.clientes_op_vigentes:
            rows_op.append({
                "nit":       c.nit,
                "nombre":    c.razon_social,
                "registros": str(len(c.rows)),
                "total":     fmt_currency(c.total_saldo),
            })
        self.page_op.load(rows_op)

        # Físicos Compras
        rows_fc = []
        for c in self.informe.clientes_fisicos_compras:
            rows_fc.append({
                "nit":       c.nit,
                "nombre":    c.razon_social,
                "registros": str(len(c.rows)),
                "total":     fmt_currency(c.saldo_final),
            })
        self.page_fc.load(rows_fc)

    def _populate_cartera_table(self):
        from generators.pdf_utils import fmt_currency

        rows = []
        for c in self.cartera.clientes:
            rows.append({
                "nit": c.nit,
                "nombre": c.razon_social,
                "registros": str(len(c.rows)),
                "total": fmt_currency(c.total_general),
            })
        self.page_cartera.load(rows)


# ═══════════════════════════════════════════════════════════════
# Entry point
# ═══════════════════════════════════════════════════════════════
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("CORREAGRO Extractor")
    app.setOrganizationName("CORREAGRO S.A.")

    # Fuente global
    app.setFont(QFont("Segoe UI", 10))

    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
