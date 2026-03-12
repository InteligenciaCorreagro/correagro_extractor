"""
Utilidades compartidas para generación de PDFs CORREAGRO
"""
import os
import sys
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdfcanvas

# ── Paleta corporativa ────────────────────────────────────────
C_DARK_BLUE  = colors.HexColor("#1F4E79")
C_MID_BLUE   = colors.HexColor("#2E75B6")
C_LIGHT_BLUE = colors.HexColor("#EBF3FB")
C_WHITE      = colors.white
C_BLACK      = colors.black
C_GREY       = colors.HexColor("#666666")

_DEV_BASE   = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
_MEI_BASE   = getattr(sys, "_MEIPASS", None)
ASSETS_DIR  = os.path.join(_MEI_BASE, "assets") if _MEI_BASE else os.path.join(_DEV_BASE, "assets")
LOGO_PATH   = os.path.join(ASSETS_DIR, "logo.png")
FOOTER_PATH = os.path.join(ASSETS_DIR, "footer.png")


def fmt_currency(value) -> str:
    """Formato COP: 1.055.000.402,38"""
    if value is None:
        return ""
    try:
        f = float(value)
        parts = f"{abs(f):,.2f}".split(".")
        int_part = parts[0].replace(",", ".")
        result   = f"{int_part},{parts[1]}"
        return f"-{result}" if f < 0 else result
    except Exception:
        return str(value)


def fmt_date(val) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    if len(s) == 8 and s.isdigit():
        return f"{s[6:8]}/{s[4:6]}/{s[0:4]}"
    return s


# ── Canvas base reutilizable ──────────────────────────────────
class CorreagroCanvasBase(pdfcanvas.Canvas):
    """
    Canvas que dibuja header y footer corporativo en cada página.
    Subclases deben implementar `_draw_header_footer(total_pages)`.
    """
    def __init__(self, filename, **kwargs):
        super().__init__(filename, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        n = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_header_footer(n)
            super().showPage()
        super().save()

    def _draw_header_footer(self, total_pages: int):
        raise NotImplementedError

    # ── Helpers de dibujo ─────────────────────────────────────
    def draw_logo(self, x, y, w=90, h=80):
        if os.path.exists(LOGO_PATH):
            self.drawImage(LOGO_PATH, x, y, width=w, height=h,
                           preserveAspectRatio=True, mask="auto")

    def draw_footer_img(self, page_width, margin=30, bottom=18):
        if os.path.exists(FOOTER_PATH):
            fw = page_width - margin * 2
            fh = fw * (67 / 1147)
            self.drawImage(FOOTER_PATH, margin, bottom, width=fw, height=fh,
                           preserveAspectRatio=True, mask="auto")

    def draw_vigilado_strip(self, x, y, h=100):
        self.saveState()
        self.setFillColor(C_DARK_BLUE)
        self.rect(x, y, 8, h, fill=1, stroke=0)
        self.setFillColor(C_WHITE)
        self.setFont("Helvetica-Bold", 5)
        self.translate(x + 4, y + h / 2)
        self.rotate(90)
        self.drawCentredString(0, 0, "VIGILADO")
        self.restoreState()

    def draw_page_number(self, page_width, page_num, total):
        self.setFont("Helvetica", 7)
        self.setFillColor(C_GREY)
        self.drawRightString(page_width - 30, 10,
                             f"Página {page_num} de {total}")
