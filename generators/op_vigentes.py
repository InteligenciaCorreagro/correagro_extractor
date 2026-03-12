"""
Generador de PDF: ANEXO EXTRACTO – OPERACIONES DE MERCADO ABIERTO (Portrait, A4)
Un PDF por cliente: {NIT} OP-VIGENTES.pdf (en subcarpeta OP-VIGENTES/)
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors

from generators.pdf_utils import (
    CorreagroCanvasBase, fmt_currency, fmt_date,
    C_DARK_BLUE, C_MID_BLUE, C_LIGHT_BLUE, C_WHITE, C_BLACK, C_GREY,
)
from core.models import ClienteOpVigente


# ── Canvas personalizado ──────────────────────────────────────
class _OPCanvas(CorreagroCanvasBase):
    def __init__(self, filename, desde_str, hasta_str, nit, razon_social, **kw):
        super().__init__(filename, **kw)
        self.desde_str = desde_str
        self.hasta_str = hasta_str
        self.nit       = nit
        self.razon_social = razon_social

    def _draw_header_footer(self, total_pages):
        w, h = A4

        self.draw_logo(30, h - 105, w=90, h=85)
        self.draw_vigilado_strip(20, h - 112, h=100)

        # Título centrado
        self.setFont("Helvetica-Bold", 14)
        self.setFillColor(C_DARK_BLUE)
        self.drawCentredString(w / 2, h - 58, "ANEXO EXTRACTO")
        self.setFont("Helvetica-Bold", 10)
        self.drawCentredString(w / 2, h - 73, "OPERACIONES DE MERCADO ABIERTO")

        # Período
        self.setFont("Helvetica-Bold", 9)
        self.setFillColor(C_BLACK)
        self.drawString(135, h - 118, "PERIODO DEL INFORME")
        self.setFont("Helvetica", 9)
        self.drawString(135, h - 131, "Desde:")
        self.drawString(190, h - 131, self.desde_str)
        self.drawString(135, h - 144, "Hasta:")
        self.drawString(190, h - 144, self.hasta_str)

        # Etiqueta cliente
        self.setFont("Helvetica-Bold", 9)
        self.setFillColor(C_DARK_BLUE)
        self.drawString(135, h - 162, "NIT Tercero")
        self.drawString(135, h - 175, "Nombre Tercero")
        self.setFont("Helvetica", 9)
        self.setFillColor(C_BLACK)
        self.drawString(210, h - 162, self.nit)
        self.drawString(210, h - 175, self.razon_social[:55])

        # Línea separadora
        self.setStrokeColor(C_MID_BLUE)
        self.setLineWidth(1.2)
        self.line(30, h - 186, w - 30, h - 186)

        self.draw_footer_img(w)
        self.draw_page_number(w, self._pageNumber, total_pages)


# ── Estilos de párrafo ────────────────────────────────────────
def _styles(compact=False):
    f_sz   = 6.5 if compact else 7.5
    f_sz_h = 7.0 if compact else 8.0
    led    = 7.5 if compact else 9.0
    led_h  = 8.0 if compact else 10.0

    hdr = ParagraphStyle("H", fontSize=f_sz_h, fontName="Helvetica-Bold",
                         textColor=C_WHITE, alignment=TA_CENTER, leading=led_h)
    cel = ParagraphStyle("C", fontSize=f_sz, fontName="Helvetica",
                         textColor=C_BLACK, alignment=TA_LEFT, leading=led)
    cer = ParagraphStyle("CR", fontSize=f_sz, fontName="Helvetica",
                         textColor=C_BLACK, alignment=TA_RIGHT, leading=led)
    tot = ParagraphStyle("T", fontSize=f_sz_h, fontName="Helvetica-Bold",
                         textColor=C_WHITE, alignment=TA_RIGHT, leading=led_h)
    tl  = ParagraphStyle("TL", fontSize=f_sz_h, fontName="Helvetica-Bold",
                         textColor=C_WHITE, alignment=TA_LEFT, leading=led_h)
    return hdr, cel, cer, tot, tl


def generate_op_vigentes_pdf(cliente: ClienteOpVigente, desde, hasta,
                              output_dir: str) -> str:
    out_dir    = os.path.join(output_dir, "OP-VIGENTES")
    os.makedirs(out_dir, exist_ok=True)
    filename   = os.path.join(out_dir, f"{cliente.nit} OP-VIGENTES.pdf")
    desde_str  = fmt_date(desde)
    hasta_str  = fmt_date(hasta)
    w, h       = A4
    usable_w   = w - 60

    doc = SimpleDocTemplate(
        filename, pagesize=A4,
        leftMargin=30, rightMargin=30,
        topMargin=195, bottomMargin=50,
    )

    # Compactar si hay más de 45 filas
    compact = len(cliente.rows) > 45
    pad = 2 if compact else 4
    hdr_s, cel_s, cer_s, tot_s, tl_s = _styles(compact)

    col_w = [
        usable_w * 0.15, usable_w * 0.12, usable_w * 0.12,
        usable_w * 0.33, usable_w * 0.12, usable_w * 0.16,
    ]

    header_row = [
        Paragraph("OPERACIONES<br/>VIGENTES", hdr_s),
        Paragraph("Documento<br/>Cruce", hdr_s),
        Paragraph("Fecha docto<br/>cruce", hdr_s),
        Paragraph("Razón Social", hdr_s),
        Paragraph("Fecha de<br/>Vencimiento", hdr_s),
        Paragraph("Suma de<br/>Saldo_final_2", hdr_s),
    ]

    data = [header_row]
    for row in cliente.rows:
        razon = row.razon_social[:58] + "…" if len(row.razon_social) > 60 else row.razon_social
        data.append([
            Paragraph(row.operacion, cel_s),
            Paragraph(str(row.doc_cruce), cel_s),
            Paragraph(row.fecha_doc, cel_s),
            Paragraph(razon, cel_s),
            Paragraph(row.fecha_vencimiento, cel_s),
            Paragraph(fmt_currency(row.saldo_final), cer_s),
        ])

    last = len(data)
    data.append([
        Paragraph("Total general", tl_s),
        Paragraph("", tl_s), Paragraph("", tl_s),
        Paragraph("", tl_s), Paragraph("", tl_s),
        Paragraph(fmt_currency(cliente.total_saldo), tot_s),
    ])

    ts = TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),    C_DARK_BLUE),
        ("VALIGN",        (0, 0), (-1, -1),   "MIDDLE"),
        ("ROWBACKGROUNDS",(0, 1), (-1, last-1),[C_LIGHT_BLUE, colors.white]),
        ("BACKGROUND",    (0, last),(-1, last), C_DARK_BLUE),
        ("SPAN",          (0, last),(4, last)),
        ("GRID",          (0, 0), (-1, last-1), 0.4, C_MID_BLUE),
        ("LINEABOVE",     (0, last),(-1, last), 1.2, C_DARK_BLUE),
        ("TOPPADDING",    (0, 0), (-1, -1), pad),
        ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
    ])

    table = Table(data, colWidths=col_w, repeatRows=1)
    table.setStyle(ts)

    def _make_canvas(fname, **kw):
        return _OPCanvas(fname, desde_str, hasta_str, cliente.nit, cliente.razon_social, **kw)

    doc.build([table], canvasmaker=_make_canvas)
    return filename
