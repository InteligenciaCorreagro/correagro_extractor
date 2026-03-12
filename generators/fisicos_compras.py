"""
Generador de PDF: EXTRACTOS – FISICOS COMPRAS (Landscape A4)
Un PDF por cliente: {NIT} EXTRACTOS.pdf (en subcarpeta EXTRACTOS/)
"""
import os
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors

from generators.pdf_utils import (
    CorreagroCanvasBase, fmt_currency, fmt_date,
    C_DARK_BLUE, C_MID_BLUE, C_LIGHT_BLUE, C_WHITE, C_BLACK, C_GREY,
)
from core.models import ClienteFisicoCompra

PAGE_SIZE = landscape(A4)   # 841.89 × 595.27 pts


# ── Canvas ────────────────────────────────────────────────────
class _FCCanvas(CorreagroCanvasBase):
    def __init__(self, filename, desde_str, hasta_str, nit, razon_social, **kw):
        super().__init__(filename, **kw)
        self.desde_str = desde_str
        self.hasta_str = hasta_str
        self.nit       = nit
        self.razon_social = razon_social

    def _draw_header_footer(self, total_pages):
        w, h = PAGE_SIZE

        self.draw_logo(30, h - 105, w=90, h=85)
        self.draw_vigilado_strip(20, h - 112, h=100)

        # Título
        self.setFont("Helvetica-Bold", 14)
        self.setFillColor(C_DARK_BLUE)
        self.drawCentredString(w / 2, h - 52, "EXTRACTOS")
        self.setFont("Helvetica-Bold", 10)
        self.drawCentredString(w / 2, h - 67, "OPERACIONES DE MERCADO ABIERTO")

        # Período
        self.setFont("Helvetica-Bold", 9)
        self.setFillColor(C_BLACK)
        self.drawString(145, h - 105, "PERIODO DEL INFORME")
        self.setFont("Helvetica", 9)
        self.drawString(145, h - 118, "Desde:")
        self.drawString(200, h - 118, self.desde_str)
        self.drawString(145, h - 131, "Hasta:")
        self.drawString(200, h - 131, self.hasta_str)

        # Cliente
        self.setFont("Helvetica-Bold", 9)
        self.setFillColor(C_DARK_BLUE)
        self.drawString(145, h - 149, "NIT Tercero")
        self.drawString(145, h - 162, "Nombre Tercero")
        self.setFont("Helvetica", 9)
        self.setFillColor(C_BLACK)
        self.drawString(235, h - 149, self.nit)
        self.drawString(235, h - 162, self.razon_social[:65])

        # Línea separadora
        self.setStrokeColor(C_MID_BLUE)
        self.setLineWidth(1.2)
        self.line(30, h - 171, w - 30, h - 171)

        self.draw_footer_img(w, margin=30, bottom=15)
        self.draw_page_number(w, self._pageNumber, total_pages)


# ── Estilos ───────────────────────────────────────────────────
def _styles(compact=False):
    f_sz   = 6.5 if compact else 7.5
    f_sz_h = 7.0 if compact else 8.0
    led    = 7.5 if compact else 9.0
    led_h  = 8.0 if compact else 10.0

    hdr = ParagraphStyle("H",  fontSize=f_sz_h, fontName="Helvetica-Bold",
                          textColor=C_WHITE,  alignment=TA_CENTER, leading=led_h)
    cel = ParagraphStyle("C",  fontSize=f_sz,   fontName="Helvetica",
                          textColor=C_BLACK,  alignment=TA_LEFT,   leading=led)
    cer = ParagraphStyle("CR", fontSize=f_sz,   fontName="Helvetica",
                          textColor=C_BLACK,  alignment=TA_RIGHT,  leading=led)
    tot = ParagraphStyle("T",  fontSize=f_sz_h, fontName="Helvetica-Bold",
                          textColor=C_WHITE,  alignment=TA_RIGHT,  leading=led_h)
    tl  = ParagraphStyle("TL", fontSize=f_sz_h, fontName="Helvetica-Bold",
                          textColor=C_WHITE,  alignment=TA_LEFT,   leading=led_h)
    sal = ParagraphStyle("SL", fontSize=8,    fontName="Helvetica-Bold",
                          textColor=C_DARK_BLUE, alignment=TA_LEFT, leading=11)
    sal_v = ParagraphStyle("SV", fontSize=8,  fontName="Helvetica",
                            textColor=C_BLACK, alignment=TA_RIGHT,  leading=11)
    return hdr, cel, cer, tot, tl, sal, sal_v


def generate_fisicos_compras_pdf(cliente: ClienteFisicoCompra, desde, hasta,
                                  output_dir: str) -> str:
    out_dir   = os.path.join(output_dir, "EXTRACTOS")
    os.makedirs(out_dir, exist_ok=True)
    filename  = os.path.join(out_dir, f"{cliente.nit} EXTRACTOS.pdf")
    desde_str = fmt_date(desde)
    hasta_str = fmt_date(hasta)
    w, h      = PAGE_SIZE
    usable_w  = w - 60

    doc = SimpleDocTemplate(
        filename, pagesize=PAGE_SIZE,
        leftMargin=30, rightMargin=30,
        topMargin=175, bottomMargin=50,
    )

    # Lógica de compactación: si hay más de 35 filas, reducir fuente
    compact = len(cliente.rows) > 35
    hdr_s, cel_s, cer_s, tot_s, tl_s, sal_s, sal_v_s = _styles(compact)
    pad = 2 if compact else 4

    # ── Tabla de saldos ───────────────────────────────────────
    saldo_data = [
        [Paragraph("SALDO INICIAL",  sal_s), Paragraph(fmt_currency(cliente.saldo_inicial), sal_v_s),
         Paragraph("",               sal_s), Paragraph("",                                  sal_v_s)],
        [Paragraph("ENTRADAS",       sal_s), Paragraph(fmt_currency(cliente.entradas),      sal_v_s),
         Paragraph("SALIDAS",        sal_s), Paragraph(fmt_currency(cliente.salidas),       sal_v_s)],
        [Paragraph("SALDO FINAL",    sal_s), Paragraph(fmt_currency(cliente.saldo_final),   sal_v_s),
         Paragraph("",               sal_s), Paragraph("",                                  sal_v_s)],
    ]
    sw = usable_w / 4
    saldo_table = Table(saldo_data, colWidths=[sw * 0.6, sw * 1.4, sw * 0.6, sw * 1.4])
    saldo_table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), colors.HexColor("#F0F7FF")),
        ("BOX",           (0, 0), (-1, -1), 0.8, C_MID_BLUE),
        ("INNERGRID",     (0, 0), (-1, -1), 0.4, C_MID_BLUE),
        ("TOPPADDING",    (0, 0), (-1, -1), pad),
        ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
    ]))

    # ── Tabla principal de movimientos ────────────────────────
    col_w = [
        usable_w * 0.09,  # Fecha_docto
        usable_w * 0.09,  # Fecha_vcto
        usable_w * 0.04,  # Docto
        usable_w * 0.06,  # Num.
        usable_w * 0.44,  # Notas
        usable_w * 0.14,  # Débitos
        usable_w * 0.14,  # Créditos
    ]

    header_row = [
        Paragraph("Fecha Docto",     hdr_s),
        Paragraph("Fecha Vcto",      hdr_s),
        Paragraph("Docto",           hdr_s),
        Paragraph("Núm.",            hdr_s),
        Paragraph("Notas",           hdr_s),
        Paragraph("Débitos",         hdr_s),
        Paragraph("Créditos",        hdr_s),
    ]

    data = [header_row]
    for row in cliente.rows:
        notas_str = row.notas[:110] + "…" if len(row.notas) > 112 else row.notas
        data.append([
            Paragraph(row.fecha_docto, cel_s),
            Paragraph(row.fecha_vcto,  cel_s),
            Paragraph(row.docto,       cel_s),
            Paragraph(str(row.num),    cel_s),
            Paragraph(notas_str,       cel_s),
            Paragraph(fmt_currency(row.debitos),  cer_s),
            Paragraph(fmt_currency(row.creditos), cer_s),
        ])

    last = len(data)
    data.append([
        Paragraph("Total general", tl_s),
        Paragraph("", tl_s), Paragraph("", tl_s),
        Paragraph("", tl_s), Paragraph("", tl_s),
        Paragraph(fmt_currency(cliente.total_debitos),  tot_s),
        Paragraph(fmt_currency(cliente.total_creditos), tot_s),
    ])

    ts = TableStyle([
        ("BACKGROUND",     (0, 0), (-1, 0),    C_DARK_BLUE),
        ("VALIGN",         (0, 0), (-1, -1),   "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, last-1),[C_LIGHT_BLUE, colors.white]),
        ("BACKGROUND",     (0, last),(-1, last), C_DARK_BLUE),
        ("SPAN",           (0, last),(4, last)),
        ("GRID",           (0, 0), (-1, last-1), 0.4, C_MID_BLUE),
        ("LINEABOVE",      (0, last),(-1, last), 1.2, C_DARK_BLUE),
        ("TOPPADDING",     (0, 0), (-1, -1), pad),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), pad),
        ("LEFTPADDING",    (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",   (0, 0), (-1, -1), 4),
    ])

    table = Table(data, colWidths=col_w, repeatRows=1)
    table.setStyle(ts)

    story = [saldo_table, Spacer(1, 10), table]

    def _make_canvas(fname, **kw):
        return _FCCanvas(fname, desde_str, hasta_str, cliente.nit, cliente.razon_social, **kw)

    doc.build(story, canvasmaker=_make_canvas)
    return filename
