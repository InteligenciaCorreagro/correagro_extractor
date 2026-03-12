"""
Generador de PDF: CARTERA (Landscape A4)
Un PDF por cliente: {NIT} CARTERA.pdf (en subcarpeta CARTERA/)
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
from core.models import ClienteCartera

PAGE_SIZE = landscape(A4)


class _CarteraCanvas(CorreagroCanvasBase):
    def __init__(self, filename, corte_str, cuenta, nit, razon_social, **kw):
        super().__init__(filename, **kw)
        self.corte_str = corte_str
        self.cuenta = cuenta
        self.nit = nit
        self.razon_social = razon_social

    def _draw_header_footer(self, total_pages):
        w, h = PAGE_SIZE

        self.draw_logo(30, h - 105, w=90, h=85)
        self.draw_vigilado_strip(20, h - 112, h=100)

        self.setFont("Helvetica-Bold", 14)
        self.setFillColor(C_DARK_BLUE)
        self.drawCentredString(w / 2, h - 52, "ESTADO DE CUENTA")
        self.setFont("Helvetica-Bold", 10)
        self.drawCentredString(w / 2, h - 67, "OPERACIONES DE MERCADO ABIERTO")

        self.setFont("Helvetica-Bold", 9)
        self.setFillColor(C_BLACK)
        self.drawString(145, h - 105, "CORTE")
        self.setFont("Helvetica", 9)
        self.drawString(200, h - 105, self.corte_str)

        if self.cuenta:
            self.setFont("Helvetica-Bold", 9)
            self.setFillColor(C_BLACK)
            self.drawString(145, h - 118, "CUENTA")
            self.setFont("Helvetica", 9)
            self.drawString(200, h - 118, self.cuenta)

        self.setFont("Helvetica-Bold", 9)
        self.setFillColor(C_DARK_BLUE)
        self.drawString(145, h - 140, "NIT Tercero")
        self.drawString(145, h - 153, "Nombre Tercero")
        self.setFont("Helvetica", 9)
        self.setFillColor(C_BLACK)
        self.drawString(235, h - 140, self.nit)
        self.drawString(235, h - 153, self.razon_social[:65])

        self.setStrokeColor(C_MID_BLUE)
        self.setLineWidth(1.2)
        self.line(30, h - 165, w - 30, h - 165)

        self.draw_footer_img(w, margin=30, bottom=15)
        self.draw_page_number(w, self._pageNumber, total_pages)


def _styles(compact=False):
    f_sz = 6.5 if compact else 7.5
    f_sz_h = 7.0 if compact else 8.0
    led = 7.5 if compact else 9.0
    led_h = 8.0 if compact else 10.0

    hdr = ParagraphStyle("H", fontSize=f_sz_h, fontName="Helvetica-Bold",
                         textColor=C_WHITE, alignment=TA_CENTER, leading=led_h)
    cel = ParagraphStyle("C", fontSize=f_sz, fontName="Helvetica",
                         textColor=C_BLACK, alignment=TA_LEFT, leading=led)
    cen = ParagraphStyle("CN", fontSize=f_sz, fontName="Helvetica",
                         textColor=C_BLACK, alignment=TA_CENTER, leading=led)
    cer = ParagraphStyle("CR", fontSize=f_sz, fontName="Helvetica",
                         textColor=C_BLACK, alignment=TA_RIGHT, leading=led)
    tot = ParagraphStyle("T", fontSize=f_sz_h, fontName="Helvetica-Bold",
                         textColor=C_WHITE, alignment=TA_RIGHT, leading=led_h)
    tl = ParagraphStyle("TL", fontSize=f_sz_h, fontName="Helvetica-Bold",
                        textColor=C_WHITE, alignment=TA_LEFT, leading=led_h)
    return hdr, cel, cen, cer, tot, tl


def generate_cartera_pdf(cliente: ClienteCartera, corte, cuenta: str, output_dir: str) -> str:
    out_dir = os.path.join(output_dir, "CARTERA")
    os.makedirs(out_dir, exist_ok=True)
    filename = os.path.join(out_dir, f"{cliente.nit} CARTERA.pdf")

    corte_str = fmt_date(corte) if corte else ""
    w, h = PAGE_SIZE
    usable_w = w - 60

    doc = SimpleDocTemplate(
        filename, pagesize=PAGE_SIZE,
        leftMargin=30, rightMargin=30,
        topMargin=175, bottomMargin=50,
    )

    compact = len(cliente.rows) > 35
    pad = 2 if compact else 4
    hdr_s, cel_s, cen_s, cer_s, tot_s, tl_s = _styles(compact)

    col_w = [
        usable_w * 0.16,
        usable_w * 0.10,
        usable_w * 0.10,
        usable_w * 0.05,
        usable_w * 0.10,
        usable_w * 0.10,
        usable_w * 0.10,
        usable_w * 0.10,
        usable_w * 0.10,
        usable_w * 0.09,
    ]

    header_row = [
        Paragraph("Documento", hdr_s),
        Paragraph("Fecha<br/>Contabilización", hdr_s),
        Paragraph("Fecha<br/>Vencimiento", hdr_s),
        Paragraph("Días", hdr_s),
        Paragraph("Corriente", hdr_s),
        Paragraph("De 1 a 30", hdr_s),
        Paragraph("De 31 a 60", hdr_s),
        Paragraph("De 61 a 90", hdr_s),
        Paragraph("De 91 a 9999", hdr_s),
        Paragraph("Total", hdr_s),
    ]

    data = [header_row]
    sum_corr = sum_1_30 = sum_31_60 = sum_61_90 = sum_91_9999 = sum_total = 0.0

    for r in cliente.rows:
        sum_corr += r.corriente
        sum_1_30 += r.de_1_30
        sum_31_60 += r.de_31_60
        sum_61_90 += r.de_61_90
        sum_91_9999 += r.de_91_9999
        sum_total += r.total
        data.append([
            Paragraph(r.documento, cel_s),
            Paragraph(r.fecha_contabilizacion, cen_s),
            Paragraph(r.fecha_vencimiento, cen_s),
            Paragraph(str(r.dias), cen_s),
            Paragraph(fmt_currency(r.corriente), cer_s),
            Paragraph(fmt_currency(r.de_1_30), cer_s),
            Paragraph(fmt_currency(r.de_31_60), cer_s),
            Paragraph(fmt_currency(r.de_61_90), cer_s),
            Paragraph(fmt_currency(r.de_91_9999), cer_s),
            Paragraph(fmt_currency(r.total), cer_s),
        ])

    last = len(data)
    data.append([
        Paragraph("Total general", tl_s),
        Paragraph("", tl_s),
        Paragraph("", tl_s),
        Paragraph("", tl_s),
        Paragraph(fmt_currency(sum_corr), tot_s),
        Paragraph(fmt_currency(sum_1_30), tot_s),
        Paragraph(fmt_currency(sum_31_60), tot_s),
        Paragraph(fmt_currency(sum_61_90), tot_s),
        Paragraph(fmt_currency(sum_91_9999), tot_s),
        Paragraph(fmt_currency(sum_total), tot_s),
    ])

    ts = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), C_DARK_BLUE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, last - 1), [C_LIGHT_BLUE, colors.white]),
        ("BACKGROUND", (0, last), (-1, last), C_DARK_BLUE),
        ("SPAN", (0, last), (3, last)),
        ("GRID", (0, 0), (-1, last - 1), 0.4, C_MID_BLUE),
        ("LINEABOVE", (0, last), (-1, last), 1.2, C_DARK_BLUE),
        ("TOPPADDING", (0, 0), (-1, -1), pad),
        ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ])

    table = Table(data, colWidths=col_w, repeatRows=1)
    table.setStyle(ts)

    story = [Spacer(1, 6), table]

    def _make_canvas(fname, **kw):
        return _CarteraCanvas(fname, corte_str, cuenta, cliente.nit, cliente.razon_social, **kw)

    doc.build(story, canvasmaker=_make_canvas)
    return filename

