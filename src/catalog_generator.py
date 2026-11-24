"""Generador de catálogos en PDF (ReportLab) y Excel (openpyxl)."""
import io
from datetime import datetime
from typing import Optional, List, Dict, Any, Union

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.pdfbase.pdfmetrics import stringWidth

INK = colors.HexColor("#1B2530")
MUTED = colors.HexColor("#66727D")
GRID = colors.HexColor("#C9D0D6")
TINT = colors.HexColor("#F2F5F7")

def parse_price(raw: Any) -> Optional[float]:
    """Normaliza y convierte strings numéricos y de moneda a float."""
    s = "".join(ch for ch in str(raw or "") if ch.isdigit() or ch in ".,-")
    if not any(ch.isdigit() for ch in s):
        return None
    last_comma, last_dot = s.rfind(","), s.rfind(".")
    if last_comma > -1 and last_dot > -1:
        if last_comma > last_dot:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif last_comma > -1:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) <= 2:
            s = parts[0] + "." + parts[1]
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None

def fmt_price(item: Dict[str, Any], currency: str) -> str:
    """Formatea el precio con la moneda correspondiente."""
    if item.get("price") is not None:
        txt = f"{item['price']:,.2f}"
        return f"{currency} {txt}".strip()
    return item.get("price_raw") or "-"

def text_color_on(hex_color: str):
    """Devuelve blanco o negro según el brillo del color de fondo."""
    c = colors.HexColor(hex_color)
    lum = 0.299 * c.red + 0.587 * c.green + 0.114 * c.blue
    return colors.black if lum > 0.62 else colors.white

def wrap_lines(text: str, font: str, size: float, max_w: float, max_lines: Optional[int] = None) -> List[str]:
    """Corta el texto en líneas calculando el ancho tipográfico real."""
    words = str(text or "").split()
    lines, cur = [], ""
    for w in words:
        test = f"{cur} {w}".strip()
        if stringWidth(test, font, size) <= max_w:
            cur = test
            continue
        if cur:
            lines.append(cur)
        while stringWidth(w, font, size) > max_w and len(w) > 1:
            k = len(w)
            while k > 1 and stringWidth(w[:k], font, size) > max_w:
                k -= 1
            lines.append(w[:k])
            w = w[k:]
        cur = w
    if cur:
        lines.append(cur)
    if not lines:
        lines = [""]
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines]
        last = lines[-1]
        while last and stringWidth(last + "...", font, size) > max_w:
            last = last[:-1]
        lines[-1] = last.rstrip() + "..."
    return lines

def draw_lines(c: canvas.Canvas, lines: List[str], x: float, y: float, lh: float, align: str = "left"):
    for j, ln in enumerate(lines):
        yy = y - j * lh
        if align == "center":
            c.drawCentredString(x, yy, ln)
        elif align == "right":
            c.drawRightString(x, yy, ln)
        else:
            c.drawString(x, yy, ln)

def draw_thumb(c: canvas.Canvas, png: bytes, x: float, y: float, w: float, h: float):
    img = ImageReader(io.BytesIO(png))
    c.drawImage(img, x, y, w, h, preserveAspectRatio=True, anchor="c", mask="auto")

def build_pdf(items: List[Dict[str, Any]], out_target: Union[str, io.BytesIO], o: Dict[str, Any]):
    """Genera el catálogo en formato PDF usando ReportLab (Tabla o Cuadrícula)."""
    c = canvas.Canvas(out_target, pagesize=A4)
    W, H = A4
    M, BOTTOM = 36, 44
    usable = W - 2 * M
    accent = colors.HexColor(o["accent"])
    on_accent = text_color_on(o["accent"])
    show_extra = o["show_extra"] and o["has_extra"]
    date_txt = datetime.now().strftime("%d/%m/%Y")

    def page_header(first: bool, page_no: int) -> float:
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 8)
        c.drawRightString(W - M, 24, f"Página {page_no}")
        if first:
            c.setFillColor(INK)
            c.setFont("Helvetica-Bold", 22)
            c.drawString(M, H - M - 16, wrap_lines(o["title"] or "Catálogo", "Helvetica-Bold", 22, usable - 150, 1)[0])
            if o["subtitle"]:
                c.setFillColor(MUTED)
                c.setFont("Helvetica", 10)
                c.drawString(M, H - M - 33, wrap_lines(o["subtitle"], "Helvetica", 10, usable - 150, 1)[0])
            c.setFillColor(MUTED)
            c.setFont("Helvetica", 9)
            c.drawRightString(W - M, H - M - 8, date_txt)
            c.drawRightString(W - M, H - M - 21, f"{len(items)} productos")
            c.setFillColor(accent)
            c.rect(M, H - M - 46, usable, 2, fill=1, stroke=0)
            return H - M - 62
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 8)
        c.drawString(M, H - M - 8, o["title"] or "Catálogo")
        c.setFillColor(GRID)
        c.rect(M, H - M - 14, usable, 0.75, fill=1, stroke=0)
        return H - M - 28

    def pdf_table():
        row_h, head_h = 112, 24
        col_sku, col_img = 86, 118
        col_price = 84 if o["show_price"] else 0
        col_desc = usable - col_sku - col_img - col_price
        x0 = M
        x1 = x0 + col_sku
        x2 = x1 + col_img
        x3 = x2 + col_desc
        i, page = 0, 1
        while True:
            y = page_header(page == 1, page)
            per = max(1, int((y - head_h - BOTTOM) // row_h))

            # Fila de encabezado
            c.setFillColor(accent)
            c.rect(M, y - head_h, usable, head_h, fill=1, stroke=0)
            c.setFillColor(on_accent)
            c.setFont("Helvetica-Bold", 10)
            c.drawCentredString(x0 + col_sku / 2, y - 16, "SKU")
            c.drawCentredString(x1 + col_img / 2, y - 16, "Foto")
            c.drawString(x2 + 10, y - 16, "Descripción")
            if o["show_price"]:
                c.drawCentredString(x3 + col_price / 2, y - 16, "Precio")

            y -= head_h
            for k, it in enumerate(items[i:i + per]):
                yt = y - k * row_h
                yb = yt - row_h
                c.setStrokeColor(GRID)
                c.setLineWidth(0.75)
                c.rect(x0, yb, col_sku, row_h, fill=0, stroke=1)
                c.rect(x1, yb, col_img, row_h, fill=0, stroke=1)
                c.rect(x2, yb, col_desc, row_h, fill=0, stroke=1)
                if o["show_price"]:
                    c.rect(x3, yb, col_price, row_h, fill=0, stroke=1)

                # SKU
                c.setFillColor(INK)
                c.setFont("Helvetica-Bold", 9.5)
                ls = wrap_lines(it["sku"], "Helvetica-Bold", 9.5, col_sku - 10, 2)
                draw_lines(c, ls, x0 + col_sku / 2, yt - row_h / 2 + (len(ls) - 1) * 6 - 3, 12, "center")

                # Foto
                bx, by = x1 + (col_img - 88) / 2, yb + (row_h - 88) / 2
                if it.get("thumb_png"):
                    try:
                        draw_thumb(c, it["thumb_png"], bx, by, 88, 88)
                    except Exception:
                        it["thumb_png"] = None
                if not it.get("thumb_png"):
                    c.setDash(3, 2)
                    c.rect(bx, by, 88, 88, fill=0, stroke=1)
                    c.setDash()
                    c.setFillColor(MUTED)
                    c.setFont("Helvetica-Oblique", 8.5)
                    c.drawCentredString(x1 + col_img / 2, yb + row_h / 2 - 3, "Sin imagen")

                # Descripción
                c.setFillColor(INK)
                c.setFont("Helvetica", 10)
                draw_lines(c, wrap_lines(it["desc"], "Helvetica", 10, col_desc - 20, 5), x2 + 10, yt - 20, 13)
                if show_extra and it.get("extra"):
                    c.setFillColor(MUTED)
                    c.setFont("Helvetica", 8.5)
                    c.drawString(x2 + 10, yb + 10, wrap_lines(it["extra"], "Helvetica", 8.5, col_desc - 20, 1)[0])

                # Precio
                if o["show_price"]:
                    c.setFillColor(INK)
                    c.setFont("Helvetica-Bold", 11)
                    lp = wrap_lines(fmt_price(it, o["currency"]), "Helvetica-Bold", 11, col_price - 8, 2)
                    draw_lines(c, lp, x3 + col_price / 2, yt - row_h / 2 + (len(lp) - 1) * 7 - 4, 13, "center")

            i += per
            if i >= len(items):
                break
            c.showPage()
            page += 1

    def pdf_grid():
        cols = o["grid_cols"]
        gap = 14
        cw = (usable - gap * (cols - 1)) / cols
        ih = cw * 0.8
        ch = ih + 78
        i, page = 0, 1
        while True:
            y = page_header(page == 1, page)
            rows = max(1, int((y - BOTTOM + gap) // (ch + gap)))
            per = rows * cols
            for k, it in enumerate(items[i:i + per]):
                col, row = k % cols, k // cols
                x = M + col * (cw + gap)
                yt = y - row * (ch + gap)
                yb = yt - ch

                c.setStrokeColor(GRID)
                c.setLineWidth(0.75)
                c.rect(x, yb, cw, ch, fill=0, stroke=1)
                c.setFillColor(TINT)
                c.rect(x + 0.4, yt - ih, cw - 0.8, ih - 0.4, fill=1, stroke=0)

                if it.get("thumb_png"):
                    try:
                        draw_thumb(c, it["thumb_png"], x + 4, yt - ih + 4, cw - 8, ih - 8)
                    except Exception:
                        it["thumb_png"] = None
                if not it.get("thumb_png"):
                    c.setFillColor(MUTED)
                    c.setFont("Helvetica-Oblique", 8.5)
                    c.drawCentredString(x + cw / 2, yt - ih / 2 - 3, "Sin imagen")

                with_extra = show_extra and it.get("extra") and cols < 4
                c.setFillColor(accent)
                c.setFont("Helvetica-Bold", 8.5)
                sku_w = cw * 0.5 if with_extra else cw - 16
                c.drawString(x + 8, yt - ih - 14, wrap_lines(it["sku"], "Helvetica-Bold", 8.5, sku_w, 1)[0])
                if with_extra:
                    c.setFillColor(MUTED)
                    c.setFont("Helvetica", 7.5)
                    c.drawRightString(x + cw - 8, yt - ih - 14, wrap_lines(it["extra"], "Helvetica", 7.5, cw * 0.4, 1)[0])

                c.setFillColor(INK)
                c.setFont("Helvetica", 9)
                draw_lines(c, wrap_lines(it["desc"], "Helvetica", 9, cw - 16, 3), x + 8, yt - ih - 28, 11)

                if o["show_price"]:
                    c.setFillColor(INK)
                    c.setFont("Helvetica-Bold", 12)
                    c.drawString(x + 8, yb + 9, wrap_lines(fmt_price(it, o["currency"]), "Helvetica-Bold", 12, cw - 16, 1)[0])

            i += per
            if i >= len(items):
                break
            c.showPage()
            page += 1

    if o["template"] == "Cuadrícula":
        pdf_grid()
    else:
        pdf_table()
    c.save()

def build_xlsx(items: List[Dict[str, Any]], out_target: Union[str, io.BytesIO], o: Dict[str, Any]):
    """Genera el catálogo en formato Excel con imágenes embebidas usando openpyxl."""
    show_extra = o["show_extra"] and o["has_extra"]
    show_price = o["show_price"]

    headers = ["SKU", "Foto", "Descripción"]
    widths = [18, 24, 60]
    if show_extra:
        headers.append("Categoría")
        widths.append(20)
    if show_price:
        headers.append("Precio")
        widths.append(16)

    wb = Workbook()
    ws = wb.active
    ws.title = "Catálogo"
    ws.append(headers)

    thin = Side(style="thin", color="C9D0D6")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    accent_hex = o["accent"].lstrip("#").upper()
    head_font = "000000" if text_color_on(o["accent"]) == colors.black else "FFFFFF"

    for ci in range(1, len(headers) + 1):
        cell = ws.cell(1, ci)
        cell.font = Font(bold=True, color=head_font)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor=accent_hex)
        cell.border = border
        ws.column_dimensions[chr(64 + ci)].width = widths[ci - 1]
    ws.row_dimensions[1].height = 24

    cur = o["currency"].replace('"', "")
    price_fmt = f'"{cur} "#,##0.00' if cur else "#,##0.00"

    for ri, it in enumerate(items, start=2):
        ws.cell(ri, 1, it["sku"]).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(ri, 3, it["desc"]).alignment = Alignment(wrap_text=True, vertical="center")
        next_col = 4
        if show_extra:
            ws.cell(ri, next_col, it.get("extra", "")).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            next_col += 1
        if show_price:
            if it.get("price") is not None:
                cp = ws.cell(ri, next_col, it["price"])
                cp.number_format = price_fmt
            else:
                cp = ws.cell(ri, next_col, it.get("price_raw") or None)
            cp.alignment = Alignment(horizontal="center", vertical="center")

        if it.get("thumb_png"):
            try:
                xl = XLImage(io.BytesIO(it["thumb_png"]))
                s = min(150 / xl.width, 150 / xl.height)
                xl.width, xl.height = int(xl.width * s), int(xl.height * s)
                ws.add_image(xl, f"B{ri}")
                ws.row_dimensions[ri].height = xl.height * 0.75 + 10
            except Exception:
                ws.cell(ri, 2, "(error imagen)").alignment = Alignment(horizontal="center", vertical="center")
                ws.row_dimensions[ri].height = 22
        else:
            ws.cell(ri, 2, "(sin imagen)").alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[ri].height = 22

        for ci in range(1, len(headers) + 1):
            ws.cell(ri, ci).border = border

    ws.freeze_panes = "A2"
    wb.save(out_target)
