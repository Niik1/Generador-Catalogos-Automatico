# =========================================================
#   GENERADOR DE CATÁLOGOS – PDF + EXCEL  (Google Colab)
#   Plantillas: Tabla / Cuadrícula
#   Opciones: título, subtítulo, moneda, precio, categoría/marca, color
#
#   Cómo usarlo:
#   1) Guarda en Colab Secrets: DROPBOX_APP_KEY, DROPBOX_APP_SECRET,
#      DROPBOX_REFRESH_TOKEN.
#   2) Edita la sección CONFIG EDITABLE (URL del Sheet, columnas, carpeta).
#   3) Ejecuta la celda, sube un Excel con los SKUs en la primera
#      columna, elige el diseño y pulsa "Generar catálogo".
# =========================================================
!pip -q install pandas openpyxl pillow requests dropbox ipywidgets reportlab gspread google-auth google-auth-oauthlib

import io
import time
import threading
from datetime import datetime
from typing import Optional

import pandas as pd
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage

import ipywidgets as widgets
from IPython.display import display, clear_output

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.pdfbase.pdfmetrics import stringWidth

from google.colab import userdata
from google.colab import auth as colab_auth
import google.auth
import gspread

import dropbox

# ----------------- CONFIG EDITABLE -----------------
GSHEET_URL   = "https://docs.google.com/spreadsheets/d/1ZwDMtvPPMYiEzbtMxOmtsC7XWb2c01lLqGxqmhxM2P4/edit"
SHEET_NAME   = "CATALOGO"

DROPBOX_ROOT = "/DISEÑO GRAFICO/GAME CENTER/SKUs/SKUs"
OUT_XLSX     = "/content/CatalogoPEND.xlsx"
OUT_PDF      = "/content/CatalogoPEND.pdf"

THUMB_W, THUMB_H = 180, 180
EXPECTED_DROPBOX_EMAIL = "gamecentersac@gmail.com"
DEBUG = False
# Nombres de las columnas en tu hoja (no importa mayúsculas/minúsculas).
COL_SKU   = "SKU GC"
COL_NAME  = "Name"
COL_PRICE = "Costo USD"
COL_EXTRA = "Marca"   # opcional: categoría o marca. Déjalo "" si no existe.
# ---------------------------------------------------

# --------- Dropbox creds (Colab Secrets) ----------
APP_KEY = userdata.get("DROPBOX_APP_KEY")
APP_SECRET = userdata.get("DROPBOX_APP_SECRET")
REFRESH_TOKEN = userdata.get("DROPBOX_REFRESH_TOKEN")
if not (APP_KEY and APP_SECRET and REFRESH_TOKEN):
    raise RuntimeError("❌ Faltan Colab Secrets: DROPBOX_APP_KEY, DROPBOX_APP_SECRET, DROPBOX_REFRESH_TOKEN")

dbx = dropbox.Dropbox(
    oauth2_refresh_token=REFRESH_TOKEN,
    app_key=APP_KEY,
    app_secret=APP_SECRET,
)

def verify_dropbox_account():
    acct = dbx.users_get_current_account()
    email = (acct.email or "").lower().strip()
    if DEBUG:
        print(f"🔐 Conectado a Dropbox como: {acct.name.display_name} | {acct.email}")
    if EXPECTED_DROPBOX_EMAIL and email != EXPECTED_DROPBOX_EMAIL.lower().strip():
        raise RuntimeError(
            f"❌ Token/Dropbox incorrecto.\nConectado como: {acct.email}\nDebe ser: {EXPECTED_DROPBOX_EMAIL}"
        )
    try:
        dbx.files_list_folder(DROPBOX_ROOT)
    except dropbox.exceptions.ApiError:
        raise RuntimeError(
            f"❌ No encuentro la carpeta de Dropbox: '{DROPBOX_ROOT}'.\n"
            "Edita DROPBOX_ROOT en CONFIG EDITABLE con la ruta real, empezando con '/' "
            "y sin incluir el nombre de tu cuenta. Ejemplo: /Catalogo/SKUs\n"
            "Revisa mayúsculas, tildes y espacios, y que el token tenga permiso sobre esa carpeta."
        )

# --------- Google Sheets PRIVADO (auth Colab) ----------
GS_CLIENT = {"gc": None}

def ensure_google_auth():
    """Abre el flujo de autorización de Google (1 vez por sesión)."""
    if GS_CLIENT["gc"] is not None:
        return GS_CLIENT["gc"]
    colab_auth.authenticate_user()
    creds, _ = google.auth.default(scopes=[
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ])
    GS_CLIENT["gc"] = gspread.authorize(creds)
    return GS_CLIENT["gc"]

def read_private_gsheet_as_df(gsheet_url: str, sheet_name: str) -> pd.DataFrame:
    gc = ensure_google_auth()
    sh = gc.open_by_url(gsheet_url)
    try:
        ws = sh.worksheet(sheet_name)
    except Exception:
        available = [w.title for w in sh.worksheets()]
        raise ValueError(f"No existe la hoja '{sheet_name}'. Hojas disponibles: {available}")
    values = ws.get_all_values()
    if not values or len(values) < 2:
        return pd.DataFrame()
    headers = [str(h).strip() for h in values[0]]
    df = pd.DataFrame(values[1:], columns=headers)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --------- Dropbox: buscar imagen ----------
def find_image(dbx, root: str, sku: str) -> Optional[bytes]:
    folder = f"{root.rstrip('/')}/{sku}"
    exts = (".jpg", ".jpeg", ".png", ".webp")

    # 1) SKU.ext
    for ext in exts:
        try:
            _, r = dbx.files_download(f"{folder}/{sku}{ext}")
            return r.content
        except Exception:
            pass

    # 2) fallback: la imagen más grande de la carpeta
    try:
        res = dbx.files_list_folder(folder)
    except Exception:
        return None
    imgs = [e for e in res.entries
            if isinstance(e, dropbox.files.FileMetadata) and e.name.lower().endswith(exts)]
    if not imgs:
        return None
    best = max(imgs, key=lambda f: f.size or 0)
    try:
        _, r = dbx.files_download(best.path_lower)
        return r.content
    except Exception:
        return None

def make_thumb_png(data: bytes, w: int, h: int) -> Optional[bytes]:
    try:
        im = PILImage.open(io.BytesIO(data))
        if im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGB")
        im.thumbnail((w, h))
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None

# --------- Utilidades de datos ----------
def parse_price(raw) -> Optional[float]:
    """Acepta '1234.50', '1,234.50', '1.234,50', 'S/ 99,90', etc."""
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

def fmt_price(item: dict, currency: str) -> str:
    if item["price"] is not None:
        txt = f"{item['price']:,.2f}"
        return f"{currency} {txt}".strip()
    return item["price_raw"] or "-"

def text_color_on(hex_color: str):
    """Blanco o negro según el brillo del color de fondo."""
    c = colors.HexColor(hex_color)
    lum = 0.299 * c.red + 0.587 * c.green + 0.114 * c.blue
    return colors.black if lum > 0.62 else colors.white

# --------- PDF ----------
INK = colors.HexColor("#1B2530")
MUTED = colors.HexColor("#66727D")
GRID = colors.HexColor("#C9D0D6")
TINT = colors.HexColor("#F2F5F7")

def wrap_lines(text, font, size, max_w, max_lines=None):
    """Corta el texto por ancho real (no por cantidad de caracteres)."""
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

def draw_lines(c, lines, x, y, lh, align="left"):
    for j, ln in enumerate(lines):
        yy = y - j * lh
        if align == "center":
            c.drawCentredString(x, yy, ln)
        elif align == "right":
            c.drawRightString(x, yy, ln)
        else:
            c.drawString(x, yy, ln)

def draw_thumb(c, png, x, y, w, h):
    img = ImageReader(io.BytesIO(png))
    c.drawImage(img, x, y, w, h, preserveAspectRatio=True, anchor="c", mask="auto")

def build_pdf(items, out_pdf, o):
    c = canvas.Canvas(out_pdf, pagesize=A4)
    W, H = A4
    M, BOTTOM = 36, 44
    usable = W - 2 * M
    accent = colors.HexColor(o["accent"])
    on_accent = text_color_on(o["accent"])
    show_extra = o["show_extra"] and o["has_extra"]
    date_txt = datetime.now().strftime("%d/%m/%Y")

    def page_header(first, page_no):
        """Dibuja encabezado y pie. Devuelve la coordenada Y donde empieza el contenido."""
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

    # ---------- Plantilla TABLA ----------
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

            # fila de encabezado
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
                if it["thumb_png"]:
                    try:
                        draw_thumb(c, it["thumb_png"], bx, by, 88, 88)
                    except Exception:
                        it["thumb_png"] = None
                if not it["thumb_png"]:
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
                if show_extra and it["extra"]:
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

    # ---------- Plantilla CUADRÍCULA ----------
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

                if it["thumb_png"]:
                    try:
                        draw_thumb(c, it["thumb_png"], x + 4, yt - ih + 4, cw - 8, ih - 8)
                    except Exception:
                        it["thumb_png"] = None
                if not it["thumb_png"]:
                    c.setFillColor(MUTED)
                    c.setFont("Helvetica-Oblique", 8.5)
                    c.drawCentredString(x + cw / 2, yt - ih / 2 - 3, "Sin imagen")

                with_extra = show_extra and it["extra"] and cols < 4
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

# --------- UI ----------
uploader = widgets.FileUpload(accept=".xlsx", multiple=False, description="📄 Elegir Excel SKUs")

w_template = widgets.ToggleButtons(options=["Tabla", "Cuadrícula"], value="Tabla", description="Plantilla:")
w_cols = widgets.Dropdown(options=[2, 3, 4], value=3, description="Por fila:")
w_title = widgets.Text(value="Catálogo de productos", description="Título:", layout=widgets.Layout(width="420px"))
w_subtitle = widgets.Text(value="", description="Subtítulo:", layout=widgets.Layout(width="420px"))
w_currency = widgets.Text(value="S/", description="Moneda:", layout=widgets.Layout(width="200px"))
w_show_price = widgets.Checkbox(value=True, description="Mostrar precio", indent=False)
w_show_extra = widgets.Checkbox(value=True, description="Mostrar categoría o marca", indent=False)
w_accent = widgets.ColorPicker(value="#0B6E8A", description="Color:", concise=False)

btn_generate = widgets.Button(description="🚀 Generar catálogo", button_style="success")
progress = widgets.IntProgress(value=0, min=0, max=100, description="0%", bar_style="", orientation="horizontal")
step_label = widgets.HTML("<i>1) Sube tu Excel de SKUs. 2) Elige el diseño. 3) Genera.</i>")
timer_label = widgets.HTML("⏱️ <b>Tiempo:</b> 00:00")
status = widgets.Output()
download_box = widgets.Output()

RUN_LOCK = {"running": False}

def fmt_elapsed(sec: float) -> str:
    sec = int(sec)
    m, s = divmod(sec, 60)
    h, m = divmod(m, 60)
    return f"{h:02d}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"

def start_stopwatch():
    """Cronómetro en vivo: arranca al pulsar 'Generar catálogo'."""
    stop = threading.Event()
    t0 = time.perf_counter()

    def tick():
        while not stop.wait(1):
            timer_label.value = f"⏱️ <b>Tiempo:</b> {fmt_elapsed(time.perf_counter() - t0)}"

    th = threading.Thread(target=tick, daemon=True)
    th.start()
    return stop, t0

def on_template_change(change):
    # "Productos por fila" solo aplica a la cuadrícula
    w_cols.layout.display = "" if w_template.value == "Cuadrícula" else "none"

w_template.observe(on_template_change, names="value")
w_cols.layout.display = "none"

def set_progress(pct: int):
    pct = max(0, min(100, int(pct)))
    progress.value = pct
    progress.description = f"{pct}%"

def get_uploaded_path() -> str:
    if not uploader.value:
        raise ValueError("Primero selecciona el archivo Excel con los SKUs.")
    item = next(iter(uploader.value.values()))
    path = f"/content/{item['metadata']['name']}"
    with open(path, "wb") as f:
        f.write(item["content"])
    return path

def read_options(has_extra: bool) -> dict:
    return {
        "template": w_template.value,
        "grid_cols": int(w_cols.value),
        "title": w_title.value.strip(),
        "subtitle": w_subtitle.value.strip(),
        "currency": w_currency.value.strip(),
        "show_price": w_price_value(),
        "show_extra": w_show_extra.value,
        "has_extra": has_extra,
        "accent": w_accent.value,
    }

def w_price_value() -> bool:
    return bool(w_show_price.value)

# --------- Excel ----------
def build_xlsx(items, out_xlsx, o):
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
            ws.cell(ri, next_col, it["extra"]).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            next_col += 1
        if show_price:
            if it["price"] is not None:
                cp = ws.cell(ri, next_col, it["price"])
                cp.number_format = price_fmt
            else:
                cp = ws.cell(ri, next_col, it["price_raw"] or None)
            cp.alignment = Alignment(horizontal="center", vertical="center")

        if it["thumb_png"]:
            xl = XLImage(io.BytesIO(it["thumb_png"]))
            s = min(150 / xl.width, 150 / xl.height)
            xl.width, xl.height = int(xl.width * s), int(xl.height * s)
            ws.add_image(xl, f"B{ri}")
            ws.row_dimensions[ri].height = xl.height * 0.75 + 10
        else:
            ws.cell(ri, 2, "(sin imagen)").alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[ri].height = 22

        for ci in range(1, len(headers) + 1):
            ws.cell(ri, ci).border = border

    ws.freeze_panes = "A2"
    wb.save(out_xlsx)

# --------- Proceso principal ----------
def on_generate(_):
    if RUN_LOCK["running"]:
        return
    RUN_LOCK["running"] = True
    btn_generate.disabled = True
    timer_label.value = "⏱️ <b>Tiempo:</b> 00:00"
    stop_timer, t_start = start_stopwatch()
    t_img = 0.0

    with status:
        clear_output()
    with download_box:
        clear_output()

    try:
        with status:
            set_progress(0)
            progress.bar_style = ""
            step_label.value = "<i>Iniciando...</i>"

            verify_dropbox_account()

            # 1) SKUs pedidos
            skus_xlsx = get_uploaded_path()
            print("✅ Archivo recibido:", skus_xlsx)
            df_skus = pd.read_excel(skus_xlsx, dtype=str)
            df_skus.columns = [str(c).strip() for c in df_skus.columns]
            sku_col = "SKU" if "SKU" in df_skus.columns else df_skus.columns[0]
            skus = [str(s).strip() for s in df_skus[sku_col].dropna().unique()]
            if not skus:
                raise ValueError("El Excel de SKUs está vacío o no tiene valores.")
            print("📦 SKUs cargados:", len(skus))

            # 2) Catálogo en Google Sheet
            step_label.value = "<b>Autorizando Google & leyendo catálogo privado...</b>"
            df_cat = read_private_gsheet_as_df(GSHEET_URL, SHEET_NAME)
            cols = {c.lower(): c for c in df_cat.columns}

            for n in (COL_SKU, COL_NAME, COL_PRICE):
                if n.lower() not in cols:
                    raise ValueError(f"Falta columna '{n}' en la hoja. Columnas: {list(df_cat.columns)}")

            has_extra = bool(COL_EXTRA) and COL_EXTRA.lower() in cols
            if w_show_extra.value and not has_extra:
                print(f"ℹ️ No se encontró la columna '{COL_EXTRA}'. Se omite categoría/marca.")

            use = [cols[COL_SKU.lower()], cols[COL_NAME.lower()], cols[COL_PRICE.lower()]]
            names = ["SKU", "Descripcion", "Precio"]
            if has_extra:
                use.append(cols[COL_EXTRA.lower()])
                names.append("Extra")
            df_cat = df_cat[use].copy()
            df_cat.columns = names
            if not has_extra:
                df_cat["Extra"] = ""
            df_cat["SKU"] = df_cat["SKU"].astype(str).str.strip()
            df_cat = df_cat.drop_duplicates(subset="SKU")

            order = {s: i for i, s in enumerate(skus)}
            df_cat = df_cat[df_cat["SKU"].isin(skus)].copy()
            df_cat["__o"] = df_cat["SKU"].map(order)
            df_cat = df_cat.sort_values("__o").drop(columns="__o")

            total = len(df_cat)
            print("🧾 Filas encontradas en catálogo:", total)
            if total == 0:
                raise ValueError("No se encontró ningún SKU del Excel dentro del Google Sheet.")
            not_found = [s for s in skus if s not in set(df_cat["SKU"])]
            if not_found:
                print(f"⚠️ SKUs que no están en el Sheet ({len(not_found)}): {', '.join(not_found[:15])}")

            # 3) Datos + imágenes
            o = read_options(has_extra)
            items, missing_img = [], []
            t_img_start = time.perf_counter()
            set_progress(5)
            for i, r in enumerate(df_cat.itertuples(index=False), start=1):
                sku = str(r.SKU).strip()
                step_label.value = f"<b>Procesando:</b> {sku} <span style='color:gray'>({i}/{total})</span>"

                price_raw = "" if pd.isna(r.Precio) else str(r.Precio).strip()
                thumb = None
                try:
                    img_data = find_image(dbx, DROPBOX_ROOT, sku)
                    if img_data:
                        thumb = make_thumb_png(img_data, THUMB_W, THUMB_H)
                except Exception as e:
                    print(f"⚠️ {sku}: error al leer la imagen ({e})")
                if not thumb:
                    missing_img.append(sku)

                items.append({
                    "sku": sku,
                    "desc": "" if pd.isna(r.Descripcion) else str(r.Descripcion).strip(),
                    "price": parse_price(price_raw),
                    "price_raw": price_raw,
                    "extra": "" if pd.isna(r.Extra) else str(r.Extra).strip(),
                    "thumb_png": thumb,
                })
                set_progress(5 + int((i / total) * 80))

            t_img = time.perf_counter() - t_img_start

            # 4) Archivos
            step_label.value = "<b>Generando Excel...</b>"
            build_xlsx(items, OUT_XLSX, o)
            set_progress(90)

            step_label.value = f"<b>Generando PDF ({o['template']})...</b>"
            build_pdf(items, OUT_PDF, o)

            set_progress(100)
            progress.bar_style = "success"
            step_label.value = "<b>✅ Terminado.</b>"

            print(f"\n✅ Excel: {OUT_XLSX}")
            print(f"✅ PDF:   {OUT_PDF}")
            print(f"🖼️  SKUs sin imagen: {len(missing_img)}" + (f" -> {', '.join(missing_img[:15])}" if missing_img else ""))
            total_s = time.perf_counter() - t_start
            print(f"⏱️ Tiempo total: {fmt_elapsed(total_s)} ({total_s:.1f} s)")
            print(f"   - Descarga de imágenes: {t_img:.1f} s ({t_img / max(total, 1):.2f} s por SKU)")
            print(f"   - Resto (lectura, Excel, PDF): {total_s - t_img:.1f} s")

        with download_box:
            clear_output()
            btn_xlsx = widgets.Button(description="⬇️ Descargar Excel", button_style="info")
            btn_pdf = widgets.Button(description="⬇️ Descargar PDF", button_style="info")

            def dl_xlsx(_):
                from google.colab import files as colab_files
                btn_xlsx.disabled = True
                btn_xlsx.description = "Descargando..."
                colab_files.download(OUT_XLSX)
                btn_xlsx.description = "✅ Excel descargado"

            def dl_pdf(_):
                from google.colab import files as colab_files
                btn_pdf.disabled = True
                btn_pdf.description = "Descargando..."
                colab_files.download(OUT_PDF)
                btn_pdf.description = "✅ PDF descargado"

            btn_xlsx.on_click(dl_xlsx)
            btn_pdf.on_click(dl_pdf)
            display(widgets.HTML("<b>Descargas:</b> (haz click 1 vez en cada botón)"))
            display(widgets.HBox([btn_xlsx, btn_pdf]))

    except Exception as e:
        with status:
            progress.bar_style = "danger"
            step_label.value = "<b>❌ Error.</b>"
            print("❌ Error:", e)
            print("\n📌 Si es error de Google Sheet privado:")
            print(" - Autoriza con una cuenta que tenga acceso al Sheet.")
            print(" - Comparte el Sheet con esa cuenta (Lector/Editor).")

    finally:
        stop_timer.set()
        timer_label.value = f"⏱️ <b>Tiempo:</b> {fmt_elapsed(time.perf_counter() - t_start)} (finalizado)"
        RUN_LOCK["running"] = False
        btn_generate.disabled = False

btn_generate.on_click(on_generate)

clear_output()
display(widgets.VBox([
    widgets.HTML("<b>1)</b> Selecciona tu Excel de SKUs (SKUs en la primera columna)"),
    uploader,
    widgets.HTML("<b>2)</b> Elige el diseño del catálogo"),
    w_template,
    w_cols,
    w_title,
    w_subtitle,
    w_currency,
    widgets.HBox([w_show_price, w_show_extra]),
    w_accent,
    widgets.HTML("<b>3)</b> Genera el catálogo (te pedirá autorizar Google la primera vez)"),
    btn_generate,
    widgets.HBox([progress, step_label]),
    timer_label,
    status,
    download_box,
]))
