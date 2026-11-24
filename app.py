"""Aplicación Web Streamlit - Generador Automático de Catálogos (PDF + Excel)."""
import io
import time
import pandas as pd
import streamlit as st

from src.config import (
    get_secret,
    DEFAULT_GSHEET_URL,
    DEFAULT_SHEET_NAME,
    DEFAULT_DROPBOX_ROOT,
    DEFAULT_EXPECTED_EMAIL,
    DEFAULT_COL_SKU,
    DEFAULT_COL_NAME,
    DEFAULT_COL_PRICE,
    DEFAULT_COL_EXTRA,
)
from src.dropbox_client import get_dropbox_client, verify_dropbox_account, find_image, make_thumb_png
from src.sheets_client import (
    read_gsheet_public_or_export,
    read_gsheet_with_service_account,
    get_demo_catalog,
)
from src.catalog_generator import build_pdf, build_xlsx, parse_price

st.set_page_config(
    page_title="Generador Automático de Catálogos",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilos CSS personalizados para apariencia moderna
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1B2530;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #556270;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border-radius: 8px;
        padding: 12px;
        border: 1px solid #E2E8F0;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# Encabezado
st.markdown('<div class="main-header">📦 Generador Automático de Catálogos</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Automatización end-to-end con integración de Dropbox API, Google Sheets y generación vectorial de PDF y Excel.</div>', unsafe_allow_html=True)

# ----------------- SIDEBAR: CONFIGURACIÓN DE DISEÑO -----------------
with st.sidebar:
    st.header("🎨 Diseño del Catálogo")
    
    template = st.radio("Plantilla:", ["Tabla", "Cuadrícula"], index=0, help="Tabla es ideal para listas densas con especificaciones; Cuadrícula para catálogos comerciales tipo e-commerce.")
    
    grid_cols = 3
    if template == "Cuadrícula":
        grid_cols = st.selectbox("Productos por fila:", [2, 3, 4], index=1)
        
    title = st.text_input("Título del catálogo:", value="Catálogo de Productos")
    subtitle = st.text_input("Subtítulo:", value="Catálogo Oficial")
    currency = st.text_input("Moneda:", value="USD", help="Símbolo o código: USD, S/, $, €")
    
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1:
        show_price = st.checkbox("Mostrar precio", value=True)
    with col_opt2:
        show_extra = st.checkbox("Mostrar marca", value=True)
        
    accent_color = st.color_picker("Color de acento:", value="#0B6E8A")

    st.markdown("---")
    st.header("⚙️ Fuente de Datos")
    
    source_mode = st.radio(
        "Modo de catálogo:",
        ["Google Sheets (Conexión Real)", "Modo Demo (5 Productos de Prueba)"],
        index=0,
        help="El modo demo permite probar la generación sin credenciales activas de Google Sheets."
    )
    
    with st.expander("Parámetros avanzados de conexión"):
        gsheet_url = st.text_input("Google Sheet URL:", value=get_secret("GSHEET_URL", DEFAULT_GSHEET_URL))
        sheet_name = st.text_input("Nombre de la Hoja:", value=get_secret("SHEET_NAME", DEFAULT_SHEET_NAME))
        dropbox_root = st.text_input("Ruta raíz Dropbox:", value=get_secret("DROPBOX_ROOT", DEFAULT_DROPBOX_ROOT))
        expected_email = st.text_input("Email esperado Dropbox:", value=get_secret("EXPECTED_DROPBOX_EMAIL", DEFAULT_EXPECTED_EMAIL))

# ----------------- ÁREA PRINCIPAL: CARGA DE ARCHIVO -----------------
tab_main, tab_about = st.tabs(["🚀 Generador", "ℹ️ Arquitectura & Documentación"])

with tab_main:
    col_upload, col_preview = st.columns([1, 1], gap="medium")
    
    with col_upload:
        st.subheader("1. Carga de SKUs")
        uploaded_file = st.file_uploader(
            "Selecciona un archivo Excel con la columna de SKUs:",
            type=["xlsx", "xls"],
            help="Debe contener los códigos de producto en la primera columna o con encabezado 'SKU'."
        )
        
        # Botón para descargar plantilla de ejemplo
        sample_df = pd.DataFrame({"SKU": ["GC-001", "GC-002", "GC-003", "GC-004"]})
        sample_buffer = io.BytesIO()
        sample_df.to_excel(sample_buffer, index=False)
        st.download_button(
            label="📄 Descargar Excel de ejemplo",
            data=sample_buffer.getvalue(),
            file_name="skus_ejemplo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    df_skus = None
    skus_list = []
    
    with col_preview:
        st.subheader("2. Vista previa")
        if uploaded_file is not None:
            try:
                df_skus = pd.read_excel(uploaded_file, dtype=str)
                df_skus.columns = [str(c).strip() for c in df_skus.columns]
                sku_col = "SKU" if "SKU" in df_skus.columns else df_skus.columns[0]
                skus_list = [str(s).strip() for s in df_skus[sku_col].dropna().unique()]
                
                st.success(f"✅ **{len(skus_list)} SKUs únicos detectados**")
                st.dataframe(df_skus[[sku_col]].head(10), use_container_width=True, height=180)
            except Exception as e:
                st.error(f"Error al leer el archivo Excel: {e}")
        else:
            st.info("Sube un archivo Excel para habilitar la generación.")

    st.markdown("---")
    
    # Botón de ejecución
    btn_disabled = (uploaded_file is None and source_mode != "Modo Demo (5 Productos de Prueba)")
    
    if st.button("🚀 Iniciar Generación de Catálogo", type="primary", disabled=btn_disabled):
        t_start = time.perf_counter()
        progress_bar = st.progress(5, text="Iniciando conexión con servicios...")
        status_box = st.status("Procesando catálogo...", expanded=True)
        
        try:
            # 1. Obtener Dropbox Client
            status_box.write("🔐 Verificando credenciales de Dropbox...")
            app_key = get_secret("DROPBOX_APP_KEY")
            app_secret = get_secret("DROPBOX_APP_SECRET")
            refresh_token = get_secret("DROPBOX_REFRESH_TOKEN")
            
            dbx = None
            if app_key and app_secret and refresh_token:
                try:
                    dbx = get_dropbox_client(app_key, app_secret, refresh_token)
                    verify_dropbox_account(dbx, expected_email, dropbox_root)
                    status_box.write("✅ Conectado a Dropbox correctamente.")
                except Exception as ex:
                    status_box.warning(f"⚠️ Aviso Dropbox: {ex}. Las imágenes se procesarán como placeholder si no hay acceso.")
            else:
                status_box.info("ℹ️ Credenciales de Dropbox no configuradas en este entorno. Se usarán placeholders para imágenes.")

            # 2. Cargar datos del catálogo
            progress_bar.progress(25, text="Obteniendo catálogo de productos...")
            status_box.write("📊 Obteniendo datos del catálogo...")
            
            if source_mode == "Modo Demo (5 Productos de Prueba)":
                df_cat = get_demo_catalog()
                if not skus_list:
                    skus_list = list(df_cat["SKU GC"])
            else:
                # Intentar lectura pública/export o service account
                df_cat = read_gsheet_public_or_export(gsheet_url, sheet_name)
                
                # Si falló, intentar con service account si existe en secrets
                if df_cat is None and "gcp_service_account" in st.secrets:
                    df_cat = read_gsheet_with_service_account(dict(st.secrets["gcp_service_account"]), gsheet_url, sheet_name)
                
                if df_cat is None:
                    raise RuntimeError(
                        "No se pudo acceder al Google Sheet. Asegúrate de que el documento esté compartido "
                        "como 'Cualquier persona con el enlace' (Lector) o añade tus credenciales de Service Account."
                    )

            # Normalizar columnas
            cols = {c.lower(): c for c in df_cat.columns}
            col_sku_real = cols.get(DEFAULT_COL_SKU.lower(), df_cat.columns[0])
            col_name_real = cols.get(DEFAULT_COL_NAME.lower(), df_cat.columns[1] if len(df_cat.columns) > 1 else df_cat.columns[0])
            col_price_real = cols.get(DEFAULT_COL_PRICE.lower(), None)
            col_extra_real = cols.get(DEFAULT_COL_EXTRA.lower(), None)

            df_cat = df_cat.rename(columns={
                col_sku_real: "SKU",
                col_name_real: "Descripcion",
            })
            df_cat["Precio"] = df_cat[col_price_real] if col_price_real else ""
            df_cat["Extra"] = df_cat[col_extra_real] if col_extra_real else ""
            
            df_cat["SKU"] = df_cat["SKU"].astype(str).str.strip()
            df_cat = df_cat.drop_duplicates(subset="SKU")
            
            # Filtrar por los SKUs solicitados manteniendo el orden
            order = {s: i for i, s in enumerate(skus_list)}
            df_cat = df_cat[df_cat["SKU"].isin(skus_list)].copy()
            df_cat["__order"] = df_cat["SKU"].map(order)
            df_cat = df_cat.sort_values("__order").drop(columns="__order")

            total_items = len(df_cat)
            if total_items == 0:
                raise ValueError("Ninguno de los SKUs cargados coincide con los productos del catálogo.")

            status_box.write(f"🧾 {total_items} productos encontrados para procesar.")

            # 3. Procesar imágenes y armar items
            progress_bar.progress(45, text="Descargando imágenes y optimizando miniaturas...")
            items = []
            missing_img = []
            
            for i, r in enumerate(df_cat.itertuples(index=False), start=1):
                sku = str(r.SKU).strip()
                price_raw = "" if pd.isna(r.Precio) else str(r.Precio).strip()
                thumb = None
                
                if dbx is not None:
                    try:
                        img_data = find_image(dbx, dropbox_root, sku)
                        if img_data:
                            thumb = make_thumb_png(img_data, 180, 180)
                    except Exception:
                        pass
                
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
                
                pct = 45 + int((i / total_items) * 35)
                progress_bar.progress(pct, text=f"Procesando imágenes: {i}/{total_items} SKUs...")

            # Opciones de diseño
            options = {
                "template": template,
                "grid_cols": grid_cols,
                "title": title,
                "subtitle": subtitle,
                "currency": currency,
                "show_price": show_price,
                "show_extra": show_extra,
                "has_extra": bool(col_extra_real),
                "accent": accent_color,
            }

            # 4. Generar PDF y Excel en memoria
            progress_bar.progress(85, text="Generando PDF y Excel vectoriales...")
            status_box.write("📑 Compilando PDF vectorial con ReportLab...")
            pdf_buffer = io.BytesIO()
            build_pdf(items, pdf_buffer, options)
            pdf_bytes = pdf_buffer.getvalue()

            status_box.write("📊 Generando libro Excel con openpyxl...")
            xlsx_buffer = io.BytesIO()
            build_xlsx(items, xlsx_buffer, options)
            xlsx_bytes = xlsx_buffer.getvalue()

            progress_bar.progress(100, text="¡Completado con éxito!")
            status_box.update(label="✅ Catálogo generado exitosamente", state="complete", expanded=False)
            
            total_elapsed = time.perf_counter() - t_start
            
            # Métricas
            st.markdown("### 📊 Resumen de Generación")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Productos Procesados", f"{total_items}")
            m2.metric("Imágenes Encontradas", f"{total_items - len(missing_img)}")
            m3.metric("Sin Imagen", f"{len(missing_img)}")
            m4.metric("Tiempo Total", f"{total_elapsed:.1f} s")

            # Botones de descarga
            st.markdown("### ⬇️ Descargar Catálogos")
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                st.download_button(
                    label="📄 Descargar Catálogo en PDF",
                    data=pdf_bytes,
                    file_name=f"Catalogo_{time.strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            with btn_col2:
                st.download_button(
                    label="📊 Descargar Catálogo en Excel",
                    data=xlsx_bytes,
                    file_name=f"Catalogo_{time.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                
        except Exception as e:
            progress_bar.progress(0, text="Error en el procesamiento")
            status_box.update(label="❌ Ocurrió un error", state="error", expanded=True)
            st.error(f"Error: {e}")

# ----------------- PESTAÑA: ARQUITECTURA Y DOCUMENTACIÓN -----------------
with tab_about:
    st.markdown("""
    ### 🏗️ Arquitectura de la Solución
    Esta herramienta automatiza un flujo de trabajo que tradicionalmente tomaba horas de trabajo manual de diseño gráfico y hojas de cálculo:
    
    1. **Entrada de Datos:** El usuario carga un archivo Excel con la lista de SKUs que desea cotizar o presentar.
    2. **Integración con Google Sheets:** Consume la base de datos de productos actualizada en tiempo real mediante `gspread` / Google Sheets API.
    3. **Integración con Dropbox API:** Localiza y descarga dinámicamente las fotografías oficiales de cada producto por SKU usando OAuth2 con Refresh Tokens.
    4. **Motor Vectorial PDF (ReportLab):** Genera documentos de alta fidelidad para impresión y distribución digital, con control tipográfico estricto y layouts dinámicos (Tabla o Cuadrícula).
    5. **Motor de Hojas de Cálculo (openpyxl):** Inserta miniaturas escaladas y formateadas directamente dentro de las celdas de Excel.
    
    ---
    ### 💻 Stack Tecnológico
    - **Frontend / Interfaz:** Streamlit
    - **Backend & Procesamiento:** Python 3.10+
    - **Generación de Documentos:** ReportLab (PDF), openpyxl (XLSX), Pillow (Procesamiento de imágenes)
    - **APIs & Conectores:** Dropbox API v2, Google Sheets API / gspread
    """)
