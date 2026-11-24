"""Módulo de configuración y gestión de secretos/variables de entorno."""
import os
from typing import Any

# Intentamos cargar dotenv si está disponible
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

def get_secret(key: str, default: Any = None) -> Any:
    """
    Obtiene una clave de configuración buscando en:
    1. streamlit secrets (si streamlit está corriendo)
    2. variables de entorno (os.environ o .env)
    3. valor por defecto proporcionado
    """
    try:
        import streamlit as st
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass

    return os.environ.get(key, default)

# Configuración por defecto de Google Sheets
DEFAULT_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1ZwDMtvPPMYiEzbtMxOmtsC7XWb2c01lLqGxqmhxM2P4/edit"
DEFAULT_SHEET_NAME = "CATALOGO"
DEFAULT_DROPBOX_ROOT = "/DISEÑO GRAFICO/GAME CENTER/SKUs/SKUs"
DEFAULT_EXPECTED_EMAIL = "gamecentersac@gmail.com"

DEFAULT_COL_SKU = "SKU GC"
DEFAULT_COL_NAME = "Name"
DEFAULT_COL_PRICE = "Costo USD"
DEFAULT_COL_EXTRA = "Marca"
