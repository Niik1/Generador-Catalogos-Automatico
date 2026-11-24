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

# Configuración por defecto (valores genéricos para el repositorio público)
DEFAULT_GSHEET_URL = ""
DEFAULT_SHEET_NAME = "CATALOGO"
DEFAULT_DROPBOX_ROOT = "/SKUs"
DEFAULT_EXPECTED_EMAIL = ""

DEFAULT_COL_SKU = "SKU"
DEFAULT_COL_NAME = "Name"
DEFAULT_COL_PRICE = "Price"
DEFAULT_COL_EXTRA = "Marca"
