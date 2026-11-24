"""Cliente y utilidades para obtener el catálogo desde Google Sheets o fuentes de datos."""
import re
from typing import Optional
import pandas as pd
import gspread

def extract_sheet_id(url: str) -> Optional[str]:
    """Extrae el ID de una URL de Google Sheets."""
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else None

def read_gsheet_public_or_export(gsheet_url: str, sheet_name: str = "") -> Optional[pd.DataFrame]:
    """
    Intenta leer Google Sheet usando el endpoint público de exportación CSV o gviz.
    Funciona si el Sheet está compartido como 'Cualquier persona con el enlace'.
    """
    sheet_id = extract_sheet_id(gsheet_url)
    if not sheet_id:
        return None

    urls_to_try = []
    if sheet_name:
        urls_to_try.append(f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}")
        urls_to_try.append(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={sheet_name}")
    urls_to_try.append(f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv")

    for u in urls_to_try:
        try:
            df = pd.read_csv(u, dtype=str)
            if not df.empty and len(df.columns) > 1:
                df.columns = [str(c).strip() for c in df.columns]
                return df
        except Exception:
            continue
    return None

def read_gsheet_with_service_account(service_account_info: dict, gsheet_url: str, sheet_name: str) -> pd.DataFrame:
    """Lee una hoja de cálculo privada utilizando credenciales de Google Service Account."""
    gc = gspread.service_account_from_dict(service_account_info)
    sh = gc.open_by_url(gsheet_url)
    ws = sh.worksheet(sheet_name)
    values = ws.get_all_values()
    if not values or len(values) < 2:
        return pd.DataFrame()
    headers = [str(h).strip() for h in values[0]]
    df = pd.DataFrame(values[1:], columns=headers)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def get_demo_catalog() -> pd.DataFrame:
    """Devuelve un catálogo de prueba para demostraciones interactivas."""
    return pd.DataFrame([
        {"SKU GC": "GC-001", "Name": "Auriculares Gamer HyperX Cloud Alpha Pro", "Costo USD": "79.99", "Marca": "HyperX"},
        {"SKU GC": "GC-002", "Name": "Teclado Mecánico Razer Huntsman Mini 60% RGB", "Costo USD": "119.50", "Marca": "Razer"},
        {"SKU GC": "GC-003", "Name": "Mouse Inalámbrico Logitech G Pro X Superlight", "Costo USD": "139.00", "Marca": "Logitech"},
        {"SKU GC": "GC-004", "Name": "Monitor Gamer Asus ROG Swift 27\" 165Hz IPS", "Costo USD": "289.90", "Marca": "Asus ROG"},
        {"SKU GC": "GC-005", "Name": "Silla Gamer Corsair T3 Rush Charcoal", "Costo USD": "249.00", "Marca": "Corsair"},
    ])
