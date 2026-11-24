"""Cliente y utilidades para interactuar con la API de Dropbox."""
import io
from typing import Optional
from PIL import Image as PILImage
import dropbox

def get_dropbox_client(app_key: str, app_secret: str, refresh_token: str) -> dropbox.Dropbox:
    """Instancia el cliente de Dropbox utilizando OAuth2 con refresh token."""
    if not (app_key and app_secret and refresh_token):
        raise ValueError("Faltan credenciales de Dropbox (DROPBOX_APP_KEY, DROPBOX_APP_SECRET, DROPBOX_REFRESH_TOKEN).")
    return dropbox.Dropbox(
        oauth2_refresh_token=refresh_token,
        app_key=app_key,
        app_secret=app_secret,
    )

def verify_dropbox_account(dbx: dropbox.Dropbox, expected_email: str = "", dropbox_root: str = "/"):
    """Verifica conexión con la cuenta y existencia de la carpeta raíz."""
    acct = dbx.users_get_current_account()
    email = (acct.email or "").lower().strip()
    if expected_email and email != expected_email.lower().strip():
        raise RuntimeError(
            f"Token de Dropbox incorrecto. Conectado como: {acct.email}. Se esperaba: {expected_email}"
        )
    if dropbox_root and dropbox_root != "/":
        try:
            dbx.files_list_folder(dropbox_root)
        except dropbox.exceptions.ApiError:
            raise RuntimeError(
                f"No se encuentra la carpeta de Dropbox: '{dropbox_root}'. "
                "Verifica la ruta y permisos en tu cuenta de Dropbox."
            )
    return acct

def find_image(dbx: dropbox.Dropbox, root: str, sku: str) -> Optional[bytes]:
    """
    Busca la imagen asociada al SKU en la carpeta de Dropbox.
    1. Intenta SKU.ext (.jpg, .jpeg, .png, .webp)
    2. Fallback: la imagen más pesada dentro de la carpeta del SKU.
    """
    folder = f"{root.rstrip('/')}/{sku}"
    exts = (".jpg", ".jpeg", ".png", ".webp")

    # 1) Intento directo SKU.ext
    for ext in exts:
        try:
            _, r = dbx.files_download(f"{folder}/{sku}{ext}")
            return r.content
        except Exception:
            pass

    # 2) Fallback: listar carpeta y tomar la imagen más grande
    try:
        res = dbx.files_list_folder(folder)
    except Exception:
        return None

    imgs = [
        e for e in res.entries
        if isinstance(e, dropbox.files.FileMetadata) and e.name.lower().endswith(exts)
    ]
    if not imgs:
        return None

    best = max(imgs, key=lambda f: f.size or 0)
    try:
        _, r = dbx.files_download(best.path_lower)
        return r.content
    except Exception:
        return None

def make_thumb_png(data: bytes, w: int, h: int) -> Optional[bytes]:
    """Convierte y redimensiona bytes de imagen a PNG manteniendo proporción."""
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
