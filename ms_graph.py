"""
ms_graph.py
Acceso a correos de Hotmail/Outlook.com via Microsoft Graph REST API.

Mucho más simple que IMAP:
  - OAuth2 Device Code Flow (autorizas 1 vez en el navegador, token se cachea)
  - HTTP GET para leer correos (JSON, no protocolo binario IMAP)
  - Solo necesita el permiso Mail.Read (fácil de configurar en Azure)

Seguridad:
  - Token cache cifrado con DPAPI de Windows (o permisos 0600 en Linux/Mac)
  - Tokens encapsulados para no exponerse en tracebacks
  - Sanitización de datos para evitar CSV injection
  - Timeouts en todas las peticiones HTTP
  - Validación de parámetros de entrada
"""

import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Generator, NamedTuple

try:
    import msal
except ImportError:
    print("\n❌  Falta 'msal'. Instálalo con:\n    pip install msal\n")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("\n❌  Falta 'requests'. Instálalo con:\n    pip install requests\n")
    sys.exit(1)

# Cifrado de cache: intentar DPAPI en Windows, fallback a permisos restrictivos
_USE_DPAPI = False
try:
    from msal_extensions import (
        FilePersistenceWithDataProtection,
        PersistedTokenCache,
    )
    _USE_DPAPI = True
except ImportError:
    pass  # Fallback: cache en texto plano con permisos restrictivos


# ─── Tipos ────────────────────────────────────────────────────────────────────

class MailData(NamedTuple):
    sender_email: str
    sender_name: str
    is_unread: bool


class _SecureToken:
    """Encapsula un access token para que no se muestre en tracebacks/logs."""
    __slots__ = ("_value",)

    def __init__(self, value: str):
        self._value = value

    @property
    def value(self) -> str:
        return self._value

    def __repr__(self) -> str:
        return "Token(***)"

    def __str__(self) -> str:
        return "Token(***)"


# ─── Configuración ────────────────────────────────────────────────────────────

CONFIG_DIR   = Path.home() / ".outlook-email-analyzer"
CONFIG_FILE  = CONFIG_DIR / "config.json"
CACHE_FILE   = CONFIG_DIR / "token_cache.json"

AUTHORITY    = "https://login.microsoftonline.com/consumers"
SCOPES       = ["Mail.Read"]
GRAPH_BASE   = "https://graph.microsoft.com/v1.0"

# Timeouts HTTP: (connect_timeout, read_timeout) en segundos
HTTP_TIMEOUT = (10, 30)

# Regex para validar nombres de carpeta (previene path traversal)
_FOLDER_RE = re.compile(r"^[a-zA-Z0-9_\-]+$")


def _ensure_dir():
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    # Permisos restrictivos en Unix
    if os.name != "nt":
        os.chmod(CONFIG_DIR, 0o700)


def _restrict_file(path: Path):
    """Aplica permisos restrictivos a un archivo (Unix: 0600)."""
    if os.name != "nt" and path.exists():
        os.chmod(path, 0o600)


def _load_json(path: Path) -> dict:
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def _save_json(path: Path, data: dict):
    _ensure_dir()
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    _restrict_file(path)


# ─── Setup del Client ID (1 sola vez) ────────────────────────────────────────

_SETUP_TEXT = """
  ╔═══════════════════════════════════════════════════════════════╗
  ║   CONFIGURACIÓN INICIAL  (solo 1 vez, toma ~2 minutos)       ║
  ╚═══════════════════════════════════════════════════════════════╝

  Microsoft requiere registrar una app para leer correos.
  Es GRATIS (no necesitas suscripción de Azure).

  ── Paso 1: Registrar la app ──────────────────────────────────

  1. Abre en tu navegador:
     https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade

  2. Inicia sesión con tu cuenta Microsoft (tu Hotmail)

  3. Click  "Nuevo registro"  (o "New registration"):
     • Nombre:           EmailAnalyzer
     • Tipos de cuenta:  "Solo cuentas personales de Microsoft"
                         (Personal Microsoft accounts only)
     • URI redirección:  Plataforma → "Cliente público/nativo"
                         URI → https://login.microsoftonline.com/common/oauth2/nativeclient

  4. Click  "Registrar"

  5. Copia el  "Id. de aplicación (cliente)"
     (Application client ID — es un UUID largo)

  ── Paso 2: Agregar permiso ───────────────────────────────────

  6. En la barra izquierda click  "Permisos de API"

  7. Click  "Agregar un permiso"  →  "Microsoft Graph"
     →  "Permisos delegados"  →  busca "Mail"
     →  marca  "Mail.Read"  →  "Agregar permisos"

  ¡Listo! Pega el Client ID aquí:
"""


def get_client_id() -> str:
    """Retorna el client_id guardado o guía al usuario por el setup."""
    config = _load_json(CONFIG_FILE)
    cid = config.get("client_id", "").strip()
    if cid:
        return cid

    print(_SETUP_TEXT)
    cid = input("  Application (client) ID: ").strip()

    _UUID_RE = re.compile(
        r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$",
        re.IGNORECASE,
    )
    if not cid or not _UUID_RE.match(cid):
        raise SystemExit(
            "\n  Client ID invalido.\n"
            "    Debe ser un UUID de 36 caracteres (ej: a1b2c3d4-e5f6-7890-abcd-ef1234567890).\n"
        )

    config["client_id"] = cid
    _save_json(CONFIG_FILE, config)
    print(f"\n  ✓ Guardado en {CONFIG_FILE}")
    print("    No necesitarás repetir esto.\n")
    return cid


# ─── Token Cache (cifrado DPAPI en Windows) ─────────────────────────────────

def _build_token_cache() -> msal.SerializableTokenCache:
    """
    Construye el cache de tokens.
    Windows: cifrado con DPAPI via msal_extensions (si está disponible).
    Otros:   archivo JSON con permisos 0600.

    Si el cache está corrupto o fue escrito con un método de cifrado diferente,
    se elimina automáticamente y se crea uno nuevo (el usuario re-autentica).
    """
    if _USE_DPAPI:
        _ensure_dir()
        try:
            persistence = FilePersistenceWithDataProtection(str(CACHE_FILE))
            cache = PersistedTokenCache(persistence)
            # Forzar lectura para detectar corrupción temprano
            cache.search(cache.CredentialType.ACCOUNT, query={})
            return cache
        except Exception:
            # Cache corrupto o incompatible — eliminar y crear nuevo
            if CACHE_FILE.exists():
                CACHE_FILE.unlink()
                print("  (Cache de sesion corrupto, se regenerara automaticamente)")
            persistence = FilePersistenceWithDataProtection(str(CACHE_FILE))
            return PersistedTokenCache(persistence)

    # Fallback sin msal_extensions
    cache = msal.SerializableTokenCache()
    if CACHE_FILE.exists():
        try:
            cache.deserialize(CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            CACHE_FILE.unlink()
    return cache


def _save_token_cache(cache: msal.SerializableTokenCache):
    """Guarda el cache (solo si no usa DPAPI, que guarda solo)."""
    if _USE_DPAPI:
        return  # PersistedTokenCache guarda automáticamente
    if cache.has_state_changed:
        _ensure_dir()
        CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")
        _restrict_file(CACHE_FILE)


# ─── OAuth2 Device Code Flow ─────────────────────────────────────────────────

def authenticate(username: str) -> _SecureToken:
    """
    Autentica con Microsoft. Retorna un _SecureToken (no expone el valor en logs).
    Primera vez: muestra código → usuario autoriza en navegador.
    Siguientes: usa token cacheado (se refresca solo).
    """
    client_id = get_client_id()
    cache = _build_token_cache()

    app = msal.PublicClientApplication(
        client_id,
        authority=AUTHORITY,
        token_cache=cache,
    )

    # Intentar token silencioso (cache)
    accounts = app.get_accounts(username=username)
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_token_cache(cache)
            print("  ✓ Autenticado (sesión guardada)")
            return _SecureToken(result["access_token"])

    # Device Code Flow (interactivo)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        error = flow.get("error_description", flow.get("error", "desconocido"))
        raise SystemExit(
            f"\n❌  Error al iniciar autenticación:\n    {error}\n\n"
            f"    Si el error menciona 'client_id', verifica que sea correcto.\n"
            f"    Para reconfigurar:  python main.py --reset-auth\n"
        )

    code = flow["user_code"]
    link = flow.get("verification_uri", "https://microsoft.com/devicelogin")

    print(f"\n  ┌────────────────────────────────────────────────┐")
    print(f"  │  1. Abre:   {link:<35}│")
    print(f"  │  2. Código: {code:<35}│")
    print(f"  │  3. Inicia sesión y acepta                     │")
    print(f"  └────────────────────────────────────────────────┘")
    print(f"\n  Esperando que autorices en el navegador...")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        _save_token_cache(cache)
        print("  ✓ Autenticado\n")
        return _SecureToken(result["access_token"])

    error = result.get("error_description", result.get("error", "desconocido"))
    raise SystemExit(f"\n❌  Autenticación fallida:\n    {error}\n")


# ─── Helpers HTTP ────────────────────────────────────────────────────────────

def _extract_token(token) -> str:
    """Extrae el string del token. Solo acepta _SecureToken."""
    if isinstance(token, _SecureToken):
        return token.value
    raise TypeError(f"Se esperaba _SecureToken, se recibio {type(token).__name__}")


def _auth_headers(token) -> dict:
    return {"Authorization": f"Bearer {_extract_token(token)}"}


def _parse_api_error(resp: requests.Response) -> str:
    """Extrae solo el mensaje de error de la API, sin exponer datos sensibles."""
    try:
        body = resp.json()
        error = body.get("error", {})
        return error.get("message", f"HTTP {resp.status_code}")
    except Exception:
        return f"HTTP {resp.status_code}"


def _api_get(url: str, headers: dict, params: dict | None = None) -> requests.Response:
    """GET con timeout, retry para 429 (throttling) y 401."""
    max_retries = 3
    for attempt in range(max_retries):
        resp = requests.get(url, headers=headers, params=params, timeout=HTTP_TIMEOUT)

        if resp.status_code == 429:
            # Rate limiting — respetar Retry-After (maximo 60s para evitar bloqueo)
            try:
                wait = min(int(resp.headers.get("Retry-After", 5)), 60)
            except (ValueError, TypeError):
                wait = 5
            print(f"\n  API throttling, esperando {wait}s...")
            time.sleep(wait)
            continue

        if resp.status_code == 401:
            raise SystemExit(
                "\n❌  Token expirado o inválido.\n"
                "    Ejecuta:  python main.py --reset-auth\n"
            )

        return resp

    raise SystemExit("\n❌  API no disponible después de 3 reintentos.\n")


def _validate_folder(folder: str) -> str:
    """Valida que el nombre de carpeta sea seguro (previene path traversal)."""
    if not _FOLDER_RE.match(folder):
        raise SystemExit(
            f"\n❌  Nombre de carpeta inválido: '{folder}'\n"
            f"    Solo se permiten letras, números, guiones y guiones bajos.\n"
            f"    Usa --list-folders para ver las disponibles.\n"
        )
    return folder


# ─── Microsoft Graph API ─────────────────────────────────────────────────────

def get_folder_info(token, folder: str = "inbox") -> dict:
    """Obtiene metadata de la carpeta (nombre, total, no leídos)."""
    folder = _validate_folder(folder)
    headers = _auth_headers(token)
    resp = _api_get(f"{GRAPH_BASE}/me/mailFolders/{folder}", headers)

    if resp.status_code == 404:
        raise SystemExit(
            f"\n❌  Carpeta '{folder}' no encontrada.\n"
            f"    Usa --list-folders para ver las disponibles.\n"
        )
    if resp.status_code != 200:
        raise SystemExit(f"\n❌  Error API: {_parse_api_error(resp)}\n")

    data = resp.json()
    return {
        "name":         data.get("displayName", folder),
        "total":        data.get("totalItemCount", 0),
        "unread":       data.get("unreadItemCount", 0),
    }


def read_emails(
    token,
    folder: str = "inbox",
    unread_only: bool = False,
    progress_callback=None,
) -> Generator[MailData, None, None]:
    """
    Lee correos via Graph API. Yielda MailData.
    Muy eficiente: solo pide los campos 'from' e 'isRead' (no descarga cuerpo).
    Pagina automáticamente (999 correos por request).
    Incluye retry para 429 (throttling) y timeout en todas las peticiones.
    """
    folder = _validate_folder(folder)
    headers = _auth_headers(token)

    url = f"{GRAPH_BASE}/me/mailFolders/{folder}/messages"
    params: dict = {
        "$select": "from,isRead",
        "$top":    999,
        "$orderby": "receivedDateTime desc",
    }
    if unread_only:
        params["$filter"] = "isRead eq false"

    fetched = 0

    while url:
        resp = _api_get(url, headers, params)

        if resp.status_code != 200:
            raise SystemExit(f"\n❌  Error API: {_parse_api_error(resp)}\n")

        data = resp.json()
        messages = data.get("value", [])

        for msg in messages:
            from_obj = msg.get("from", {}).get("emailAddress", {})
            email_addr = (from_obj.get("address") or "").lower().strip()
            name       = (from_obj.get("name") or email_addr).strip()
            is_unread  = not msg.get("isRead", True)

            if not email_addr:
                fetched += 1
                continue

            fetched += 1
            if progress_callback:
                progress_callback(fetched)

            yield MailData(
                sender_email=email_addr,
                sender_name=name,
                is_unread=is_unread,
            )

        # Paginación: Graph API retorna @odata.nextLink con la URL de la siguiente página
        url = data.get("@odata.nextLink")
        params = {}  # nextLink ya incluye los parámetros

    # Asegurar callback final
    if progress_callback:
        progress_callback(fetched)


def list_folders(token) -> list[tuple[str, int, int]]:
    """Lista carpetas de correo. Retorna [(nombre, total, no_leídos), ...]."""
    headers = _auth_headers(token)
    resp = _api_get(
        f"{GRAPH_BASE}/me/mailFolders",
        headers,
        params={"$top": 100},
    )
    if resp.status_code != 200:
        return []

    result = []
    for f in resp.json().get("value", []):
        name   = f.get("displayName", "?")
        total  = f.get("totalItemCount", 0)
        unread = f.get("unreadItemCount", 0)
        result.append((name, total, unread))
    return result


# ─── Utilidades ───────────────────────────────────────────────────────────────

def reset_auth():
    """Borra configuración y tokens guardados (pide confirmación)."""
    files = [p for p in (CONFIG_FILE, CACHE_FILE) if p.exists()]
    if not files:
        print("  (No había configuración guardada)\n")
        return

    print("  Se eliminarán:")
    for p in files:
        print(f"    - {p}")
    confirm = input("\n  ¿Confirmar? (s/N): ").strip().lower()
    if confirm not in ("s", "si", "sí", "y", "yes"):
        print("  Cancelado.\n")
        return

    for path in files:
        path.unlink()
        print(f"  ✓ Eliminado: {path}")
    print("  La próxima ejecución pedirá configurar desde cero.\n")
