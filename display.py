"""
display.py
Formateo y presentacion del ranking en consola.
Soporta modo normal (ranking) y modo newsletters (con link de desuscripcion).
"""

import sys
from typing import Sequence

from analyzer import SenderStats

# ─── Anchos de columna ───────────────────────────────────────────────────────

# Modo normal
COL_NUM   = 5
COL_NAME  = 30
COL_EMAIL = 38
COL_COUNT = 10
COL_BAR   = 20

# Modo newsletters (columna link reemplaza la barra)
COL_UNSUB = 5   # "SI" / "NO"
COL_LINK  = 55  # URL truncada

_NORMAL_WIDTH = COL_NUM + COL_NAME + COL_EMAIL + COL_COUNT + COL_BAR + 10
_NEWS_WIDTH   = COL_NUM + COL_NAME + COL_EMAIL + COL_COUNT + COL_UNSUB + COL_LINK + 14

SEPARATOR    = "─" * _NORMAL_WIDTH
HEADER_LINE  = "═" * _NORMAL_WIDTH
NL_SEPARATOR = "─" * _NEWS_WIDTH
NL_HEADER    = "═" * _NEWS_WIDTH


def _truncate(text: str, max_len: int) -> str:
    if len(text) <= max_len:
        return text
    return text[: max_len - 1] + "…"


def _bar(count: int, max_count: int, width: int = COL_BAR) -> str:
    if max_count == 0:
        return ""
    filled = int((count / max_count) * width)
    return "█" * filled + "░" * (width - filled)


# ─── Header ──────────────────────────────────────────────────────────────────

def print_header() -> None:
    print()
    print(HEADER_LINE)
    print("  OUTLOOK EMAIL ANALYZER -- Ranking de remitentes")
    print(HEADER_LINE)
    print()


# ─── Tabla normal (sin newsletters) ─────────────────────────────────────────

def print_table(senders: Sequence[SenderStats], total: int) -> None:
    if not senders:
        print("  (No se encontraron correos para mostrar)")
        return

    max_count = senders[0].count

    print(SEPARATOR)
    print(
        f"  {'#':>{COL_NUM - 2}}  "
        f"{'Remitente':<{COL_NAME}}  "
        f"{'Email':<{COL_EMAIL}}  "
        f"{'Cantidad':>{COL_COUNT - 2}}  "
        f"{'Distribucion':<{COL_BAR}}"
    )
    print(SEPARATOR)

    for idx, s in enumerate(senders, 1):
        print(
            f"  {idx:>{COL_NUM - 2}}  "
            f"{_truncate(s.name, COL_NAME):<{COL_NAME}}  "
            f"{_truncate(s.email, COL_EMAIL):<{COL_EMAIL}}  "
            f"{s.count:>{COL_COUNT - 2},}  "
            f"{_bar(s.count, max_count)}"
        )

    print(SEPARATOR)


# ─── Tabla con newsletters ───────────────────────────────────────────────────

def print_table_newsletters(senders: Sequence[SenderStats], total: int) -> None:
    """Tabla que incluye columna de desuscripcion y link."""
    if not senders:
        print("  (No se encontraron correos para mostrar)")
        return

    print(NL_SEPARATOR)
    print(
        f"  {'#':>{COL_NUM - 2}}  "
        f"{'Remitente':<{COL_NAME}}  "
        f"{'Email':<{COL_EMAIL}}  "
        f"{'Cant.':>{COL_COUNT - 2}}  "
        f"{'Unsub':<{COL_UNSUB}}  "
        f"{'Link de desuscripcion':<{COL_LINK}}"
    )
    print(NL_SEPARATOR)

    for idx, s in enumerate(senders, 1):
        unsub_flag = "SI" if s.is_newsletter else "--"
        link = _truncate(s.unsubscribe_link, COL_LINK) if s.unsubscribe_link else ""

        print(
            f"  {idx:>{COL_NUM - 2}}  "
            f"{_truncate(s.name, COL_NAME):<{COL_NAME}}  "
            f"{_truncate(s.email, COL_EMAIL):<{COL_EMAIL}}  "
            f"{s.count:>{COL_COUNT - 2},}  "
            f"{unsub_flag:<{COL_UNSUB}}  "
            f"{link}"
        )

    print(NL_SEPARATOR)


# ─── Resumen ─────────────────────────────────────────────────────────────────

def print_summary(senders: Sequence[SenderStats], total: int, unread_only: bool) -> None:
    print()
    print("RESUMEN:")

    mode = "no leidos" if unread_only else "totales"
    print(f"  Correos analizados ({mode}):  {total:>10,}")
    print(f"  Remitentes unicos:           {len(senders):>10,}")

    if senders:
        top = senders[0]
        print(
            f"  Top remitente:               {top.name} "
            f"({top.count:,} correos, {top.percentage}%)"
        )

    print()
    print(HEADER_LINE)
    print()


def print_summary_newsletters(
    senders: Sequence[SenderStats], total: int, unread_only: bool
) -> None:
    """Resumen extendido con estadisticas de newsletters."""
    nl_senders = [s for s in senders if s.is_newsletter]
    nl_emails  = sum(s.count for s in nl_senders)
    non_nl     = len(senders) - len(nl_senders)

    print()
    print("RESUMEN:")

    mode = "no leidos" if unread_only else "totales"
    print(f"  Correos analizados ({mode}):  {total:>10,}")
    print(f"  Remitentes unicos:           {len(senders):>10,}")

    if senders:
        top = senders[0]
        print(
            f"  Top remitente:               {top.name} "
            f"({top.count:,} correos, {top.percentage}%)"
        )

    print()
    print(f"  NEWSLETTERS:")
    nl_pct = round(len(nl_senders) / len(senders) * 100, 1) if senders else 0
    mail_pct = round(nl_emails / total * 100, 1) if total else 0
    print(f"    Remitentes con desuscripcion: {len(nl_senders):>6} ({nl_pct}%)")
    print(f"    Correos de newsletters:       {nl_emails:>6,} ({mail_pct}% del total)")
    print(f"    Remitentes sin desuscripcion: {non_nl:>6}")

    print()
    print(NL_HEADER)
    print()


# ─── Progreso ────────────────────────────────────────────────────────────────

def print_progress(current: int, total: int) -> None:
    pct = int((current / total) * 100) if total else 0
    sys.stdout.write(f"\r  Procesando correo {current:,}/{total:,}... {pct}%")
    sys.stdout.flush()


def print_progress_done(total: int) -> None:
    sys.stdout.write(f"\r  Procesando correo {total:,}/{total:,}... 100%\n")
    sys.stdout.flush()
