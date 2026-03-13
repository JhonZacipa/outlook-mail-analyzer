"""
display.py
Formateo y presentación del ranking en consola.
"""

import sys
from typing import Sequence

from analyzer import SenderStats

# Anchos de columna
COL_NUM = 5
COL_NAME = 30
COL_EMAIL = 38
COL_COUNT = 10
COL_BAR = 20

SEPARATOR = "─" * (COL_NUM + COL_NAME + COL_EMAIL + COL_COUNT + COL_BAR + 10)
HEADER_LINE = "═" * len(SEPARATOR)


def _truncate(text: str, max_len: int) -> str:
    """Trunca texto con elipsis si supera max_len."""
    if len(text) <= max_len:
        return text
    return text[: max_len - 1] + "…"


def _bar(count: int, max_count: int, width: int = COL_BAR) -> str:
    """Genera una barra ASCII proporcional."""
    if max_count == 0:
        return ""
    filled = int((count / max_count) * width)
    return "█" * filled + "░" * (width - filled)


def print_header() -> None:
    print()
    print(HEADER_LINE)
    print("  OUTLOOK EMAIL ANALYZER — Ranking de remitentes")
    print(HEADER_LINE)
    print()


def print_table(senders: Sequence[SenderStats], total: int) -> None:
    """Imprime la tabla de ranking."""
    if not senders:
        print("  (No se encontraron correos para mostrar)")
        return

    max_count = senders[0].count if senders else 1

    # Encabezado de tabla
    print(SEPARATOR)
    header = (
        f"  {'#':>{COL_NUM - 2}}  "
        f"{'Remitente':<{COL_NAME}}  "
        f"{'Email':<{COL_EMAIL}}  "
        f"{'Cantidad':>{COL_COUNT - 2}}  "
        f"{'Distribución':<{COL_BAR}}"
    )
    print(header)
    print(SEPARATOR)

    # Filas
    for idx, sender in enumerate(senders, start=1):
        num = f"{idx:>{COL_NUM - 2}}"
        name = _truncate(sender.name, COL_NAME)
        email = _truncate(sender.email, COL_EMAIL)
        count = f"{sender.count:>{COL_COUNT - 2},}"
        bar = _bar(sender.count, max_count)

        print(
            f"  {num}  "
            f"{name:<{COL_NAME}}  "
            f"{email:<{COL_EMAIL}}  "
            f"{count}  "
            f"{bar}"
        )

    print(SEPARATOR)


def print_summary(senders: Sequence[SenderStats], total: int, unread_only: bool) -> None:
    """Imprime el resumen al final."""
    print()
    print("RESUMEN:")

    mode = "no leídos" if unread_only else "totales"
    print(f"  Correos analizados ({mode}):  {total:>10,}")
    print(f"  Remitentes únicos:           {len(senders):>10,}")

    if senders:
        top = senders[0]
        print(
            f"  Top remitente:               {top.name} "
            f"({top.count:,} correos, {top.percentage}%)"
        )

    print()
    print(HEADER_LINE)
    print()


def print_progress(current: int, total: int) -> None:
    """Imprime progreso en la misma línea."""
    pct = int((current / total) * 100) if total else 0
    sys.stdout.write(f"\r  Procesando correo {current:,}/{total:,}... {pct}%")
    sys.stdout.flush()


def print_progress_done(total: int) -> None:
    """Finaliza la línea de progreso."""
    sys.stdout.write(f"\r  Procesando correo {total:,}/{total:,}... 100% ✓\n")
    sys.stdout.flush()
