"""
main.py
Email Analyzer para cuentas Microsoft (Hotmail, Outlook.com, Live.com).

Usa Microsoft Graph API — más simple y confiable que IMAP.
OAuth2 Device Code Flow: autorizas 1 vez en el navegador, después conecta solo.

Uso:
    python main.py                                    # analiza todo el inbox
    python main.py --unread-only                      # solo no leídos
    python main.py --top 20                           # top 20 remitentes
    python main.py --unread-only --top 20             # combinar
    python main.py --list-folders                     # ver carpetas
    python main.py --folder sentitems                 # otra carpeta
    python main.py --export resultados.csv            # exportar
    python main.py --reset-auth                       # reconfigurar OAuth
"""

import argparse
import csv
import os
import sys

import display
from analyzer import SenderStats, analyze
from ms_graph import MailData, authenticate, get_folder_info, list_folders, read_emails, reset_auth


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Analiza correos de Hotmail/Outlook.com y muestra ranking por remitente.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python main.py
  python main.py --unread-only --top 20
  python main.py --list-folders
  python main.py --export resultados.csv
  python main.py --reset-auth

Carpetas conocidas: inbox, sentitems, drafts, deleteditems, junkemail
        """,
    )
    parser.add_argument("--user", help="Email (ej: tu-correo@hotmail.com). Si no se da, se pide.")
    parser.add_argument("--unread-only", action="store_true", help="Solo correos no leídos.")
    parser.add_argument("--top", type=int, metavar="N", help="Mostrar solo top N remitentes.")
    parser.add_argument("--folder", default="inbox",
                        help="Carpeta a analizar (default: inbox). Usa --list-folders para ver opciones.")
    parser.add_argument("--list-folders", action="store_true", help="Listar carpetas y salir.")
    parser.add_argument("--export", metavar="FILE.csv", help="Exportar resultado a CSV.")
    parser.add_argument("--reset-auth", action="store_true", help="Borrar OAuth guardado y reconfigurar.")
    return parser.parse_args()


def _sanitize_csv(value: str) -> str:
    """
    Previene CSV Injection (fórmula injection en Excel).
    Un atacante podría enviar correos con nombre tipo =CMD|'/C calc'
    que Excel ejecutaría como fórmula al abrir el CSV.
    """
    if value and value[0] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + value
    return value


def export_csv(path: str, senders: list[SenderStats]) -> None:
    try:
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["#", "Remitente", "Email", "Cantidad", "Porcentaje"])
            for i, s in enumerate(senders, 1):
                writer.writerow([
                    i,
                    _sanitize_csv(s.name),
                    _sanitize_csv(s.email),
                    s.count,
                    f"{s.percentage}%",
                ])
        print(f"  ✓ Exportado a: {path}")
    except OSError as e:
        print(f"  Error al exportar: {e}", file=sys.stderr)


def main() -> None:
    args = parse_args()

    if args.reset_auth:
        reset_auth()
        return

    display.print_header()

    # ── Usuario ───────────────────────────────────────────────────────────────
    username = os.environ.get("EMAIL_USER") or args.user
    if not username:
        username = input("  Email (ej: tu-correo@hotmail.com): ").strip()
    if not username:
        raise SystemExit("\n❌  Debes ingresar tu dirección de correo.\n")

    # ── Autenticar ────────────────────────────────────────────────────────────
    print(f"\n  Conectando como {username}...")
    token = authenticate(username)

    # ── Listar carpetas ───────────────────────────────────────────────────────
    if args.list_folders:
        folders = list_folders(token)
        print("  Carpetas disponibles:\n")
        print(f"    {'Carpeta':<30} {'Total':>8} {'No leídos':>10}")
        print(f"    {'─' * 30} {'─' * 8} {'─' * 10}")
        for name, total, unread in folders:
            print(f"    {name:<30} {total:>8,} {unread:>10,}")
        print()
        return

    # ── Info de carpeta ───────────────────────────────────────────────────────
    folder_info = get_folder_info(token, args.folder)
    folder_name = folder_info["name"]
    total_in_folder = folder_info["total"]
    unread_count = folder_info["unread"]

    if total_in_folder == 0:
        print(f"  La carpeta '{folder_name}' está vacía.")
        return

    mode = "no leídos" if args.unread_only else "todos"
    expected = unread_count if args.unread_only else total_in_folder
    print(f"  Carpeta: {folder_name} | {total_in_folder:,} correos ({unread_count:,} no leídos)")
    print(f"  Modo: {mode} → analizando ~{expected:,} correos...\n")

    # ── Leer correos ──────────────────────────────────────────────────────────
    collected: list[MailData] = []

    def on_progress(fetched: int):
        if fetched % 100 == 0 or fetched == expected:
            display.print_progress(fetched, expected)

    try:
        for mail in read_emails(token, folder=args.folder,
                                unread_only=args.unread_only,
                                progress_callback=on_progress):
            collected.append(mail)
    except KeyboardInterrupt:
        print("\n\n  Interrumpido por el usuario.")

    if collected:
        display.print_progress_done(len(collected))
    print()

    if not collected:
        msg = "no leídos " if args.unread_only else ""
        print(f"  No se encontraron correos {msg}en '{folder_name}'.")
        return

    # ── Analizar y mostrar ────────────────────────────────────────────────────
    senders, total = analyze(collected, top=args.top)
    display.print_table(senders, total)
    display.print_summary(senders, total, args.unread_only)

    if args.export:
        export_csv(args.export, senders)


if __name__ == "__main__":
    main()
