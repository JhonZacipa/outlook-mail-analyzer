"""
main.py
Email Analyzer para cuentas Microsoft (Hotmail, Outlook.com, Live.com).

Usa Microsoft Graph API. OAuth2 Device Code Flow: autorizas 1 vez en el navegador.

Uso:
    python main.py                                    # analiza todo el inbox
    python main.py --unread-only                      # solo no leidos
    python main.py --top 20                           # top 20 remitentes
    python main.py --newsletters                      # detectar newsletters + link unsub
    python main.py --newsletters --top 20             # combinar
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
from ms_graph import (
    MailData, authenticate, fetch_unsubscribe_links,
    get_folder_info, list_folders, read_emails, reset_auth,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Analiza correos de Hotmail/Outlook.com y muestra ranking por remitente.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python main.py
  python main.py --unread-only --top 20
  python main.py --newsletters
  python main.py --newsletters --top 10 --export newsletters.csv
  python main.py --list-folders
  python main.py --reset-auth

Carpetas conocidas: inbox, sentitems, drafts, deleteditems, junkemail
        """,
    )
    parser.add_argument("--user", help="Email (ej: tu-correo@hotmail.com). Si no se da, se pide.")
    parser.add_argument("--unread-only", action="store_true", help="Solo correos no leidos.")
    parser.add_argument("--newsletters", action="store_true",
                        help="Detectar newsletters y mostrar link de desuscripcion por remitente.")
    parser.add_argument("--top", type=int, metavar="N", help="Mostrar solo top N remitentes.")
    parser.add_argument("--folder", default="inbox",
                        help="Carpeta a analizar (default: inbox). Usa --list-folders para ver opciones.")
    parser.add_argument("--list-folders", action="store_true", help="Listar carpetas y salir.")
    parser.add_argument("--export", metavar="FILE.csv", help="Exportar resultado a CSV.")
    parser.add_argument("--reset-auth", action="store_true", help="Borrar OAuth guardado y reconfigurar.")
    return parser.parse_args()


def _sanitize_csv(value: str) -> str:
    """
    Previene CSV Injection (formula injection en Excel).
    Un atacante podria enviar correos con nombre tipo =CMD|'/C calc'
    que Excel ejecutaria como formula al abrir el CSV.
    """
    if value and value[0] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + value
    return value


def export_csv(path: str, senders: list[SenderStats], newsletters_mode: bool) -> None:
    try:
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)

            if newsletters_mode:
                writer.writerow(["#", "Remitente", "Email", "Cantidad", "Porcentaje",
                                 "Newsletter", "Link desuscripcion"])
            else:
                writer.writerow(["#", "Remitente", "Email", "Cantidad", "Porcentaje"])

            for i, s in enumerate(senders, 1):
                row = [
                    i,
                    _sanitize_csv(s.name),
                    _sanitize_csv(s.email),
                    s.count,
                    f"{s.percentage}%",
                ]
                if newsletters_mode:
                    row.append("SI" if s.is_newsletter else "NO")
                    row.append(_sanitize_csv(s.unsubscribe_link))
                writer.writerow(row)

        print(f"  Exportado a: {path}")
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
        raise SystemExit("\n  Debes ingresar tu direccion de correo.\n")

    # ── Autenticar ────────────────────────────────────────────────────────────
    print(f"\n  Conectando como {username}...")
    token = authenticate(username)

    # ── Listar carpetas ───────────────────────────────────────────────────────
    if args.list_folders:
        folders = list_folders(token)
        print("  Carpetas disponibles:\n")
        print(f"    {'Carpeta':<30} {'Total':>8} {'No leidos':>10}")
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
        print(f"  La carpeta '{folder_name}' esta vacia.")
        return

    mode = "no leidos" if args.unread_only else "todos"
    expected = unread_count if args.unread_only else total_in_folder

    extra = " + deteccion de newsletters" if args.newsletters else ""
    print(f"  Carpeta: {folder_name} | {total_in_folder:,} correos ({unread_count:,} no leidos)")
    print(f"  Modo: {mode}{extra} -- analizando ~{expected:,} correos...\n")

    # ── Leer correos ──────────────────────────────────────────────────────────
    collected: list[MailData] = []

    def on_progress(fetched: int):
        if fetched % 100 == 0 or fetched == expected:
            display.print_progress(fetched, expected)

    try:
        for mail in read_emails(token, folder=args.folder,
                                unread_only=args.unread_only,
                                detect_newsletters=args.newsletters,
                                progress_callback=on_progress):
            collected.append(mail)
    except KeyboardInterrupt:
        print("\n\n  Interrumpido por el usuario.")

    if collected:
        display.print_progress_done(len(collected))
    print()

    if not collected:
        msg = "no leidos " if args.unread_only else ""
        print(f"  No se encontraron correos {msg}en '{folder_name}'.")
        return

    # ── Analizar y mostrar ────────────────────────────────────────────────────
    senders, total = analyze(collected, top=args.top)

    if args.newsletters:
        # Segundo pase: buscar links REALES de desuscripcion en el body HTML
        # Solo 1 request por remitente (no por correo), asi que es rapido
        sender_emails = [s.email for s in senders]
        print(f"  Buscando links de desuscripcion en {len(sender_emails)} remitentes...\n")

        def on_unsub_progress(current: int, total: int):
            sys.stdout.write(f"\r  Escaneando remitente {current}/{total}...")
            sys.stdout.flush()

        unsub_links = fetch_unsubscribe_links(
            token, sender_emails, folder=args.folder,
            progress_callback=on_unsub_progress,
        )
        sys.stdout.write(f"\r  Escaneando remitentes... {len(unsub_links)} links encontrados\n\n")

        # Actualizar senders con los links reales del body
        updated = []
        for s in senders:
            body_link = unsub_links.get(s.email, "")
            updated.append(SenderStats(
                email=s.email,
                name=s.name,
                count=s.count,
                percentage=s.percentage,
                is_newsletter=bool(body_link) or s.is_newsletter,
                unsubscribe_link=body_link or s.unsubscribe_link,
            ))
        senders = updated

        display.print_table_newsletters(senders, total)
        display.print_summary_newsletters(senders, total, args.unread_only)
    else:
        display.print_table(senders, total)
        display.print_summary(senders, total, args.unread_only)

    if args.export:
        export_csv(args.export, senders, args.newsletters)


if __name__ == "__main__":
    main()
