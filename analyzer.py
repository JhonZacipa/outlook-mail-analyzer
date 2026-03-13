"""
analyzer.py
Lógica de agrupación y ranking de correos por remitente.
"""

from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Iterable

from ms_graph import MailData


@dataclass
class SenderStats:
    email: str
    name: str
    count: int
    percentage: float = 0.0
    is_newsletter: bool = False
    unsubscribe_link: str = ""


def analyze(emails: Iterable[MailData], top: int | None = None) -> tuple[list[SenderStats], int]:
    """
    Agrupa correos por remitente y retorna un ranking ordenado de mayor a menor.

    Args:
        emails: Iterable de MailData.
        top: Si se indica, retorna solo los primeros N resultados.

    Returns:
        Tupla (lista de SenderStats ordenada desc, total de correos procesados).
    """
    counter: Counter = Counter()

    # Conteo auxiliar para elegir el nombre mas representativo por email
    name_votes: dict[str, Counter] = defaultdict(Counter)

    # Guardar el primer link de desuscripcion encontrado por remitente
    # (los correos vienen ordenados por fecha desc, asi que el primero es el mas reciente)
    unsub_links: dict[str, str] = {}

    total = 0

    for mail in emails:
        email = mail.sender_email or "(sin direccion)"
        counter[email] += 1
        name_votes[email][mail.sender_name] += 1

        # Guardar link de desuscripcion (solo el primero/mas reciente)
        if mail.unsubscribe_link and email not in unsub_links:
            unsub_links[email] = mail.unsubscribe_link

        total += 1

    if total == 0:
        return [], 0

    # Resolver nombre mas frecuente por remitente
    names: dict[str, str] = {}
    for email, votes in name_votes.items():
        names[email] = votes.most_common(1)[0][0]

    # Construir lista ordenada
    ranked = [
        SenderStats(
            email=email,
            name=names.get(email, email),
            count=count,
            percentage=round((count / total) * 100, 1),
            is_newsletter=email in unsub_links,
            unsubscribe_link=unsub_links.get(email, ""),
        )
        for email, count in counter.most_common()
    ]

    if top is not None and top > 0:
        ranked = ranked[:top]

    return ranked, total
