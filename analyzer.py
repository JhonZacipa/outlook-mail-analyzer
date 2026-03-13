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
    names: dict[str, str] = {}  # email -> nombre más frecuente

    # Conteo auxiliar para elegir el nombre más representativo por email
    name_votes: dict[str, Counter] = defaultdict(Counter)

    total = 0

    for mail in emails:
        email = mail.sender_email or "(sin dirección)"
        counter[email] += 1
        name_votes[email][mail.sender_name] += 1
        total += 1

    if total == 0:
        return [], 0

    # Resolver nombre más frecuente por remitente
    for email, votes in name_votes.items():
        names[email] = votes.most_common(1)[0][0]

    # Construir lista ordenada
    ranked = [
        SenderStats(
            email=email,
            name=names.get(email, email),
            count=count,
            percentage=round((count / total) * 100, 1),
        )
        for email, count in counter.most_common()
    ]

    if top is not None and top > 0:
        ranked = ranked[:top]

    return ranked, total
