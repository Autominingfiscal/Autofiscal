"""Comparacao de textos digitados de jeitos diferentes."""

import re
import unicodedata


def sem_acento(texto):
    """'Líquido Ç' -> 'Liquido C'. None vira ''."""
    if texto is None:
        return ""
    s = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in s if not unicodedata.combining(c))


def norm(texto):
    """Maiusculas, sem acento, so letras e numeros: 'PLACA DO \\nCAVALO' -> 'PLACADOCAVALO'.

    Serve para achar cabecalho de coluna, placa e modelo de caminhao mesmo com
    acento, espaco ou traco diferente."""
    return re.sub(r"[^A-Z0-9]", "", sem_acento(texto).upper())
