"""Numeros escritos do jeito brasileiro, como aparecem nas planilhas."""

import math
import re

_MILHAR = re.compile(r"-?\d{1,3}(\.\d{3})+")


def numero_br(valor):
    """Devolve float se a celula tiver numero utilizavel, senao None.

    Numero de verdade passa direto. Texto e lido no jeito brasileiro:
        "R$ 1.234,56" -> 1234.56     "14,80" -> 14.8      "1.234" -> 1234
        "14.80"       -> 14.8        "(10,00)" -> -10     "1 234,5" -> 1234.5
    Vazio, "-", formula ("=A1"), erro do Excel ("#REF!") e texto -> None.
    """
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor) if math.isfinite(valor) else None
    texto = (str(valor).replace("R$", "").replace(" ", "")
             .replace(" ", "").strip())
    if not texto or texto.startswith(("=", "#")) or set(texto) <= {"-"}:
        return None
    negativo = texto.startswith("(") and texto.endswith(")")
    if negativo:
        texto = texto[1:-1]
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif _MILHAR.fullmatch(texto):
        texto = texto.replace(".", "")
    try:
        n = float(texto)
    except ValueError:
        return None
    if not math.isfinite(n):            # "nan", "inf" escritos na celula
        return None
    return -n if negativo else n
