"""Chave de acesso da NF-e (44 digitos).

    posicao  0-1  UF        2-5  AAMM       6-19  CNPJ do emitente
            20-21 modelo   22-24 serie     25-33 numero da nota
            34    tipo de emissao          35-42 codigo     43  digito verificador
"""

import re


def so_digitos(texto):
    return re.sub(r"\D", "", texto or "")


def dv_confere(chave):
    """Digito verificador: modulo 11, pesos 2..9 da direita para a esquerda."""
    if len(chave) != 44 or not chave.isdigit():
        return False
    soma, peso = 0, 2
    for ch in reversed(chave[:43]):
        soma += int(ch) * peso
        peso = peso + 1 if peso < 9 else 2
    dv = 11 - soma % 11
    return (0 if dv >= 10 else dv) == int(chave[43])


def nota_da_chave(chave):
    """Numero da nota embutido na chave (posicoes 26 a 34)."""
    return int(chave[25:34])


def modelo_da_chave(chave):
    """'55' = NF-e, '65' = NFC-e."""
    return chave[20:22]


def formatar(chave):
    """11 grupos de 4 digitos, como e digitado nas planilhas."""
    return " ".join(chave[i:i + 4] for i in range(0, 44, 4))
