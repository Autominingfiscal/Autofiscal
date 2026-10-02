"""Leitura da planilha de expedicao (EXPED_CONCENTRADO).

Cada aba com nome no padrao "dd.mm" e um dia de carregamento. A tabela
principal comeca na linha 3, e as colunas usadas sao:

    D (4)  Nº_NOTA        F (6)  CODIGO LACRE     G (7)  TRANSPORTADORA
    I (9)  PLACA CAVALO   N (14) PESO_ENTRADA     O (15) PESO_SAIDA
    P (16) PESO_LIQUIDO

Detalhe importante do porte: PESO_LIQUIDO e uma formula
(PESO_SAIDA - PESO_ENTRADA) e o arquivo nao guarda o resultado em cache,
entao o valor e calculado aqui em vez de lido.
"""

import datetime

from .planilha import ErroDeUso, como_numero, texto

E_NOTA, E_LACRE, E_TRANSP, E_PLACA = 4, 6, 7, 9
E_ENTRADA, E_SAIDA, E_LIQUIDO = 14, 15, 16

LIN_INI = 3


class Nota:
    __slots__ = ("nf", "data", "placa", "transportadora", "lacre", "liquido")

    def __init__(self, nf, data, placa, transportadora, lacre, liquido):
        self.nf = nf
        self.data = data
        self.placa = placa
        self.transportadora = transportadora
        self.lacre = lacre
        self.liquido = liquido


def nome_eh_data(nome):
    """Aceita nomes de aba no padrao 'dd.mm'."""
    if "." not in nome:
        return False
    partes = nome.split(".")
    if len(partes) != 2:
        return False
    try:
        dia, mes = int(partes[0]), int(partes[1])
    except ValueError:
        return False
    return 1 <= dia <= 31 and 1 <= mes <= 12


def data_da_aba(nome, ano):
    dia, mes = nome.split(".")
    return datetime.date(ano, int(mes), int(dia))


def peso_liquido(ws, r):
    """PESO_LIQUIDO da linha: usa o valor se houver, senao saida - entrada."""
    direto = como_numero(ws.cell(r, E_LIQUIDO).value)
    if direto is not None:
        return direto
    entrada = como_numero(ws.cell(r, E_ENTRADA).value)
    saida = como_numero(ws.cell(r, E_SAIDA).value)
    if entrada is None or saida is None:
        return None
    return saida - entrada


def ler(wb, ano):
    """Devolve {nf: Nota} com as notas de todas as abas de dia.

    Igual ao VBA: a primeira ocorrencia de uma NF e a que vale.
    """
    notas = {}
    sem_peso = []

    for ws in wb.worksheets:
        if not nome_eh_data(ws.title):
            continue
        data = data_da_aba(ws.title, ano)

        for r in range(LIN_INI, ws.max_row + 1):
            nf = como_numero(ws.cell(r, E_NOTA).value)
            if nf is None or nf <= 0 or nf in notas:
                continue
            liquido = peso_liquido(ws, r)
            if liquido is None:
                sem_peso.append((ws.title, r, int(nf)))
                continue
            notas[nf] = Nota(
                nf=nf,
                data=data,
                placa=texto(ws.cell(r, E_PLACA).value),
                transportadora=texto(ws.cell(r, E_TRANSP).value),
                lacre=texto(ws.cell(r, E_LACRE).value),
                liquido=liquido,
            )

    if not notas:
        raise ErroDeUso(
            "Nao encontrei nenhuma nota na planilha de expedicao.\n"
            "Confira se o caminho aponta para o arquivo certo e se as abas "
            "de dia estao no padrao 'dd.mm'."
        )
    return notas, sem_peso
