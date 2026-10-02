"""Lancamento das remessas na planilha de faturamento.

Le a aba "Remessas Porto" e lanca cada nota na aba do mes correspondente
(09-26, 10-26, ...). Para cada nota, uma de tres coisas acontece:

    - ja lancada e completa     -> nao mexe
    - linha existe com so o NF  -> completa a linha ali mesmo
    - nao existe                -> insere na posicao certa pela data

Sobre o bloco de resumo: as formulas somam uma faixa fixa de linhas. No Excel
a insercao de linhas dentro dessa faixa fazia a formula esticar sozinha. Aqui
quem estica e o modulo `refs`, chamado por `planilha.inserir_linhas` - por
isso a rotina confere no fim se todas as faixas realmente cobrem os dados.

Diferenca em relacao a macro: como a insercao acontece na posicao certa pela
data, nota retroativa (um dia esquecido, lancado depois) entra no lugar dela
em vez de ir para o fim.
"""

import re
from collections import defaultdict
from copy import copy

import openpyxl

from .planilha import (
    ErroDeUso,
    achar_rodape,
    como_data,
    como_numero,
    fazer_backup,
    garantir_brl,
    gravar,
    inserir_linhas,
    texto,
)

MARCA_RESUMO = "SUB-TOTAL"
LIN_INI = 3

F_DATA, F_NF, F_VALOR = 1, 2, 3
F_UF, F_CFOP, F_OPER, F_DESCR, F_KG = 13, 14, 16, 17, 18
F_ULTCOL = 33

R_DATA, R_NF, R_TRANSP, R_LIQ, R_UNIT, R_TOTAL = 1, 2, 4, 5, 7, 8


# ------------------------------------------------------- leitura da remessa --
def ler_remessa(caminho):
    """{nf: (data, valor, kg)} da aba Remessas Porto.

    VL. TOTAL e formula (=G*E) e o arquivo nao guarda o resultado em cache,
    entao o valor e calculado aqui: unitario x peso liquido.
    """
    from .remessas import ABA

    wb = openpyxl.load_workbook(caminho, data_only=True)
    if ABA not in wb.sheetnames:
        wb.close()
        raise ErroDeUso(f'A planilha escolhida nao tem a aba "{ABA}".')
    ws = wb[ABA]

    rodape = achar_rodape(ws, coluna=R_TRANSP, ate=20000) or (ws.max_row + 1)

    notas = {}
    for r in range(2, rodape):
        nf = como_numero(ws.cell(r, R_NF).value)
        data = como_data(ws.cell(r, R_DATA).value)
        if nf is None or nf <= 0 or data is None or nf in notas:
            continue
        kg = como_numero(ws.cell(r, R_LIQ).value)
        valor = como_numero(ws.cell(r, R_TOTAL).value)
        if valor is None:
            unit = como_numero(ws.cell(r, R_UNIT).value)
            valor = (unit * kg) if (unit is not None and kg is not None) else None
        notas[nf] = (data, valor, kg)
    wb.close()

    if not notas:
        raise ErroDeUso(f'Nao encontrei notas na aba "{ABA}".')
    return notas


# -------------------------------------------------------------- principal --
def atualizar(cfg, relato=print):
    caminho_rem = cfg.remessa
    caminho_fat = cfg.faturamento

    relato(f"Lendo a remessa: {caminho_rem.name}")
    notas = ler_remessa(caminho_rem)
    relato(f"  {len(notas)} nota(s)")

    por_mes = defaultdict(dict)
    for nf, reg in notas.items():
        por_mes[reg[0].strftime("%m-%y")][nf] = reg

    relato(f"Abrindo o faturamento: {caminho_fat.name}")
    wb = openpyxl.load_workbook(caminho_fat)

    backup = fazer_backup(caminho_fat)
    relato(f"  copia de seguranca: Backup\\{backup.name}")

    total = {"completadas": 0, "inseridas": 0, "ja_lancadas": 0, "sem_aba": 0,
             "valor_atualizado": 0}
    avisos = []
    resumo_mes = []

    for mes in sorted(por_mes, key=lambda m: m[3:] + m[:2]):
        if mes not in wb.sheetnames:
            total["sem_aba"] += len(por_mes[mes])
            avisos.append(
                f'Aba "{mes}" nao existe: {len(por_mes[mes])} nota(s) nao lancada(s).'
            )
            continue
        parcial = _processar_mes(wb, wb[mes], mes, por_mes[mes], cfg, avisos)
        for chave in ("completadas", "inseridas", "ja_lancadas", "valor_atualizado"):
            total[chave] += parcial[chave]
        if sum(parcial.values()):
            resumo_mes.append(
                f"{mes}: {parcial['completadas']} completada(s), "
                f"{parcial['inseridas']} inserida(s), "
                f"{parcial['ja_lancadas']} ja lancada(s)"
                + (f", {parcial['valor_atualizado']} com valor atualizado"
                   if parcial["valor_atualizado"] else "")
            )
            relato(f"  {resumo_mes[-1]}")

    gravar(wb, caminho_fat)

    total["notas"] = len(notas)
    total["avisos"] = avisos
    total["por_mes"] = resumo_mes
    total["backup"] = backup
    return total


# -------------------------------------------------------------- um mes ----
def _processar_mes(wb, ws, mes, notas, cfg, avisos):
    conta = {"completadas": 0, "inseridas": 0, "ja_lancadas": 0, "valor_atualizado": 0}

    inicio_resumo = _achar_inicio_resumo(ws)
    if inicio_resumo == 0:
        avisos.append(
            f'Aba {mes}: nao achei o bloco de resumo ("{MARCA_RESUMO}" na coluna B). '
            "Nada foi lancado nela."
        )
        return conta

    existentes = _linhas_com_nota(ws, inicio_resumo)
    onde_esta = {}
    for linha, nf, _ in existentes:
        onde_esta.setdefault(nf, linha)
    ult_nota = existentes[-1][0] if existentes else LIN_INI - 1

    modelo = _achar_modelo(wb, ws, LIN_INI, ult_nota, cfg.operacao_faturamento)

    # --- completar o que ja existe, separar o que falta ---------------------
    a_inserir = {}
    for nf, reg in notas.items():
        if nf in onde_esta:
            linha = onde_esta[nf]
            if _esta_completa(ws, linha):
                conta["ja_lancadas"] += 1
                # VL. UNITARIO mudou na remessa depois do lancamento:
                # o valor da nota acompanha (so a coluna VALOR e tocada)
                novo = reg[1]
                atual = como_numero(ws.cell(linha, F_VALOR).value)
                if novo is not None and atual is not None and abs(novo - atual) > 0.005:
                    ws.cell(linha, F_VALOR).value = novo
                    garantir_brl(ws.cell(linha, F_VALOR))
                    conta["valor_atualizado"] += 1
            else:
                _escrever_nota(ws, linha, nf, reg, cfg, modelo)
                conta["completadas"] += 1
        else:
            a_inserir[nf] = reg

    if not a_inserir:
        return conta

    ordem = sorted(a_inserir, key=lambda nf: (a_inserir[nf][0], nf))

    # relê as datas DEPOIS de completar as linhas que so tinham o numero:
    # sem isso uma linha recem-completada nao conta na hora de decidir onde
    # cada nota nova entra, e a ordem de datas sai furada
    datas = [
        (linha, data)
        for linha, _, data in _linhas_com_nota(ws, inicio_resumo)
        if data
    ]

    no_meio = defaultdict(list)   # nota retroativa: entra antes de alguem
    no_fim = []
    for nf in ordem:
        data = a_inserir[nf][0]
        alvo = next((linha for linha, dt in datas if dt > data), None)
        if alvo is None:
            no_fim.append(nf)
        else:
            no_meio[alvo].append(nf)

    # --- retroativas, do alvo mais baixo na planilha para o mais alto -------
    for alvo in sorted(no_meio, reverse=True):
        grupo = no_meio[alvo]
        inserir_linhas(wb, ws, alvo, len(grupo))
        # a linha-modelo foi escolhida antes da insercao: se estava daqui para
        # baixo, desceu junto. Sem o ajuste, o estilo viria de uma linha recem-
        # inserida (vazia) e a data sairia como numero cru.
        if modelo is not None and modelo[0] is ws and modelo[1] >= alvo:
            modelo = (ws, modelo[1] + len(grupo))
        for i, nf in enumerate(grupo):
            _escrever_nota(ws, alvo + i, nf, a_inserir[nf], cfg, modelo)
        conta["inseridas"] += len(grupo)
        ult_nota += len(grupo)
        inicio_resumo += len(grupo)

    # --- as do fim: usa as linhas em branco, cria so o que faltar -----------
    if no_fim:
        livres = max(0, (inicio_resumo - 1) - ult_nota)
        faltam = len(no_fim) - livres

        if faltam <= 0:
            onde = ult_nota + 1

        elif ult_nota >= LIN_INI:
            # inserir DENTRO da faixa somada (na ultima nota) e devolver essa
            # nota para o lugar dela, preservando a ordem de datas
            inserir_linhas(wb, ws, ult_nota, faltam)
            _mover_linha(ws, ult_nota + faltam, ult_nota)
            inicio_resumo += faltam
            onde = ult_nota + 1

        else:
            # aba vazia: criar a partir da SEGUNDA linha da faixa, nunca da
            # primeira, senao as notas novas caem fora da soma
            inserir_linhas(wb, ws, LIN_INI + 1, faltam)
            inicio_resumo += faltam
            onde = LIN_INI

        for i, nf in enumerate(no_fim):
            _escrever_nota(ws, onde + i, nf, a_inserir[nf], cfg, modelo)
        conta["inseridas"] += len(no_fim)
        ult_nota = onde + len(no_fim) - 1

    # --- conferir se o resumo acompanhou -----------------------------------
    fim_min = _menor_fim_do_resumo(ws, _achar_inicio_resumo(ws))
    if fim_min and fim_min < ult_nota:
        avisos.append(
            f"Aba {mes}: ha formula(s) no resumo somando so ate a linha {fim_min}, "
            f"mas os dados vao ate {ult_nota}. Confira essas faixas antes de usar "
            "os totais."
        )

    return conta


# ---------------------------------------------------------------- apoio ----
def _achar_inicio_resumo(ws):
    for r in range(LIN_INI, min(ws.max_row, 20000) + 1):
        v = ws.cell(r, F_NF).value
        if isinstance(v, str) and v.strip().upper() == MARCA_RESUMO:
            return r
    return 0


def _linhas_com_nota(ws, inicio_resumo):
    saida = []
    for r in range(LIN_INI, inicio_resumo):
        nf = como_numero(ws.cell(r, F_NF).value)
        if nf is None:
            continue
        saida.append((r, nf, como_data(ws.cell(r, F_DATA).value)))
    return saida


def _esta_completa(ws, linha):
    return (
        como_data(ws.cell(linha, F_DATA).value) is not None
        and como_numero(ws.cell(linha, F_VALOR).value) is not None
    )


def _escrever_nota(ws, linha, nf, reg, cfg, modelo):
    data, valor, kg = reg
    if modelo is not None:
        ws_modelo, lin_modelo = modelo
        if not (ws_modelo is ws and lin_modelo == linha):
            for c in range(1, F_ULTCOL + 1):
                ws.cell(linha, c)._style = copy(ws_modelo.cell(lin_modelo, c)._style)
    ws.cell(linha, F_DATA).value = data
    ws.cell(linha, F_NF).value = nf
    ws.cell(linha, F_VALOR).value = valor
    garantir_brl(ws.cell(linha, F_VALOR))
    ws.cell(linha, F_UF).value = cfg.uf
    ws.cell(linha, F_CFOP).value = cfg.cfop_faturamento
    ws.cell(linha, F_OPER).value = cfg.operacao_faturamento
    ws.cell(linha, F_DESCR).value = cfg.descricao
    ws.cell(linha, F_KG).value = kg


def _mover_linha(ws, de, para):
    """Leva valores e formatacao de uma linha para outra e limpa a origem."""
    for c in range(1, F_ULTCOL + 1):
        origem = ws.cell(de, c)
        destino = ws.cell(para, c)
        destino.value = origem.value
        destino._style = copy(origem._style)
        origem.value = None


def _achar_modelo(wb, ws, ini, fim, operacao):
    """Linha-modelo de formatacao: uma linha da mesma operacao ja lancada.

    Procura primeiro no proprio mes; se o mes estiver vazio, pega a de outro
    mes - senao a data sairia como numero cru.
    """
    achado = _modelo_na_aba(ws, ini, fim, operacao)
    if achado:
        return achado
    for outra in wb.worksheets:
        if outra is ws or not _eh_aba_de_mes(outra.title):
            continue
        limite = _achar_inicio_resumo(outra)
        if limite > LIN_INI:
            achado = _modelo_na_aba(outra, LIN_INI, limite - 1, operacao)
            if achado:
                return achado
    return None


def _modelo_na_aba(ws, ini, fim, operacao):
    alvo = operacao.strip().upper()
    for r in range(ini, fim + 1):
        if texto(ws.cell(r, F_OPER).value).upper() != alvo:
            continue
        if como_data(ws.cell(r, F_DATA).value) is not None:
            return (ws, r)
    return None


def _eh_aba_de_mes(nome):
    return bool(re.fullmatch(r"(0[1-9]|1[0-2])-\d{2}", nome))


_FIM_DA_FAIXA = re.compile(r"\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?(\d+)")


def _menor_fim_do_resumo(ws, inicio_resumo):
    """Menor linha final entre as faixas somadas pelo bloco de resumo."""
    if not inicio_resumo:
        return 0
    menor = 0
    for r in range(inicio_resumo, min(inicio_resumo + 40, ws.max_row + 1)):
        for c in range(3, 13):
            f = ws.cell(r, c).value
            if not isinstance(f, str) or "SUMIF" not in f.upper():
                continue
            for m in _FIM_DA_FAIXA.finditer(f):
                n = int(m.group(1))
                if menor == 0 or n < menor:
                    menor = n
    return menor
