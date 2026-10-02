"""Apoio para mexer nas planilhas sem estragar o que ja esta la."""

import datetime
import os
import re
import shutil
from copy import copy
from pathlib import Path

from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import get_column_letter

from .refs import deslocar_faixa, deslocar_formula

MARCA_RODAPE = "VOLUME EMBARCADO"

# Formato Contabil em Real, o mesmo que as planilhas ja usam. O codigo [$R$-416]
# e o "Moeda: R$ - Portugues (Brasil)" do Excel: aparece como R$ 1.234,56.
FORMATO_BRL = (
    '_-[$R$-416]\\ * #,##0.00_-;\\-[$R$-416]\\ * #,##0.00_-;'
    '_-[$R$-416]\\ * "-"??_-;_-@_-'
)


class ErroDeUso(Exception):
    """Problema que o usuario consegue resolver - mostrado sem traceback."""


# ---------------------------------------------------------------- backup ----
MANTER_BACKUPS = 30     # copias guardadas de CADA planilha; as mais velhas saem


def fazer_backup(caminho, manter=MANTER_BACKUPS):
    """Grava uma copia do arquivo na subpasta Backup e devolve o caminho.

    Depois apaga as copias mais antigas DESTA planilha alem das `manter` mais
    recentes. So mexe em arquivos com o carimbo que esta funcao grava."""
    caminho = Path(caminho)
    pasta = caminho.parent / "Backup"
    pasta.mkdir(exist_ok=True)
    carimbo = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    destino = pasta / f"{carimbo}_{caminho.name}"
    n = 2
    while destino.exists():     # dois backups no mesmo segundo (rodar tudo)
        destino = pasta / f"{carimbo}_{caminho.stem} ({n}){caminho.suffix}"
        n += 1
    shutil.copy2(caminho, destino)
    _limpar_backups(pasta, caminho.name, manter)
    return destino


def _limpar_backups(pasta, nome, manter):
    base, ext = os.path.splitext(nome)
    desta = re.compile(r"^\d{4}-\d{2}-\d{2}_\d{6}_" + re.escape(base)
                       + r"(?: \(\d+\))?" + re.escape(ext) + "$")
    copias = sorted(p for p in pasta.iterdir() if p.is_file() and desta.match(p.name))
    for velha in copias[:-manter] if manter > 0 else []:
        try:
            velha.unlink()
        except OSError:
            pass        # aberta no Excel, por exemplo: sai na proxima vez


def gravar(wb, caminho):
    """Salva a pasta de trabalho, com recado claro se ela estiver aberta."""
    caminho = Path(caminho)
    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass
    try:
        wb.save(caminho)
    except PermissionError:
        raise ErroDeUso(
            f"Nao consegui salvar '{caminho.name}'.\n"
            "O arquivo provavelmente esta aberto no Excel. "
            "Feche a planilha e rode de novo."
        )


# ----------------------------------------------------------- formatacao ----
def copiar_formato(ws_origem, lin_origem, ws_destino, lin_destino, col_ate):
    """Copia a formatacao de uma linha inteira para outra (so o formato)."""
    if ws_origem is ws_destino and lin_origem == lin_destino:
        return
    for c in range(1, col_ate + 1):
        origem = ws_origem.cell(lin_origem, c)
        destino = ws_destino.cell(lin_destino, c)
        destino._style = copy(origem._style)


def guardar_formato(ws, linha, col_ate):
    """Fotografa a formatacao de uma linha para aplicar depois."""
    return [copy(ws.cell(linha, c)._style) for c in range(1, col_ate + 1)]


def aplicar_formato(ws, linha, estilos):
    for i, estilo in enumerate(estilos, start=1):
        ws.cell(linha, i)._style = copy(estilo)


def garantir_brl(celula):
    """Poe a celula em Real se ela ainda nao estiver num formato com R$.

    Formato que ja tem R$ (com 2 ou 4 casas, contabil ou moeda) e mantido.
    """
    if "R$" not in (celula.number_format or ""):
        celula.number_format = FORMATO_BRL


def limpar_conteudo(ws, lin_ini, lin_fim, col_ate):
    """Apaga os valores mantendo a formatacao, como o ClearContents do Excel."""
    for r in range(lin_ini, lin_fim + 1):
        for c in range(1, col_ate + 1):
            ws.cell(r, c).value = None


# -------------------------------------------------------- inserir linhas ----
def inserir_linhas(wb, ws, at, n):
    """Insere n linhas na posicao `at` fazendo o que o Excel faria.

    Alem de empurrar as celulas (que e o unico servico do openpyxl), ajusta:
      - as formulas de TODAS as abas que apontam para esta;
      - os intervalos mesclados;
      - as faixas de formatacao condicional e de validacao de dados;
      - as alturas de linha;
      - os nomes definidos da pasta de trabalho.
    """
    if n <= 0:
        return

    alturas = _guardar_alturas(ws)
    ws.insert_rows(at, n)
    _restaurar_alturas(ws, alturas, at, n)

    alvo = ws.title
    for outra in wb.worksheets:
        for linha in outra.iter_rows():
            for celula in linha:
                if isinstance(celula.value, str) and celula.value.startswith("="):
                    celula.value = deslocar_formula(
                        celula.value, at, n, alvo, outra.title
                    )

    _deslocar_mescladas(ws, at, n)
    _deslocar_condicional(ws, at, n)
    _deslocar_validacao(ws, at, n)
    _deslocar_nomes(wb, at, n, alvo)


def _guardar_alturas(ws):
    return {
        idx: dim.height
        for idx, dim in ws.row_dimensions.items()
        if dim.height is not None
    }


def _restaurar_alturas(ws, alturas, at, n):
    for idx in list(ws.row_dimensions):
        if idx >= at:
            ws.row_dimensions[idx].height = None
    for idx, altura in sorted(alturas.items(), reverse=True):
        novo = idx + n if idx >= at else idx
        ws.row_dimensions[novo].height = altura


def _deslocar_mescladas(ws, at, n):
    antigas = [str(m) for m in ws.merged_cells.ranges]
    novas = [deslocar_faixa(f, at, n) for f in antigas]
    if novas == antigas:
        return
    for faixa in antigas:
        ws.unmerge_cells(faixa)
    for faixa in novas:
        ws.merge_cells(faixa)


def _deslocar_condicional(ws, at, n):
    """Reconstroi a lista inteira.

    Nao da para so trocar o sqref de cada regra: o openpyxl guarda as regras
    num dicionario cuja chave e o proprio objeto, e o hash dele sai do sqref.
    Mexer no sqref no lugar quebra o dicionario, e o erro so aparece na hora
    de salvar.
    """
    itens = [
        (str(cf.sqref), list(regras))
        for cf, regras in ws.conditional_formatting._cf_rules.items()
    ]
    if not itens:
        return
    nova = ConditionalFormattingList()
    for sqref, regras in itens:
        faixa = " ".join(deslocar_faixa(f, at, n) for f in sqref.split())
        for regra in regras:
            nova.add(faixa, regra)
    ws.conditional_formatting = nova


def _deslocar_validacao(ws, at, n):
    for dv in ws.data_validations.dataValidation:
        faixas = [deslocar_faixa(str(f), at, n) for f in dv.sqref.ranges]
        dv.sqref = " ".join(faixas)


def _deslocar_nomes(wb, at, n, alvo):
    for nome in wb.defined_names.values():
        try:
            nome.value = deslocar_formula("=" + nome.value, at, n, alvo, None)[1:]
        except Exception:
            pass


# ------------------------------------------------------------- leituras ----
def achar_rodape(ws, coluna=4, marca=MARCA_RODAPE, ate=5000):
    """Linha do 'VOLUME EMBARCADO', que fecha a area de dados. 0 se nao houver."""
    for r in range(min(ws.max_row, ate), 1, -1):
        v = ws.cell(r, coluna).value
        if isinstance(v, str) and v.strip().upper() == marca:
            return r
    return 0


def como_numero(valor):
    """Devolve float se a celula tiver numero utilizavel, senao None.

    Numero de verdade passa direto. Texto e lido no jeito brasileiro:
    "R$ 1.234,56", "14,80" e "1.234" viram 1234.56, 14.8 e 1234.
    """
    if isinstance(valor, bool) or valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        texto = valor.replace("R$", "").replace("\u00a0", "").replace(" ", "").strip()
        if not texto or texto.startswith("="):
            return None
        if "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", texto):
            texto = texto.replace(".", "")
        try:
            return float(texto)
        except ValueError:
            return None
    return None


def como_data(valor):
    if isinstance(valor, datetime.datetime):
        return valor.date()
    if isinstance(valor, datetime.date):
        return valor
    return None


def texto(valor):
    if valor is None:
        return ""
    return str(valor).strip()


def coluna(n):
    return get_column_letter(n)
