"""Deslocamento de referencias em formulas.

Quando o Excel insere linhas, ele reescreve sozinho todas as formulas da pasta
de trabalho: uma faixa que terminava na linha 408 passa a terminar na 412 se
quatro linhas forem inseridas dentro dela. O openpyxl NAO faz isso - ele move
as celulas de lugar e deixa o texto das formulas intacto.

Este modulo reproduz a regra do Excel:

    inserindo N linhas na posicao P,
      - referencia de celula com linha >= P  ->  linha + N
      - faixa inicio:fim  ->  cada ponta segue a regra acima

O efeito colateral util e o mesmo do Excel: inserir DENTRO de uma faixa
(inclusive na ultima linha dela) estica a faixa; inserir depois do fim, nao.
"""

import re

# Nome de aba, com ou sem aspas, seguido de "!"
_ABA = r"(?:'(?:[^']|'')+'|[A-Za-z_À-ɏ][\w.À-ɏ]*)!"
# Referencia A1, com ou sem $
_REF = r"\$?[A-Za-z]{1,3}\$?\d{1,7}"

_PADRAO = re.compile(
    r"(?<![A-Za-z0-9_.$!])"          # nao e continuacao de outro identificador
    rf"({_ABA})?"                     # 1 - aba (opcional)
    rf"({_REF})"                      # 2 - referencia (ou inicio da faixa)
    rf"(?::({_REF}))?"                # 3 - fim da faixa (opcional)
    r"(?![A-Za-z0-9_])"               # nao e prefixo de outro identificador
    r"(?!\s*[\(\[])"                  # nao e nome de funcao nem tabela estruturada
)

_PARTES_REF = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$")


def _sem_aspas(nome):
    if nome.startswith("'") and nome.endswith("'"):
        return nome[1:-1].replace("''", "'")
    return nome


def _mover(ref, at, n):
    m = _PARTES_REF.match(ref)
    if not m:
        return ref
    cifrao_col, col, cifrao_lin, lin = m.groups()
    linha = int(lin)
    if linha >= at:
        linha += n
    return f"{cifrao_col}{col}{cifrao_lin}{linha}"


def deslocar_formula(formula, at, n, aba_alvo, aba_da_formula):
    """Devolve a formula com as referencias a `aba_alvo` deslocadas.

    at   - primeira linha inserida
    n    - quantas linhas foram inseridas
    aba_alvo        - aba onde as linhas entraram
    aba_da_formula  - aba onde esta a formula (para saber se uma referencia
                      sem prefixo aponta para a aba alvo ou para outra)
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    if n == 0:
        return formula

    mesma_aba = (aba_da_formula == aba_alvo)
    saida = []
    pos = 0

    # percorre pulando o conteudo entre aspas duplas (texto literal da formula)
    for trecho, literal in _fatiar(formula):
        if literal:
            saida.append(trecho)
            continue
        saida.append(_deslocar_trecho(trecho, at, n, aba_alvo, mesma_aba))
    return "".join(saida)


def _fatiar(formula):
    """Quebra a formula em pedacos (texto, e_literal)."""
    partes = []
    atual = []
    dentro = False
    i = 0
    while i < len(formula):
        ch = formula[i]
        if ch == '"':
            if dentro and i + 1 < len(formula) and formula[i + 1] == '"':
                atual.append('""')
                i += 2
                continue
            atual.append(ch)
            if dentro:
                partes.append(("".join(atual), True))
                atual = []
                dentro = False
            else:
                if len(atual) > 1:
                    partes.append(("".join(atual[:-1]), False))
                atual = ['"']
                dentro = True
            i += 1
            continue
        atual.append(ch)
        i += 1
    if atual:
        partes.append(("".join(atual), dentro))
    return partes


def _deslocar_trecho(trecho, at, n, aba_alvo, mesma_aba):
    def troca(m):
        aba, ini, fim = m.group(1), m.group(2), m.group(3)
        if aba:
            nome = _sem_aspas(aba[:-1])
            if nome != aba_alvo:
                return m.group(0)
        else:
            if not mesma_aba:
                return m.group(0)
        if fim:
            return f"{aba or ''}{_mover(ini, at, n)}:{_mover(fim, at, n)}"
        return f"{aba or ''}{_mover(ini, at, n)}"

    return _PADRAO.sub(troca, trecho)


def deslocar_faixa(faixa, at, n):
    """Desloca uma faixa solta do tipo 'A1:C10' ou 'B5' (intervalos mesclados,
    formatacao condicional, validacao de dados)."""
    if ":" in faixa:
        ini, fim = faixa.split(":", 1)
        return f"{_mover(ini, at, n)}:{_mover(fim, at, n)}"
    return _mover(faixa, at, n)
