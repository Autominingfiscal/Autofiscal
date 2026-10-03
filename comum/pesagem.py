"""Regras do carregamento: peso bruto alvo por modelo e faixa de aceite.

O Lancador pinta a celula do PESO_SAIDA e o Dashboard mostra a mesma faixa:
os dois usam estas funcoes, para nunca discordarem.
"""

from .texto import norm

# faixas devolvidas por faixa_de_aceite
VERMELHO, AMARELO, VERDE, DENTRO = "vermelho", "amarelo", "verde", "dentro"


def peso_alvo(alvos, modelo):
    """Peso bruto alvo do modelo ('LS 4 EIXOS', 'vanderleia'...). None se nao achar.

    `alvos` e {modelo: peso}, com o nome do modelo escrito de qualquer jeito.
    Primeiro procura o nome igual; depois um que contenha o outro."""
    m = norm(modelo)
    if not m:
        return None
    por_nome = {norm(k): v for k, v in alvos.items()}
    if m in por_nome:
        return por_nome[m]
    for k, v in por_nome.items():
        if k and (k in m or m in k):
            return v
    return None


def faixa_de_aceite(bruto, alvo, tol_excesso, tol_subcarga):
    """Compara o peso bruto com o alvo.

        VERMELHO  excesso acima de tol_excesso
        AMARELO   excesso de ate tol_excesso
        VERDE     subcarregado: mais de tol_subcarga abaixo do alvo
        DENTRO    entre (alvo - tol_subcarga) e o alvo
    """
    dif = bruto - alvo
    if dif > tol_excesso:
        return VERMELHO
    if dif > 0:
        return AMARELO
    if -dif > tol_subcarga:
        return VERDE
    return DENTRO
