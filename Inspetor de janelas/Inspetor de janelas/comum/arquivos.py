"""Leitura de arquivos de texto e .ini editados a mao no Bloco de Notas.

O Bloco de Notas pode salvar em UTF-8 (com ou sem BOM) ou em ANSI (cp1252),
dependendo da versao do Windows e de quem salvou. Aqui os dois funcionam.
"""

import configparser


def ler_texto(caminho):
    """Conteudo do arquivo: UTF-8 (com ou sem BOM) ou, se nao for, ANSI."""
    with open(caminho, "rb") as f:
        bruto = f.read()
    try:
        return bruto.decode("utf-8-sig")
    except UnicodeDecodeError:
        return bruto.decode("cp1252")


def ler_ini(caminho, **opcoes):
    """ConfigParser ja lido. Sem interpolacao: '%' no caminho nao quebra nada.

    `opcoes` vai para o ConfigParser (ex.: inline_comment_prefixes=(";",))."""
    opcoes.setdefault("interpolation", None)
    cp = configparser.ConfigParser(**opcoes)
    cp.read_string(ler_texto(caminho), source=str(caminho))
    return cp
