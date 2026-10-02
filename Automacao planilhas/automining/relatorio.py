"""Relatorio em tela, no lugar do MsgBox do VBA."""

import sys

LARGURA = 70


def _escrever(texto=""):
    try:
        print(texto)
    except UnicodeEncodeError:
        print(texto.encode("ascii", "replace").decode("ascii"))


def cabecalho(titulo):
    _escrever()
    _escrever("=" * LARGURA)
    _escrever(f"  AUTOMACAO MVV  -  {titulo}")
    _escrever("=" * LARGURA)
    _escrever()


def linha():
    _escrever("-" * LARGURA)


def fim(recado=""):
    linha()
    if recado:
        _escrever(recado)
    _escrever()
    _pausar()


def erro(mensagem):
    _escrever()
    _escrever("!" * LARGURA)
    _escrever("  NAO DEU CERTO")
    _escrever("!" * LARGURA)
    _escrever()
    _escrever(mensagem)
    _escrever()
    _pausar()


def _pausar():
    """Segura a janela aberta quando o script foi chamado com duplo clique."""
    if sys.stdin is not None and sys.stdin.isatty():
        try:
            input("Pressione Enter para fechar...")
        except (EOFError, KeyboardInterrupt):
            pass
