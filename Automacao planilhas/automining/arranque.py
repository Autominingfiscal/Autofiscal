"""Conferencias antes de comecar: versao do Python e openpyxl instalado."""

import sys

MINIMO = (3, 8)


def checar():
    if sys.version_info < MINIMO:
        _parar(
            f"Este Python e a versao {sys.version_info.major}."
            f"{sys.version_info.minor}, e a rotina precisa da "
            f"{MINIMO[0]}.{MINIMO[1]} ou mais nova."
        )
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        _parar(
            "Falta a biblioteca openpyxl, que e quem abre os arquivos .xlsx.\n\n"
            "Instale abrindo o Prompt de Comando e rodando:\n\n"
            "    pip install openpyxl\n\n"
            "Se o pip nao for reconhecido, tente:\n\n"
            "    python -m pip install openpyxl\n\n"
            "Se a empresa bloquear o download, peca ao TI para instalar o "
            "pacote openpyxl. Nenhuma outra biblioteca e necessaria."
        )


def _parar(mensagem):
    print()
    print("!" * 70)
    print("  NAO DEU PARA COMECAR")
    print("!" * 70)
    print()
    print(mensagem)
    print()
    if sys.stdin is not None and sys.stdin.isatty():
        try:
            input("Pressione Enter para fechar...")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(1)
