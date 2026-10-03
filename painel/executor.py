"""Roda o script de uma ferramenta quando o Autofiscal esta instalado como .exe.

Instalado, nao existe python.exe no computador: o painel chama o
"Autofiscal Executor.exe" (que e este mesmo programa, com console) assim:

    "Autofiscal Executor.exe" --rodar-script <script.py> [args...]

e ele roda o script como o python.exe rodaria. As bibliotecas (openpyxl,
pypdf, pillow, pywin32) vao junto dentro do instalador.
"""

import os
import runpy
import sys

OPCAO = "--rodar-script"
EXECUTOR = "Autofiscal Executor.exe"


def instalado():
    """True quando roda de dentro do .exe gerado pelo PyInstaller."""
    return bool(getattr(sys, "frozen", False))


def pasta_do_programa():
    """Pasta com as ferramentas: a do .exe, quando instalado."""
    if instalado():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def caminho_executor():
    return os.path.join(pasta_do_programa(), EXECUTOR)


def pedido_de_script(argv):
    """O script e os argumentos, se o .exe foi chamado para rodar um script."""
    if len(argv) >= 3 and argv[1] == OPCAO:
        return argv[2], list(argv[3:])
    return None


def rodar_script(script, args):
    """Faz o que `python -u script args` faria. Devolve o codigo de saida."""
    script = os.path.abspath(script)
    for fluxo in (sys.stdout, sys.stderr):
        try:
            fluxo.reconfigure(encoding="utf-8", errors="replace", write_through=True)
        except (AttributeError, ValueError):
            pass
    try:
        sys.stdin.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass
    sys.argv = [script] + list(args)
    sys.path.insert(0, os.path.dirname(script))
    try:
        runpy.run_path(script, run_name="__main__")
    except SystemExit as e:
        if e.code is None or isinstance(e.code, int):
            return e.code or 0
        print(e.code, file=sys.stderr)
        return 1
    return 0
