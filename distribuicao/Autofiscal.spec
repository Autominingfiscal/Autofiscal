# -*- mode: python ; coding: utf-8 -*-
# Receita do PyInstaller. Nao rode direto: use gerar_instalador.py, que passa
# os valores abaixo pelas variaveis de ambiente AUTOFISCAL_*.
#
# Gera dois .exe que dividem a mesma pasta _internal:
#   Autofiscal.exe            a janela (sem console)
#   Autofiscal Executor.exe   roda os scripts das ferramentas (com console,
#                             que o painel esconde) - veja painel/executor.py

import os

raiz = os.environ["AUTOFISCAL_RAIZ"]
ocultos = [m for m in os.environ.get("AUTOFISCAL_OCULTOS", "").split(os.pathsep) if m]
icone = os.environ.get("AUTOFISCAL_ICONE")
info_versao = os.environ.get("AUTOFISCAL_INFO_VERSAO")

a = Analysis(
    [os.path.join(raiz, "Autofiscal.pyw")],
    pathex=[raiz],
    hiddenimports=ocultos + ["comum", "painel.app"],
    datas=[(icone, "painel")],                 # icone da janela (painel/app.py)
    excludes=["pytest", "IPython", "matplotlib", "numpy", "pandas"],
    noarchive=False,
)
pyz = PYZ(a.pure)

comum = dict(exclude_binaries=True, icon=icone, version=info_versao, upx=False)
janela = EXE(pyz, a.scripts, [], name="Autofiscal", console=False, **comum)
executor = EXE(pyz, a.scripts, [], name="Autofiscal Executor", console=True, **comum)

COLLECT(janela, executor, a.binaries, a.datas, name="Autofiscal", upx=False)
