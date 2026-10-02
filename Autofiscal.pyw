# -*- coding: utf-8 -*-
"""AUTOFISCAL - painel com todas as ferramentas do trabalho.

Dois cliques neste arquivo (ou em "Abrir Autofiscal.bat") abrem a janela.
Cada subpasta com um arquivo ferramenta.json aparece como uma ferramenta.
"""

import os
import sys
import traceback

RAIZ = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, RAIZ)

try:
    from painel.app import iniciar
    iniciar(RAIZ)
except Exception:
    erro = traceback.format_exc()
    try:
        with open(os.path.join(RAIZ, "autofiscal_erro.txt"), "w", encoding="utf-8") as f:
            f.write(erro)
    except OSError:
        pass
    try:
        from tkinter import Tk, messagebox
        r = Tk()
        r.withdraw()
        messagebox.showerror("Autofiscal", "O painel nao abriu por causa de um erro:\n\n"
                             + erro[-1200:] + "\n\n(salvo em autofiscal_erro.txt)")
    except Exception:
        print(erro)
        input("Enter para fechar...")
