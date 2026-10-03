"""Caminhos do painel e o que ele lembra de cada usuario.

caminho_do_usuario e existe moram em comum/caminhos.py e sao repassados daqui.
"""

import json
import os

from comum.caminhos import caminho_do_usuario, existe  # noqa: F401  (usados pelo app)


# ------------------------------------------------------------------ estado ----
# O que o painel "lembra" (ultima pasta escolhida etc.) fica na pasta de dados
# de CADA usuario do Windows, para um nao atrapalhar o outro.
def _arquivo_estado():
    raiz = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    pasta = os.path.join(raiz, "Autofiscal")
    os.makedirs(pasta, exist_ok=True)
    return os.path.join(pasta, "estado.json")


def ler_estado():
    try:
        with open(_arquivo_estado(), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def gravar_estado(estado):
    try:
        with open(_arquivo_estado(), "w", encoding="utf-8") as f:
            json.dump(estado, f, ensure_ascii=False, indent=2)
    except OSError:
        pass
