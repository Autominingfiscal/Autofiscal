"""Caminhos que funcionam para qualquer usuario do Windows."""

import glob
import json
import os
import re


def caminho_do_usuario(texto, base=None):
    """Ajusta um caminho para o usuario que esta usando o computador.

    - aceita variaveis do Windows: %USERPROFILE%, %OneDrive%, %USERNAME%, ~
    - caminho relativo vale a partir de `base`
    - caminho de OUTRO usuario (C:\\Users\\fulano\\...) que nao existe aqui e
      trocado pela pasta do usuario atual; a parte "OneDrive..." vira a
      OneDrive dele (mesmo que o nome seja "OneDrive - Empresa").
    """
    if not texto:
        return texto
    t = str(texto).strip().strip('"')
    t = re.sub(r"%([^%]+)%", lambda m: os.environ.get(m.group(1), m.group(0)), t)
    t = os.path.expanduser(t)
    if base and not os.path.isabs(t) and not t.startswith(("\\\\", "//")):
        t = os.path.normpath(os.path.join(base, t))

    if existe(t):
        return t
    m = re.match(r"^[A-Za-z]:[\\/]+Users[\\/]+[^\\/]+[\\/]*(.*)$", t, re.I)
    if not m:
        return t
    resto = m.group(1)
    candidatos = []
    partes = re.split(r"[\\/]+", resto, maxsplit=1)
    for var in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        od = os.environ.get(var)
        if od and partes[0].lower().startswith("onedrive"):
            candidatos.append(os.path.join(od, partes[1]) if len(partes) > 1 else od)
    candidatos.append(os.path.join(os.path.expanduser("~"), resto))
    for c in candidatos:
        if existe(c):
            return c
    return t


def existe(caminho):
    if not caminho:
        return False
    if any(x in caminho for x in "*?"):
        return bool(glob.glob(caminho))
    return os.path.exists(caminho)


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
