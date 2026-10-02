# -*- coding: utf-8 -*-
"""Exemplo minimo de ferramenta para o painel Autofiscal."""

import os
import sys


def main():
    pasta = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(f"Pasta: {pasta}")
    arquivos = os.listdir(pasta)
    print(f"{len(arquivos)} item(ns) na pasta.")
    nome = input("Como voce se chama? ")
    print(f"Ola, {nome}! Terminei.")


if __name__ == "__main__":
    main()
    if sys.stdin is not None and sys.stdin.isatty():
        input("Pressione Enter para fechar...")
