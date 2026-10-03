# -*- coding: utf-8 -*-
"""Exemplo minimo de ferramenta para o painel Autofiscal.

Ponto de partida para uma ferramenta nova (Manutencao > Nova ferramenta copia
esta pasta). Mostra os tres combinados de toda ferramenta:
  1. o bloco "comum" logo abaixo, para usar o codigo compartilhado;
  2. print() para escrever no registro do painel;
  3. pausa no fim so quando aberto fora do painel.
"""

# --- pasta "comum" do Autofiscal ---------------------------------------------
# Fica na pasta Autofiscal, logo acima desta ferramenta, ou dentro dela quando a
# ferramenta foi exportada para outro PC (Manutencao > Exportar ferramenta).
import os
import sys

for _pasta in (os.path.dirname(os.path.abspath(__file__)),
               os.path.dirname(os.path.dirname(os.path.abspath(__file__)))):
    if os.path.isdir(os.path.join(_pasta, "comum")):
        sys.path.insert(0, _pasta)
        break
else:
    sys.exit("Nao achei a pasta 'comum' do Autofiscal, nem nesta pasta nem na de cima.\n"
             "Para usar a ferramenta fora da pasta Autofiscal, copie-a pela ferramenta\n"
             "Manutencao > Exportar ferramenta, que leva a pasta 'comum' junto.")
# -----------------------------------------------------------------------------

from comum.caminhos import caminho_do_usuario
from comum.numeros import numero_br


def main():
    pasta = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    pasta = caminho_do_usuario(pasta)
    print(f"Pasta: {pasta}")
    arquivos = os.listdir(pasta)
    print(f"{len(arquivos)} item(ns) na pasta.")
    valor = numero_br(input("Digite um valor (ex.: R$ 1.234,56): "))
    if valor is None:
        print("ATENÇÃO: nao entendi o valor.")
    else:
        print(f"Valor lido: {valor:.2f}")
    print("Terminei.")


if __name__ == "__main__":
    main()
    if sys.stdin is not None and sys.stdin.isatty():
        try:
            input("Pressione Enter para fechar...")
        except (EOFError, KeyboardInterrupt):
            pass
