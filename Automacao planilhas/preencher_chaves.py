"""Le os PDFs da rede e preenche a coluna CHAVE ACESSO.

Equivale a macro PreencherChavesAcesso.
Rode com um duplo clique em  2 - preencher chaves.bat
ou pelo prompt:  python preencher_chaves.py

Para testar um PDF isolado:  python preencher_chaves.py "caminho\\do\\arquivo.pdf" 19007
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

from pathlib import Path

from automining.arranque import checar

checar()

from automining import chaves, config
from automining.planilha import ErroDeUso
from automining.relatorio import cabecalho, erro, fim, linha


def conferir_um(caminho, nf):
    cabecalho("CONFERIR UM PDF")
    print(f"Arquivo: {caminho}")
    r = chaves.conferir_um(Path(caminho), nf)
    linha()
    if r["ok"]:
        print(f"Chave lida ......... {r['chave']}")
        print(f"Nota dentro da chave {r['nota_na_chave']}")
        if nf:
            print(f"Nota esperada ...... {int(nf)}")
    else:
        print(f"Nao validou: {r['motivo']}")
    fim()


def main():
    cabecalho("CHAVES DE ACESSO")
    cfg = config.carregar()
    r = chaves.preencher(cfg)

    linha()
    print(f"Preenchidas agora ....... {r['preenchidas']}")
    print(f"Ja estavam preenchidas .. {r['ja_tinha']}")
    print(f"Sem PDF encontrado ...... {r['sem_pdf']}")
    print(f"Nao validaram ........... {r['nao_validaram']}")
    linha()
    print(f"PDFs indexados: {r['pdfs']} em {r['pastas']} pasta(s)")
    print("Subpastas: " + ("incluidas" if r["subpastas"] else "ignoradas (config.ini)"))

    if r["pendencias"]:
        linha()
        print("Pendencias:")
        for p in r["pendencias"][:40]:
            print(f"  {p}")
        if len(r["pendencias"]) > 40:
            print(f"  ... e mais {len(r['pendencias']) - 40}")

    linha()
    print(f"Copia de seguranca: Backup\\{r['backup'].name}")
    fim("Confira a planilha.")


if __name__ == "__main__":
    try:
        if len(sys.argv) > 1:
            nf = float(sys.argv[2]) if len(sys.argv) > 2 else None
            conferir_um(sys.argv[1], nf)
        else:
            main()
    except ErroDeUso as e:
        erro(str(e))
        sys.exit(1)
    except Exception as e:  # noqa: BLE001
        erro(f"A rotina parou com erro inesperado.\n\n{type(e).__name__}: {e}")
        raise
