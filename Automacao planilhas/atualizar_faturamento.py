"""Lanca as remessas na aba do mes do faturamento.

Equivale a macro AtualizarFaturamento.
Rode com um duplo clique em  3 - atualizar faturamento.bat
ou pelo prompt:  python atualizar_faturamento.py
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

from automining.arranque import checar

checar()

from automining import config, faturamento
from automining.planilha import ErroDeUso
from automining.relatorio import cabecalho, erro, fim, linha


def main():
    cabecalho("FATURAMENTO")
    cfg = config.carregar()
    r = faturamento.atualizar(cfg)

    linha()
    print(f"Notas na remessa ..... {r['notas']}")
    print(f"Linhas completadas ... {r['completadas']}")
    print(f"Linhas inseridas ..... {r['inseridas']}")
    print(f"Ja estavam lancadas .. {r['ja_lancadas']}")
    if r["valor_atualizado"]:
        print(f"  com valor atualizado {r['valor_atualizado']} (VL. UNITARIO mudou na remessa)")
    if r["sem_aba"]:
        print(f"Sem aba do mes ....... {r['sem_aba']}")

    if r["por_mes"]:
        linha()
        print("Por mes:")
        for item in r["por_mes"]:
            print(f"  {item}")

    if r["avisos"]:
        linha()
        print("ATENCAO:")
        for aviso in r["avisos"]:
            print(f"  {aviso}")

    linha()
    print(f"Copia de seguranca: Backup\\{r['backup'].name}")
    fim("Confira a planilha.")


if __name__ == "__main__":
    try:
        main()
    except ErroDeUso as e:
        erro(str(e))
        sys.exit(1)
    except Exception as e:  # noqa: BLE001
        erro(f"A rotina parou com erro inesperado.\n\n{type(e).__name__}: {e}\n\n"
             "Nada foi salvo alem da copia na pasta Backup.")
        raise
