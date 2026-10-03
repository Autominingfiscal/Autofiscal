"""Lanca as notas da expedicao na aba Remessas Porto.

Equivale a macro AtualizarRemessasPorto.
Rode com um duplo clique em  1 - atualizar remessas.bat
ou pelo prompt:  python atualizar_remessas.py
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

from automining import config, remessas
from automining.planilha import ErroDeUso
from automining.relatorio import cabecalho, erro, fim, linha


def main():
    cabecalho("REMESSAS PORTO")
    cfg = config.carregar()
    r = remessas.atualizar(cfg)

    linha()
    print(f"Notas lancadas ............... {r['notas']}")
    print(f"Dias ......................... {r['dias']}")
    print(f"Chaves de acesso preenchidas . {r['com_chave']}")
    print(f"Chaves de acesso em branco ... {r['sem_chave']}")
    if r["linhas_inseridas"]:
        print(f"Linhas criadas na area ....... {r['linhas_inseridas']}")
    if r["unit_padrao"]:
        linha()
        padrao = f"{r['valor_padrao']:.2f}".replace(".", ",")
        print(f"Notas novas com o VL. UNITARIO padrao ({padrao}) - confira e corrija na planilha:")
        print("  NF " + ", ".join(str(int(nf)) for nf in r["unit_padrao"][:30]))
        if len(r["unit_padrao"]) > 30:
            print(f"  ... e mais {len(r['unit_padrao']) - 30}")
    if r["sem_peso"]:
        linha()
        print("Notas sem peso liquido (nao lancadas):")
        for aba, lin, nf in r["sem_peso"][:20]:
            print(f"  NF {nf} - aba {aba}, linha {lin}")
        if len(r["sem_peso"]) > 20:
            print(f"  ... e mais {len(r['sem_peso']) - 20}")
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
