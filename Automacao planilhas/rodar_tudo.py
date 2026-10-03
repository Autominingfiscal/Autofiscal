"""Roda a cadeia inteira, na ordem:

    EXPEDICAO -> REMESSA -> (PDFs -> chave de acesso) -> FATURAMENTO

Se uma etapa parar, as seguintes nao rodam - e o que voce quer, porque cada
uma alimenta a proxima.

Rode com um duplo clique em  0 - rodar tudo.bat
ou pelo prompt:  python rodar_tudo.py
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

from automining import chaves, config, faturamento, remessas
from automining.planilha import ErroDeUso
from automining.relatorio import cabecalho, erro, fim, linha


def main():
    cabecalho("CADEIA COMPLETA")
    cfg = config.carregar()

    print(">> 1/3  REMESSAS PORTO")
    r1 = remessas.atualizar(cfg)
    print(f"   {r1['notas']} nota(s) lancada(s), {r1['sem_chave']} sem chave")
    linha()

    print(">> 2/3  CHAVES DE ACESSO")
    try:
        r2 = chaves.preencher(cfg)
        print(f"   {r2['preenchidas']} preenchida(s), "
              f"{r2['sem_pdf'] + r2['nao_validaram']} pendente(s)")
    except ErroDeUso as e:
        r2 = None
        print(f"   pulada: {e}")
    linha()

    print(">> 3/3  FATURAMENTO")
    r3 = faturamento.atualizar(cfg)
    print(f"   {r3['completadas']} completada(s), {r3['inseridas']} inserida(s), "
          f"{r3['ja_lancadas']} ja lancada(s)")
    if r3["valor_atualizado"]:
        print(f"   {r3['valor_atualizado']} com valor atualizado (VL. UNITARIO mudou na remessa)")

    linha()
    if r3["avisos"]:
        print("ATENCAO:")
        for aviso in r3["avisos"]:
            print(f"  {aviso}")
        linha()
    print("Copias de seguranca na subpasta Backup de cada planilha.")
    fim("Confira as planilhas.")


if __name__ == "__main__":
    try:
        main()
    except ErroDeUso as e:
        erro(str(e))
        sys.exit(1)
    except Exception as e:  # noqa: BLE001
        erro(f"A rotina parou com erro inesperado.\n\n{type(e).__name__}: {e}")
        raise
