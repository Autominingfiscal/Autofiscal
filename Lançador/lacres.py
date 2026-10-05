# -*- coding: utf-8 -*-
"""LACRES DO TURNO - registra o par de lacres de cada motorista e imprime as etiquetas.

Com a planilha EXPED_CONCENTRADO ABERTA no Excel:
  1. le os motoristas da tabela do turno (matutino ou vespertino) da aba do dia;
  2. pede o par de lacres de quem ainda nao tem (ex.: "9999 1265");
  3. confere: dois numeros, nenhum deles usado em outra linha ou outro dia;
  4. grava "9999 - 1265" na coluna CODIGO LACRE da tabela do turno (dai o
     Lancador leva para a expedicao e para as OBSERVACOES ADICIONAIS);
  5. gera as etiquetas (15 por folha A4, com linha de corte) e abre no navegador.

Quem ja tem lacre nao e perguntado: para trocar, apague a celula no Excel e
rode de novo. Rodar de novo com todos preenchidos so reimprime as etiquetas.

    python lacres.py --turno vespertino
    python lacres.py --turno matutino --aba 02.10
    python lacres.py --turno vespertino --teste ARQ.xlsx   nao grava na planilha

Codigos de saida: 0 tudo certo | 1 erro | 2 ficou motorista sem lacre
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

import argparse
import datetime as dt
import html
import re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import etapa3_lancador as lan  # noqa: E402

from comum.texto import norm  # noqa: E402

PASTA_ETIQUETAS = os.path.join(lan.PASTA_SCRIPT, "Etiquetas")
LACRES_POR_MOTORISTA = 2
POR_FOLHA = 15
TURNOS = {"matutino": 1, "vespertino": 2}
COL_LACRE = "CÓDIGO LACRE"


# ----------------------------------------------------------------- lacres ----
def numeros(texto):
    """Os numeros de lacre de um texto: '9999 - 1265' -> ['9999', '1265']."""
    if texto is None:
        return []
    if isinstance(texto, float) and texto.is_integer():
        texto = int(texto)
    return re.findall(r"\d+", str(texto))


def texto_da_celula(par):
    return " - ".join(par)


def lacres_em_uso(planilha):
    """{numero: 'aba 02.10, PEDRO MELO SILVA'} de todas as abas de dia da planilha."""
    uso = {}
    for aba in planilha.abas():
        if not re.fullmatch(r"\d{2}\.\d{2}", aba):
            continue
        for tab in planilha.tabelas(aba):
            c_lacre, c_mot = tab.col(COL_LACRE), tab.col("MOTORISTA")
            if c_lacre is None:
                continue
            for lin in tab.linhas:
                for n in numeros(lin[c_lacre]):
                    quem = lin[c_mot] if c_mot is not None and lin[c_mot] else "sem nome"
                    uso.setdefault(n, f"aba {aba}, {quem}")
    return uso


def motoristas_do_turno(planilha, aba, n_turno):
    """(tabela do turno, [motoristas]) - cada motorista com a linha e os dados da etiqueta."""
    turnos = lan.tabelas_turno(planilha.tabelas(aba))
    if len(turnos) < n_turno:
        return None, []
    tab = turnos[n_turno - 1]
    val = lambda lin, nome: lin[tab.col(nome)] if tab.col(nome) is not None else None
    lista = []
    for i, lin in enumerate(tab.linhas):
        nome = val(lin, "MOTORISTA")
        if lan.vazio(nome):
            continue
        lista.append({
            "linha": tab.linha_excel(i),
            "motorista": str(nome).strip(),
            "cavalo": lan.formatar_placa(val(lin, "PLACA DO CAVALO") or ""),
            "reboque": lan.formatar_placa(val(lin, "PLACA DO REBOQUE") or ""),
            "modelo": "" if lan.vazio(val(lin, "MODELO")) else str(val(lin, "MODELO")).strip(),
            "transportadora": ("" if lan.vazio(val(lin, "TRANSPORTADORA"))
                               else str(val(lin, "TRANSPORTADORA")).strip()),
            "lacres": numeros(val(lin, COL_LACRE)),
        })
    return tab, lista


def pedir_par(m, em_uso, perguntar):
    """Pergunta ate vir um par valido. None se a resposta for vazia (pula) ou acabar a entrada."""
    rotulo = f"{m['motorista']}" + (f" ({m['cavalo']})" if m["cavalo"] else "")
    while True:
        try:
            resposta = perguntar(f"Lacres de {rotulo}: ")
        except EOFError:
            return None
        if not resposta.strip():
            return None
        par = numeros(resposta)
        if len(par) != LACRES_POR_MOTORISTA:
            print(f"  ATENÇÃO: digite {LACRES_POR_MOTORISTA} numeros (ex.: 9999 1265). "
                  f"Li {len(par)}: {' '.join(par) or 'nenhum'}.")
            continue
        if par[0] == par[1]:
            print("  ATENÇÃO: os dois lacres sao o mesmo numero. Digite de novo.")
            continue
        repetidos = [f"{n} (ja esta em {em_uso[n]})" for n in par if n in em_uso]
        if repetidos:
            print("  ATENÇÃO: lacre ja usado: " + "; ".join(repetidos) + ". Digite de novo.")
            continue
        return par


def registrar(planilha, aba, tab, motoristas, perguntar=input):
    """Pede e grava o lacre de quem nao tem. Devolve quantos gravou."""
    faltam = [m for m in motoristas if not m["lacres"]]
    if not faltam:
        print("Todos os motoristas do turno ja tem lacre.")
        return 0
    print(f"{len(faltam)} motorista(s) sem lacre. Digite o par de cada um "
          "(Enter vazio pula o motorista).\n")
    em_uso = lacres_em_uso(planilha)
    col = tab.col_ini + tab.col(COL_LACRE)
    exped = lan.tabela_expedicao(planilha.tabelas(aba))
    gravados = 0
    for k, m in enumerate(faltam, start=1):
        print(f"[{k}/{len(faltam)}]", end=" ")
        par = pedir_par(m, em_uso, perguntar)
        if par is None:
            print(f"  pulado: {m['motorista']} ficou sem lacre.")
            continue
        planilha.escrever(aba, m["linha"], col, [texto_da_celula(par)])
        m["lacres"] = par
        for n in par:
            em_uso[n] = f"aba {aba}, {m['motorista']}"
        gravados += 1
        _completar_expedicao(planilha, aba, exped, m, par)
        print(f"  ok: {m['motorista']} -> {' - '.join(par)}")
    return gravados


def _completar_expedicao(planilha, aba, exped, m, par):
    """Caminhao que ja passou pela balanca: a linha dele na expedicao ganha o lacre
    tambem (o Lancador so copia o lacre do turno na hora da tara)."""
    if exped is None or exped.col(COL_LACRE) is None or exped.col("PLACA DO CAVALO") is None:
        return
    placa = norm(m["cavalo"])
    if not placa:
        return
    for i, lin in enumerate(exped.linhas):
        if norm(lin[exped.col("PLACA DO CAVALO")]) == placa and lan.vazio(lin[exped.col(COL_LACRE)]):
            planilha.escrever(aba, exped.linha_excel(i), exped.col_ini + exped.col(COL_LACRE),
                              [texto_da_celula(par)])
            print(f"  (ja tinha passado pela balanca: lacre gravado tambem na linha "
                  f"{exped.linha_excel(i)} da expedicao)")


# -------------------------------------------------------------- etiquetas ----
CSS = """
@page { size: A4; margin: 8mm 10mm; }
* { box-sizing: border-box; }
body { margin: 0; font-family: "Segoe UI", Arial, sans-serif; color: #000; background: #e5e7eb; }
.barra { padding: 12px 16px; background: #1e293b; color: #fff; font-size: 14px; }
.barra button { font: inherit; padding: 6px 16px; margin-right: 12px; cursor: pointer; }
.folha { width: 210mm; height: 297mm; padding: 8mm 10mm; margin: 12px auto; background: #fff;
         box-shadow: 0 1px 4px rgba(0,0,0,.25); page-break-after: always; }
.folha:last-child { page-break-after: auto; }
.etiqueta { height: calc(281mm / 15); display: flex; align-items: center;
            border-bottom: 0.3mm dashed #888; padding: 0 3mm; }
.esq { flex: 1; min-width: 0; padding-right: 4mm; }
.nome { font-size: 11pt; font-weight: 700; text-transform: uppercase; white-space: nowrap;
        overflow: hidden; text-overflow: ellipsis; line-height: 1.2; }
.carro { font-size: 7.5pt; margin-top: 0.6mm; white-space: nowrap; }
.transp { font-size: 5.8pt; letter-spacing: .04em; text-transform: uppercase; margin-top: 0.4mm; }
.dir { width: 52mm; border-left: 0.4mm solid #000; text-align: center; padding-left: 3mm; }
.rot { font-size: 5.5pt; font-weight: 700; letter-spacing: .25em; }
.num { font-size: 15pt; font-weight: 800; line-height: 1.1; white-space: nowrap; }
.data { font-size: 5.8pt; margin-top: 0.3mm; }
@media print {
  body { background: #fff; }
  .barra { display: none; }
  .folha { margin: 0; box-shadow: none; width: auto; height: auto; padding: 0; }
  .etiqueta { height: calc(281mm / 15); }
}
"""


def _etiqueta(m, turno, data):
    e = html.escape
    carro = []
    if m["cavalo"]:
        carro.append(f"Cavalo <b>{e(m['cavalo'])}</b>")
    if m["reboque"]:
        carro.append(f"Reboque <b>{e(m['reboque'])}</b>")
    if m["modelo"]:
        carro.append(e(m["modelo"].upper()))
    return (f'<div class="etiqueta"><div class="esq">'
            f'<div class="nome">{e(m["motorista"])}</div>'
            f'<div class="carro">{" · ".join(carro)}</div>'
            f'<div class="transp">{e(m["transportadora"])}</div></div>'
            f'<div class="dir"><div class="rot">LACRES</div>'
            f'<div class="num">{" – ".join(e(n) for n in m["lacres"])}</div>'
            f'<div class="data">{e(turno.capitalize())} · {data:%d/%m/%Y}</div></div></div>')


def etiquetas_html(motoristas, turno, data):
    com_lacre = [m for m in motoristas if m["lacres"]]
    folhas = [com_lacre[i:i + POR_FOLHA] for i in range(0, len(com_lacre), POR_FOLHA)]
    corpo = "".join('<div class="folha">' + "".join(_etiqueta(m, turno, data) for m in f)
                    + "</div>" for f in folhas)
    return (f'<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">'
            f"<title>Etiquetas de lacres - {turno} {data:%d/%m/%Y}</title><style>{CSS}</style></head>"
            f'<body><div class="barra"><button onclick="window.print()">Imprimir</button>'
            f"{len(com_lacre)} etiqueta(s) em {len(folhas)} folha(s). Na impressão: papel A4, "
            f"escala 100%, margens padrão.</div>{corpo}</body></html>")


def salvar_etiquetas(motoristas, turno, data, abrir=True):
    os.makedirs(PASTA_ETIQUETAS, exist_ok=True)
    caminho = os.path.join(PASTA_ETIQUETAS, f"lacres_{data:%Y-%m-%d}_{turno}.html")
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(etiquetas_html(motoristas, turno, data))
    if abrir:
        try:
            os.startfile(caminho)                 # abre no navegador
        except (AttributeError, OSError):
            pass
    return caminho


# --------------------------------------------------------------- programa ----
def data_da_aba(aba, hoje):
    """'02.10' -> 02/10 deste ano; se cair muito no futuro, e do ano passado
    (aba 28.12 aberta em janeiro)."""
    dia, mes = (int(x) for x in aba.split("."))
    data = dt.date(hoje.year, mes, dia)
    return data.replace(year=hoje.year - 1) if (data - hoje).days > 31 else data


def main(argv=None, perguntar=input):
    ap = argparse.ArgumentParser(description="Lacres do turno: registra e imprime as etiquetas")
    ap.add_argument("--turno", choices=sorted(TURNOS), required=True)
    ap.add_argument("--aba", help="aba do dia (ex.: 02.10). Padrao: a de hoje")
    ap.add_argument("--teste", metavar="ARQUIVO_XLSX", help="usa o arquivo salvo e nao grava nada nele")
    ap.add_argument("--nao-abrir", action="store_true", help="nao abre as etiquetas no navegador")
    args = ap.parse_args(argv)

    hoje = dt.date.today()
    aba = args.aba or lan.nome_aba(hoje)
    if not re.fullmatch(r"\d{2}\.\d{2}", aba):
        print(f"ERRO: aba '{aba}' fora do formato dd.mm (ex.: 02.10).")
        return 1
    if args.teste:
        planilha = lan.PlanilhaTeste(args.teste)
    else:
        cfg = lan.carregar_config()
        planilha = lan.conectar_excel(cfg)
        if planilha is None:
            print(f"ERRO: abra a planilha {cfg['planilha']} no Excel e rode de novo.")
            return 1
    print(f"Planilha: {planilha.nome}   aba: {aba}   turno: {args.turno.upper()}\n")
    if aba not in planilha.abas():
        print(f"ERRO: a planilha nao tem a aba {aba}.")
        return 1
    tab, motoristas = motoristas_do_turno(planilha, aba, TURNOS[args.turno])
    if tab is None:
        print(f"ERRO: nao achei a tabela do turno {args.turno} (coluna PESO A CARREGAR) na aba {aba}.")
        return 1
    if tab.col(COL_LACRE) is None:
        print(f"ERRO: a tabela do turno nao tem a coluna {COL_LACRE}.")
        return 1
    if not motoristas:
        print(f"ERRO: a tabela do turno {args.turno} esta sem motoristas.")
        return 1

    gravados = registrar(planilha, aba, tab, motoristas, perguntar)
    if gravados and not args.teste:
        planilha.salvar()
        print(f"\n{gravados} lacre(s) gravado(s) e planilha salva.")

    sem = [m["motorista"] for m in motoristas if not m["lacres"]]
    if any(m["lacres"] for m in motoristas):
        caminho = salvar_etiquetas(motoristas, args.turno, data_da_aba(aba, hoje),
                                   abrir=not args.nao_abrir)
        print(f"\nEtiquetas: {caminho}")
        print("Abriram no navegador: clique em Imprimir (papel A4, escala 100%).")
    if sem:
        print(f"\nATENÇÃO: {len(sem)} motorista(s) sem lacre e sem etiqueta: {', '.join(sem)}")
        return 2
    return 0


if __name__ == "__main__":
    codigo = main()
    if sys.stdin is not None and sys.stdin.isatty():
        try:
            input("Pressione Enter para fechar...")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(codigo)
