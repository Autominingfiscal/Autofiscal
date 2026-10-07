# -*- coding: utf-8 -*-
"""DASHBOARD FLUIG - indicadores das solicitacoes de Entrada de Notas Fiscais.

Le a exportacao "Resultado da consulta de solicitacoes" do Fluig (.xlsx) e gera
um dashboard em HTML (abre no navegador, funciona sem internet) com:
  - total do ano, abertas no mes, em aberto agora, finalizadas, reprovadas,
    valor total das notas e SLA das tratativas (so solicitacoes finalizadas);
  - ranking de finalizadas por tecnico fiscal e ranking de fornecedores;
  - quantidade por tipo (servico, produtos gerais, energia ONS, CTE Rodogranel);
  - abertas x finalizadas por mes e distribuicao do SLA.

Os dados ficam so neste computador. As regras (o que e finalizada, reprovada,
cada tipo) ficam no dashboard_fluig_config.ini. O DIAGNOSTICO no fim mostra so
os valores das colunas de categoria (situacao, aprovacao, tipo), nunca
fornecedor, CNPJ ou valor: da para mandar para ajustar as regras.

    python dashboard_fluig.py                    planilha do .ini (a mais recente)
    python dashboard_fluig.py "<arquivo.xlsx>"   outra planilha
    python dashboard_fluig.py --diagnostico      so o diagnostico, sem gerar

Codigos de saida: 0 tudo certo | 1 erro
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
import glob
import json
import re
from collections import Counter

from comum.arquivos import ler_ini
from comum.caminhos import caminho_do_usuario
from comum.numeros import numero_br
from comum.texto import norm, sem_acento

PASTA = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PASTA)
from modelo_dashboard import MODELO  # noqa: E402

CONFIG = os.path.join(PASTA, "dashboard_fluig_config.ini")
SAIDA = os.path.join(PASTA, "Dashboard gerado")

# campo do dashboard -> cabecalho padrao da exportacao do Fluig
COLUNAS_PADRAO = {
    "solicitacao": "Solicitação",
    "situacao": "Situação da Solicitação",
    "inicio": "Início",
    "fim": "Fim",
    "tecnico": "Atividade - EntradadenotasFiscais - Fiscal - Responsável",
    "aprovacao": "aprovacao - EntradadenotasFiscais",
    "fornecedor": "razao_social_fornecedor - EntradadenotasFiscais",
    "cnpj": "cnpj_fornecedor - EntradadenotasFiscais",
    "valor": "valor_nota - EntradadenotasFiscais",
    "tipo_nota": "tipo_nota - EntradadenotasFiscais",
    "tipo_documento": "tipo_documento_nota - EntradadenotasFiscais",
    "tipo_contrato": "tipoContrato - EntradadenotasFiscais",
}
# so estas colunas aparecem no diagnostico (nenhuma tem fornecedor, CNPJ ou valor)
DO_DIAGNOSTICO = ("situacao", "aprovacao", "tipo_nota", "tipo_documento", "tipo_contrato")


# ----------------------------------------------------------------- config ----
def _lista(texto):
    return [norm(p) for p in re.split(r"[;,]", texto or "") if norm(p)]


def ler_config(caminho=CONFIG):
    cp = ler_ini(caminho, inline_comment_prefixes=(";",))
    colunas = dict(COLUNAS_PADRAO)
    if cp.has_section("colunas"):
        for campo in colunas:
            if cp.get("colunas", campo, fallback="").strip():
                colunas[campo] = cp.get("colunas", campo).strip()
    sit = cp["situacao"] if cp.has_section("situacao") else {}
    # [categorias]  1 = Serviço | SERVI   (nome no dashboard | palavras), na ordem do numero
    categorias = []
    if cp.has_section("categorias"):
        regras = sorted(cp.items("categorias"), key=lambda kv: (len(kv[0]), kv[0]))
        for _, valor in regras:
            nome, _, palavras = valor.partition("|")
            if nome.strip() and _lista(palavras):
                categorias.append((nome.strip(), _lista(palavras)))
    return {
        "planilha": cp.get("geral", "planilha", fallback=""),
        "colunas": colunas,
        "finalizada": _lista(sit.get("finalizada", "FINALIZ, CONCLU")),
        "cancelada": _lista(sit.get("cancelada", "CANCEL")),
        "aberta": _lista(sit.get("aberta", "ABERT, ANDAMENTO, PENDENT")),
        "reprovada": _lista(cp.get("aprovacao", "reprovada", fallback="REPROV")),
        "reprovada_igual": _lista(cp.get("aprovacao", "reprovada_se_igual", fallback="NAO, N, FALSE, 0")),
        "categorias": categorias,
    }


def achar_planilha(padrao):
    """O arquivo mais recente que bate com o padrao (aceita *)."""
    caminho = caminho_do_usuario(padrao)
    achados = [c for c in glob.glob(caminho) if not os.path.basename(c).startswith("~$")]
    return max(achados, key=os.path.getmtime) if achados else None


# ------------------------------------------------------------------ leitura ----
def ler_linhas(caminho):
    """(cabecalho, linhas) da 1a aba. Tenta o modo rapido; se vier vazio, o normal."""
    import openpyxl
    for rapido in (True, False):
        wb = openpyxl.load_workbook(caminho, read_only=rapido, data_only=True)
        try:
            ws = wb.worksheets[0]
            it = ws.iter_rows(values_only=True)
            cab = [str(c).strip() if c is not None else "" for c in next(it, [])]
            linhas = [r for r in it if any(v not in (None, "") for v in r)]
        finally:
            wb.close()
        if linhas or not rapido:
            return cab, linhas
    return [], []


def para_data(v):
    """datetime de uma celula: data do Excel ou texto 'dd/mm/aaaa hh:mm(:ss)' / 'aaaa-mm-dd'."""
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    if not isinstance(v, str) or not v.strip():
        return None
    s = v.strip()
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})(?:\D+(\d{1,2}):(\d{2})(?::(\d{2}))?)?", s)
    if m:
        d, mes, a, h, mi, se = m.groups()
    else:
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?", s)
        if not m:
            return None
        a, mes, d, h, mi, se = m.groups()
    try:
        return dt.datetime(int(a), int(mes), int(d), int(h or 0), int(mi or 0), int(se or 0))
    except ValueError:
        return None


def para_valor(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    return numero_br(v) if isinstance(v, str) else None


def texto(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return " ".join(str(v).split())


def situacao_de(valor, cfg):
    n = norm(valor)
    for codigo, chave in (("C", "cancelada"), ("F", "finalizada"), ("A", "aberta")):
        if any(p in n for p in cfg[chave]):
            return codigo
    return "?"


def reprovada(valor, cfg):
    n = norm(valor)
    return bool(n) and (any(p in n for p in cfg["reprovada"]) or n in cfg["reprovada_igual"])


def palavras_de(textos):
    """'Serviço CT-e' -> ['SERVICO', 'CTE'] (sem acento; hifen e ponto nao separam)."""
    junto = sem_acento(" ".join(textos)).upper().replace("-", "").replace(".", "")
    return re.findall(r"[A-Z0-9]+", junto)


def categoria_de(textos, cfg):
    """A 1a regra com uma palavra que COMECE com uma das palavras dela. Pelo comeco
    da palavra, e nao por pedaco: 'ONS' nao pode bater em 'CONSTRUCAO'."""
    palavras = palavras_de(textos)
    for nome, chaves in cfg["categorias"]:
        if any(p.startswith(c) for c in chaves for p in palavras):
            return nome
    return "Outros"


def montar_registros(cab, linhas, cfg):
    """Uma solicitacao por registro. Linha repetida da mesma solicitacao completa
    os campos vazios da anterior. Devolve (registros, avisos, diagnostico)."""
    idx = {norm(c): j for j, c in enumerate(cab)}
    col = {campo: idx.get(norm(nome)) for campo, nome in cfg["colunas"].items()}
    faltando = [cfg["colunas"][c] for c in ("solicitacao", "situacao", "inicio") if col[c] is None]
    if faltando:
        raise ValueError("a planilha nao tem a(s) coluna(s): " + "; ".join(faltando))
    pega = lambda lin, campo: lin[col[campo]] if col[campo] is not None and col[campo] < len(lin) else None

    por_numero, ordem = {}, []
    for lin in linhas:
        num = texto(pega(lin, "solicitacao"))
        if not num:
            continue
        atual = {campo: pega(lin, campo) for campo in col}
        if num in por_numero:
            antigo = por_numero[num]
            for k, v in atual.items():
                if v not in (None, ""):
                    antigo[k] = v
        else:
            por_numero[num] = atual
            ordem.append(num)

    diag = {c: Counter() for c in DO_DIAGNOSTICO}
    registros, sem_data = [], 0
    for num in ordem:
        r = por_numero[num]
        for c in DO_DIAGNOSTICO:
            diag[c][texto(r[c]) or "(vazio)"] += 1
        inicio, fim = para_data(r["inicio"]), para_data(r["fim"])
        if inicio is None:
            sem_data += 1
            continue
        sit = situacao_de(r["situacao"], cfg)
        sla = None
        if sit == "F" and fim and fim >= inicio:
            sla = round((fim - inicio).total_seconds() / 3600, 2)
        fornecedor = texto(r["fornecedor"]) or texto(r["cnpj"]) or "(sem fornecedor)"
        registros.append({
            "n": num,
            "i": inicio.strftime("%Y-%m-%d"),
            "f": fim.strftime("%Y-%m-%d") if fim and sit == "F" else None,
            "s": sit,
            "t": texto(r["tecnico"]) or "(sem técnico)",
            "r": 1 if reprovada(r["aprovacao"], cfg) else 0,
            "fo": fornecedor,
            "v": para_valor(r["valor"]),
            "k": categoria_de([texto(r[c]) for c in ("tipo_nota", "tipo_documento", "tipo_contrato")], cfg),
            "h": sla,
        })
    avisos = []
    repetidas = sum(1 for lin in linhas if texto(pega(lin, "solicitacao"))) - len(ordem)
    if repetidas:
        avisos.append(f"{repetidas} linha(s) repetiam uma solicitacao e foram juntadas a ela")
    if sem_data:
        avisos.append(f"{sem_data} solicitacao(oes) sem data de inicio ficaram de fora")
    sem_col = [cfg["colunas"][c] for c, j in col.items() if j is None]
    if sem_col:
        avisos.append("colunas nao encontradas (o indicador delas fica vazio): " + "; ".join(sem_col))
    return registros, avisos, diag


# ------------------------------------------------------------------- saida ----
def compactar(registros):
    """Tecnicos e fornecedores viram indices de uma lista: o HTML fica bem menor."""
    tecnicos = sorted({r["t"] for r in registros})
    fornecedores = sorted({r["fo"] for r in registros})
    categorias = sorted({r["k"] for r in registros}, key=lambda k: (k == "Outros", k))
    it, ifo, ik = ({v: j for j, v in enumerate(l)} for l in (tecnicos, fornecedores, categorias))
    linhas = [[r["n"], r["i"], r["f"], r["s"], it[r["t"]], r["r"], ifo[r["fo"]], r["v"], ik[r["k"]], r["h"]]
              for r in registros]
    return {"tecnicos": tecnicos, "fornecedores": fornecedores, "categorias": categorias,
            "linhas": linhas}


def gerar_html(registros, origem, cfg, agora=None):
    agora = agora or dt.datetime.now()
    dados = compactar(registros)
    dados["gerado"] = agora.strftime("%d/%m/%Y %H:%M")
    dados["origem"] = os.path.basename(origem)
    dados["hoje"] = agora.strftime("%Y-%m-%d")
    dados["ordem_categorias"] = [n for n, _ in cfg["categorias"]] + ["Outros"]
    js = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    return MODELO.replace("/*__DADOS__*/null", js)


def mostrar_diagnostico(diag, registros, cfg):
    print("\nDIAGNOSTICO (so colunas de categoria: pode ser enviado para ajustar as regras)")
    nomes = {"situacao": "Situação da Solicitação", "aprovacao": "aprovacao",
             "tipo_nota": "tipo_nota", "tipo_documento": "tipo_documento_nota",
             "tipo_contrato": "tipoContrato"}
    for c in DO_DIAGNOSTICO:
        print(f"\n  {nomes[c]}:")
        for valor, n in diag[c].most_common(15):
            extra = ""
            if c == "situacao" and valor != "(vazio)":
                extra = {"F": "finalizada", "A": "aberta", "C": "cancelada"}.get(
                    situacao_de(valor, cfg), "NAO RECONHECIDA - ajuste [situacao] no .ini")
                extra = f"  -> {extra}"
            elif c == "aprovacao" and valor != "(vazio)":
                extra = "  -> reprovada" if reprovada(valor, cfg) else ""
            print(f"    {n:>7}  {valor[:60]}{extra}")
        if len(diag[c]) > 15:
            print(f"    ... e mais {len(diag[c]) - 15} valor(es)")
    tipos = Counter(r["k"] for r in registros)
    print("\n  Tipo de cada solicitacao (pelas regras de [categorias]):")
    for nome in [n for n, _ in cfg["categorias"]] + ["Outros"]:
        print(f"    {tipos.get(nome, 0):>7}  {nome}")
    if tipos.get("Outros"):
        print("  ATENÇÃO: ha solicitacoes em 'Outros'. Veja os valores de tipo acima e")
        print("  acrescente as palavras deles em [categorias] no .ini.")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Dashboard Fluig")
    ap.add_argument("planilha", nargs="?", help="exportacao do Fluig (.xlsx). Padrao: a do .ini")
    ap.add_argument("--diagnostico", action="store_true", help="so mostra o diagnostico")
    ap.add_argument("--nao-abrir", action="store_true", help="nao abre o dashboard no navegador")
    args = ap.parse_args(argv)

    try:
        cfg = ler_config()
    except Exception as e:
        print(f"ERRO no {os.path.basename(CONFIG)}: {e}")
        return 1
    caminho = caminho_do_usuario(args.planilha) if args.planilha else achar_planilha(cfg["planilha"])
    if not caminho or not os.path.isfile(caminho):
        print(f"ERRO: nao achei a planilha ({args.planilha or cfg['planilha']}).")
        print("Exporte a consulta no Fluig ou ajuste 'planilha' em Configuracoes.")
        return 1
    modificada = dt.datetime.fromtimestamp(os.path.getmtime(caminho))
    print(f"Planilha: {caminho}")
    print(f"Salva em {modificada:%d/%m/%Y %H:%M}. Lendo...", flush=True)
    try:
        cab, linhas = ler_linhas(caminho)
        registros, avisos, diag = montar_registros(cab, linhas, cfg)
    except PermissionError:
        print("ERRO: a planilha esta aberta em outro programa. Feche e rode de novo.")
        return 1
    except ValueError as e:
        print(f"ERRO: {e}")
        return 1
    print(f"{len(registros)} solicitacao(oes) lida(s) de {len(linhas)} linha(s).")
    for a in avisos:
        print(f"ATENÇÃO: {a}")
    mostrar_diagnostico(diag, registros, cfg)
    if args.diagnostico:
        return 0
    if not registros:
        print("\nERRO: nenhuma solicitacao com data de inicio: nada para mostrar.")
        return 1

    os.makedirs(SAIDA, exist_ok=True)
    destino = os.path.join(SAIDA, "Dashboard Fluig.html")
    with open(destino, "w", encoding="utf-8") as f:
        f.write(gerar_html(registros, caminho, cfg))
    print(f"\nDashboard: {destino}")
    if not args.nao_abrir:
        try:
            os.startfile(destino)
        except (AttributeError, OSError):
            pass
    return 0


if __name__ == "__main__":
    codigo = main()
    if sys.stdin is not None and sys.stdin.isatty():
        try:
            input("Pressione Enter para fechar...")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(codigo)
