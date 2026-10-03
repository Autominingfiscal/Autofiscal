# -*- coding: utf-8 -*-
"""
CONFERÊNCIA DE RETENÇÕES  -  PLANILHÃO  x  2-CONFERENCIA RETENÇÕES 2026

Compara, nota a nota, os impostos retidos do PLANILHÃO com a planilha
de conferência e gera um relatório em Excel dizendo o que está OK,
o que não bate e o que está faltando.

Como usar:
  1. Salve e feche as duas planilhas no Excel (o script lê o valor salvo).
  2. Coloque este arquivo na mesma pasta das planilhas.
  3. Dê dois cliques no "rodar_conferencia.bat" (ou rode: python conferencia_retencoes.py).
  4. Se ele não achar os arquivos sozinho, abre uma janela para você escolher.

Requer apenas Python 3 + openpyxl.
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

import re
import glob
from datetime import datetime
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import column_index_from_string, get_column_letter
except ImportError:
    print("ERRO: o openpyxl não está instalado. Rode:  pip install openpyxl")
    if sys.stdin is not None and sys.stdin.isatty():
        input("Enter para sair...")
    sys.exit(1)

from comum.numeros import numero_br


# =====================================================================
# CONFIGURAÇÃO  (se o layout das planilhas mudar, ajuste só aqui)
# =====================================================================

# Diferença máxima aceita (em R$) para considerar OK. Cobre arredondamento.
TOLERANCIA = 0.01

# ---- PLANILHÃO (1 linha por ITEM da nota -> o script soma por nota) ----
PLAN_LINHA_INICIAL = 2
PLAN_COL_NF = "A"
PLAN_COL_COD_FORN = "H"
PLAN_COL_CNPJ = "I"
PLAN_COL_RAZAO = "J"

# ---- CONFERÊNCIA ----
CONF_LINHA_INICIAL = 4          # dados começam na linha 4
CONF_COL_COD_FORN = "A"
CONF_COL_FORNECEDOR = "B"
CONF_COL_CNPJ = "C"
CONF_COL_TITULO = "F"
CONF_COL_PARADA = "H"           # para de ler quando achar "TOTAL" nesta coluna

# ---- IMPOSTOS: coluna do PLANILHÃO  x  colunas de VALOR da CONFERÊNCIA ----
# Obs.: nas faixas que você passou existem colunas de CÓDIGO, ALÍQUOTA e
# BASE (V, Z, AC, AD, AE, AI, AJ, M, N). Elas NÃO entram na soma, senão o
# código "1708" seria somado como se fosse dinheiro. Só as de valor entram.
IMPOSTOS = [
    {"nome": "IRRF", "plan": "AW", "cab_plan": "Valor IRRF", "conf": ["W", "X", "Y"],           # IRRF 1708 / 3280 / 3208
     "desc_conf": "IRRF 1708 + 3208 (W:Y)"},
    {"nome": "PCC", "plan": "AY", "cab_plan": "Valor PCC", "conf": ["AA"],                     # PCC 5952
     "desc_conf": "PCC 5952 (AA)"},
    {"nome": "INSS", "plan": "BA", "cab_plan": "Valor INSS", "conf": ["AF", "AG", "AK"],        # INSS 2631 (AF,AG) + 2100 (AK)
     "desc_conf": "INSS 2631 (AF:AG) + INSS 2100 (AK)"},
    {"nome": "ISS", "plan": "BC", "cab_plan": "Valor ISSQN", "conf": ["O", "P", "Q", "R", "S", "T"],  # ISS 2090 e demais
     "desc_conf": "ISS 2090 (O:T)"},
]

# Nomes para achar os arquivos automaticamente na pasta do script
PADRAO_PLANILHAO = "*PLANILH*.xls*"
PADRAO_CONFERENCIA = "*CONFER*RETEN*.xls*"


# =====================================================================
# FUNÇÕES AUXILIARES
# =====================================================================

def col(letra):
    return column_index_from_string(letra)


def para_numero(v):
    """Converte o conteúdo da célula em número. ' - ', vazio, texto -> 0.
    Leitura do número em comum/numeros.py ("R$ 1.234,56", "(10,00)"...)."""
    return numero_br(v) or 0.0


def chave_nf(v):
    """Normaliza o número da nota: 000123 / 123.0 / 'NF 123' / '123/1' -> '123'."""
    if v is None:
        return None
    if isinstance(v, float):
        if v != v:  # NaN
            return None
        v = int(v) if v.is_integer() else v
    s = str(v).strip()
    if not s:
        return None
    m = re.search(r"\d+", s)
    if not m:
        return s.upper()
    return str(int(m.group()))


def chave_forn(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s or None


def texto(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).replace("\t", " ").strip()


def r2(x):
    return round(x + 0.0, 2)


# =====================================================================
# ESCOLHA DOS ARQUIVOS E DA ABA
# =====================================================================

def pasta_do_script():
    return os.path.dirname(os.path.abspath(sys.argv[0]))


def achar_arquivo(padrao, titulo):
    achados = [f for f in glob.glob(os.path.join(pasta_do_script(), padrao))
               if not os.path.basename(f).startswith("~$")
               and "Relatorio_Conferencia" not in os.path.basename(f)]
    if len(achados) == 1:
        return achados[0]
    # 0 ou vários -> pergunta
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        caminho = filedialog.askopenfilename(
            title=titulo, initialdir=pasta_do_script(),
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        root.destroy()
        return caminho or None
    except Exception:
        print(titulo)
        for i, f in enumerate(achados, 1):
            print(f"  {i}) {os.path.basename(f)}")
        resp = input("Número da lista ou caminho completo: ").strip().strip('"')
        if resp.isdigit() and 1 <= int(resp) <= len(achados):
            return achados[int(resp) - 1]
        return resp or None


def escolher_aba(wb, caminho_planilhao):
    nomes = wb.sheetnames
    # Tenta pelo mês do nome do PLANILHÃO (ex.: "PLANILHÃO 09" -> aba "09_2026")
    m = re.search(r"(\d{1,2})", os.path.basename(caminho_planilhao))
    if m:
        mes = m.group(1).zfill(2)
        candidatas = [n for n in nomes if n.strip().startswith(mes + "_")]
        if len(candidatas) == 1:
            return candidatas[0]
    print("\nAbas da conferência:")
    for i, n in enumerate(nomes, 1):
        print(f"  {i}) {n.strip()}")
    while True:
        resp = input("Qual aba conferir? (número): ").strip()
        if resp.isdigit() and 1 <= int(resp) <= len(nomes):
            return nomes[int(resp) - 1]


# =====================================================================
# LEITURA
# =====================================================================

def _norm_cab(v):
    return re.sub(r"\s+", " ", str(v or "")).strip().upper()


def localizar_colunas_impostos(ws):
    """Usa a letra configurada, mas confere pelo nome do cabeçalho (linha 1).
    Se o planilhão mudar de layout de novo, acha a coluna pelo nome."""
    cab = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    cab = [_norm_cab(v) for v in cab]
    i_imp = {}
    for imp in IMPOSTOS:
        idx = col(imp["plan"]) - 1
        esperado = _norm_cab(imp.get("cab_plan"))
        atual = cab[idx] if idx < len(cab) else ""
        if esperado and atual != esperado:
            if esperado in cab:
                novo = cab.index(esperado)
                print(f'  AVISO: "{imp["cab_plan"]}" não está na coluna {imp["plan"]}; '
                      f'usando a coluna {get_column_letter(novo + 1)}.')
                idx = novo
            else:
                print(f'  AVISO: cabeçalho "{imp["cab_plan"]}" não encontrado; '
                      f'usando a coluna {imp["plan"]} ("{atual}").')
        i_imp[imp["nome"]] = idx
        imp["plan_usada"] = get_column_letter(idx + 1)
    return i_imp


def ler_planilhao(caminho):
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    i_nf, i_cod = col(PLAN_COL_NF) - 1, col(PLAN_COL_COD_FORN) - 1
    i_cnpj, i_raz = col(PLAN_COL_CNPJ) - 1, col(PLAN_COL_RAZAO) - 1
    i_imp = localizar_colunas_impostos(ws)

    # notas[nf][cod_forn] = {"razao", "cnpj", "IRRF": x, ...}
    notas = defaultdict(dict)
    linhas = 0
    for row in ws.iter_rows(min_row=PLAN_LINHA_INICIAL, values_only=True):
        if not row or len(row) <= i_nf:
            continue
        nf = chave_nf(row[i_nf])
        if nf is None:
            continue
        linhas += 1
        forn = chave_forn(row[i_cod]) if len(row) > i_cod else None
        reg = notas[nf].setdefault(forn, {
            "razao": texto(row[i_raz]) if len(row) > i_raz else "",
            "cnpj": texto(row[i_cnpj]) if len(row) > i_cnpj else "",
            **{imp["nome"]: 0.0 for imp in IMPOSTOS}})
        for nome, idx in i_imp.items():
            if len(row) > idx:
                reg[nome] += para_numero(row[idx])
    wb.close()
    return notas, linhas


def ler_conferencia(caminho, aba):
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb[aba]
    notas = defaultdict(dict)
    linhas = 0
    formulas_sem_valor = False
    for r in range(CONF_LINHA_INICIAL, ws.max_row + 1):
        parada = ws[f"{CONF_COL_PARADA}{r}"].value
        if isinstance(parada, str) and parada.strip().upper() == "TOTAL":
            break
        nf = chave_nf(ws[f"{CONF_COL_TITULO}{r}"].value)
        if nf is None:
            continue
        linhas += 1
        forn = chave_forn(ws[f"{CONF_COL_COD_FORN}{r}"].value)
        reg = notas[nf].setdefault(forn, {
            "razao": texto(ws[f"{CONF_COL_FORNECEDOR}{r}"].value),
            "cnpj": texto(ws[f"{CONF_COL_CNPJ}{r}"].value),
            "linhas": [],
            **{imp["nome"]: 0.0 for imp in IMPOSTOS}})
        reg["linhas"].append(r)
        for imp in IMPOSTOS:
            for c in imp["conf"]:
                reg[imp["nome"]] += para_numero(ws[f"{c}{r}"].value)
        if ws[f"I{r}"].value is not None and ws[f"K{r}"].value is None:
            formulas_sem_valor = True
    wb.close()
    return notas, linhas, formulas_sem_valor


# =====================================================================
# COMPARAÇÃO
# =====================================================================

def juntar(regs):
    """Soma vários fornecedores/linhas num registro só."""
    out = {"razao": "", "cnpj": "", "linhas": []}
    for imp in IMPOSTOS:
        out[imp["nome"]] = 0.0
    for reg in regs:
        out["razao"] = out["razao"] or reg.get("razao", "")
        out["cnpj"] = out["cnpj"] or reg.get("cnpj", "")
        out["linhas"] += reg.get("linhas", [])
        for imp in IMPOSTOS:
            out[imp["nome"]] += reg[imp["nome"]]
    return out


def tem_retencao(reg):
    return any(abs(reg[imp["nome"]]) > TOLERANCIA for imp in IMPOSTOS)


def comparar(plan, conf):
    """Gera a lista de resultados, uma linha por nota (ou nota+fornecedor)."""
    pares = []   # (nf, reg_plan ou None, reg_conf ou None, observacao)
    for nf in set(plan) | set(conf):
        p, c = plan.get(nf, {}), conf.get(nf, {})
        if len(p) <= 1 and len(c) <= 1:
            obs = ""
            if p and c:
                fp, fc = next(iter(p)), next(iter(c))
                if fp and fc and fp != fc:
                    obs = f"Cód. fornecedor diferente (planilhão {fp} / conferência {fc})"
            pares.append((nf, juntar(p.values()) if p else None,
                          juntar(c.values()) if c else None, obs))
        else:
            # mesma NF para fornecedores diferentes -> casa pelo código do fornecedor
            obs = "NF repetida para mais de um fornecedor - casado pelo cód. fornecedor"
            for forn in set(p) | set(c):
                pares.append((nf, juntar([p[forn]]) if forn in p else None,
                              juntar([c[forn]]) if forn in c else None, obs))

    resultados, ignoradas = [], 0
    for nf, rp, rc, obs in pares:
        if rp is not None and rc is None and not tem_retencao(rp):
            ignoradas += 1        # nota sem retenção e fora da conferência: normal
            continue
        linha = {"nf": nf, "obs": obs,
                 "razao": (rp or {}).get("razao") or (rc or {}).get("razao", ""),
                 "cnpj": (rp or {}).get("cnpj") or (rc or {}).get("cnpj", ""),
                 "linhas_conf": ", ".join(str(x) for x in (rc or {}).get("linhas", []))}
        if rc is None:
            linha["geral"] = "FALTANDO NA CONFERÊNCIA"
        elif rp is None:
            linha["geral"] = "FALTANDO NO PLANILHÃO"
        else:
            linha["geral"] = "OK"
        for imp in IMPOSTOS:
            n = imp["nome"]
            vp = r2(rp[n]) if rp else None
            vc = r2(rc[n]) if rc else None
            if rp is None or rc is None:
                existente = vp if rp is not None else vc
                st = "FALTANDO" if abs(existente) > TOLERANCIA else "-"
                dif = None
            else:
                dif = r2(vp - vc)
                if abs(dif) <= TOLERANCIA + 1e-9:
                    st = "OK"
                elif vc == 0:
                    st = "FALTA NA CONFERÊNCIA"
                elif vp == 0:
                    st = "FALTA NO PLANILHÃO"
                else:
                    st = "NÃO BATE"
                if st != "OK":
                    linha["geral"] = "NÃO BATE"
            linha[n] = (vp, vc, dif, st)
        resultados.append(linha)

    ordem = {"NÃO BATE": 0, "FALTANDO NA CONFERÊNCIA": 1, "FALTANDO NO PLANILHÃO": 2, "OK": 3}
    resultados.sort(key=lambda x: (ordem[x["geral"]], int(x["nf"]) if x["nf"].isdigit() else 0))
    return resultados, ignoradas


# =====================================================================
# RELATÓRIO EM EXCEL
# =====================================================================

VERDE = PatternFill("solid", start_color="C6EFCE")
VERMELHO = PatternFill("solid", start_color="FFC7CE")
AMARELO = PatternFill("solid", start_color="FFEB9C")
CINZA = PatternFill("solid", start_color="D9D9D9")
AZUL = PatternFill("solid", start_color="1F4E78")
FONTE = Font(name="Arial", size=10)
NEGRITO = Font(name="Arial", size=10, bold=True)
BRANCO = Font(name="Arial", size=10, bold=True, color="FFFFFF")
FINA = Side(style="thin", color="BFBFBF")
BORDA = Border(left=FINA, right=FINA, top=FINA, bottom=FINA)
MOEDA = '#,##0.00;[Red]-#,##0.00;"-"'


def cor_status(st):
    if st == "-":
        return CINZA
    if st == "OK":
        return VERDE
    if st.startswith("FALT"):
        return AMARELO
    return VERMELHO


def gerar_relatorio(resultados, ignoradas, info, destino):
    wb = openpyxl.Workbook()

    # ------------------------- DETALHE -------------------------
    ws = wb.active
    ws.title = "Detalhe"
    cab1 = ["", "", "", "", ""]
    cab2 = ["Nota Fiscal", "Fornecedor", "CNPJ", "Situação geral", "Linha(s) na conferência"]
    for imp in IMPOSTOS:
        cab1 += [imp["nome"], "", "", ""]
        cab2 += ["Planilhão", "Conferência", "Diferença", "Status"]
    cab1.append("")
    cab2.append("Observação")
    ws.append(cab1)
    ws.append(cab2)

    c = 6
    for imp in IMPOSTOS:
        ws.merge_cells(start_row=1, start_column=c, end_row=1, end_column=c + 3)
        ws.cell(1, c).value = f'{imp["nome"]}  -  planilhão {imp.get("plan_usada", imp["plan"])}  x  {imp["desc_conf"]}'
        c += 4
    for cell in ws[1] + ws[2]:
        cell.font = BRANCO
        cell.fill = AZUL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDA

    for res in resultados:
        linha = [int(res["nf"]) if res["nf"].isdigit() else res["nf"],
                 res["razao"], res["cnpj"], res["geral"], res["linhas_conf"]]
        for imp in IMPOSTOS:
            linha += list(res[imp["nome"]])
        linha.append(res["obs"])
        ws.append(linha)
        r = ws.max_row
        for cell in ws[r]:
            cell.font = FONTE
            cell.border = BORDA
        ws.cell(r, 4).fill = cor_status(res["geral"])
        ws.cell(r, 4).font = NEGRITO
        c = 6
        for imp in IMPOSTOS:
            for k in range(3):
                ws.cell(r, c + k).number_format = MOEDA
            ws.cell(r, c + 3).fill = cor_status(res[imp["nome"]][3])
            c += 4

    larguras = [12, 38, 17, 24, 14] + [13, 13, 12, 21] * len(IMPOSTOS) + [55]
    for i, w in enumerate(larguras, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 28
    ws.freeze_panes = "B3"
    ws.auto_filter.ref = f"A2:{get_column_letter(ws.max_column)}{max(ws.max_row, 2)}"

    # ------------------------- RESUMO -------------------------
    rs = wb.create_sheet("Resumo", 0)
    cont = defaultdict(int)
    for res in resultados:
        cont[res["geral"]] += 1

    rs["A1"] = "CONFERÊNCIA DE RETENÇÕES"
    rs["A1"].font = Font(name="Arial", size=14, bold=True, color="1F4E78")
    dados = [
        ("Gerado em", info["quando"]),
        ("Planilhão", info["planilhao"]),
        ("Conferência", info["conferencia"]),
        ("Aba conferida", info["aba"]),
        ("Tolerância (R$)", TOLERANCIA),
        ("Linhas lidas no planilhão (itens)", info["linhas_plan"]),
        ("Linhas lidas na conferência", info["linhas_conf"]),
        ("Notas do planilhão sem retenção (fora do relatório)", ignoradas),
    ]
    r = 3
    for k, v in dados:
        rs.cell(r, 1, k).font = NEGRITO
        rs.cell(r, 2, v).font = FONTE
        r += 1

    r += 1
    rs.cell(r, 1, "Situação").font = BRANCO
    rs.cell(r, 2, "Qtd. notas").font = BRANCO
    rs.cell(r, 1).fill = rs.cell(r, 2).fill = AZUL
    r += 1
    for st in ["OK", "NÃO BATE", "FALTANDO NA CONFERÊNCIA", "FALTANDO NO PLANILHÃO"]:
        rs.cell(r, 1, st).fill = cor_status(st)
        rs.cell(r, 1).font = NEGRITO
        rs.cell(r, 2, cont[st]).font = FONTE
        r += 1
    rs.cell(r, 1, "Total de notas no relatório").font = NEGRITO
    rs.cell(r, 2, len(resultados)).font = NEGRITO

    r += 2
    for i, t in enumerate(["Imposto", "Total planilhão", "Total conferência", "Diferença", "Notas com problema"], 1):
        rs.cell(r, i, t).font = BRANCO
        rs.cell(r, i).fill = AZUL
    r += 1
    for imp in IMPOSTOS:
        n = imp["nome"]
        tp = r2(sum(x[n][0] or 0 for x in resultados))
        tc = r2(sum(x[n][1] or 0 for x in resultados))
        prob = sum(1 for x in resultados if x[n][3] not in ("OK", "-"))
        rs.cell(r, 1, n).font = NEGRITO
        rs.cell(r, 2, tp).number_format = MOEDA
        rs.cell(r, 3, tc).number_format = MOEDA
        rs.cell(r, 4, r2(tp - tc)).number_format = MOEDA
        rs.cell(r, 5, prob)
        for i in range(2, 6):
            rs.cell(r, i).font = FONTE
        r += 1

    r += 1
    legenda = [
        ("OK", "Valor do planilhão igual ao da conferência (dentro da tolerância)."),
        ("NÃO BATE", "Os dois lados têm valor, mas diferentes."),
        ("FALTA NA CONFERÊNCIA", "Planilhão tem o imposto e a conferência está zerada."),
        ("FALTA NO PLANILHÃO", "Conferência tem o imposto e o planilhão está zerado."),
        ("FALTANDO NA CONFERÊNCIA", "Nota com retenção no planilhão que não existe na conferência (coluna TÍTULO)."),
        ("FALTANDO NO PLANILHÃO", "Título da conferência que não existe no planilhão."),
    ]
    rs.cell(r, 1, "Legenda").font = NEGRITO
    r += 1
    for st, d in legenda:
        rs.cell(r, 1, st).fill = cor_status(st)
        rs.cell(r, 1).font = FONTE
        rs.cell(r, 2, d).font = FONTE
        r += 1

    rs.column_dimensions["A"].width = 48
    rs.column_dimensions["B"].width = 60
    for L in "CDE":
        rs.column_dimensions[L].width = 20

    wb.save(destino)


# =====================================================================
# PRINCIPAL
# =====================================================================

def main():
    print("=" * 60)
    print(" CONFERÊNCIA DE RETENÇÕES  -  PLANILHÃO x CONFERÊNCIA")
    print("=" * 60)

    cam_plan = achar_arquivo(PADRAO_PLANILHAO, "Selecione o PLANILHÃO")
    if not cam_plan:
        print("Nenhum PLANILHÃO selecionado.")
        return
    cam_conf = achar_arquivo(PADRAO_CONFERENCIA, "Selecione a 2-CONFERÊNCIA RETENÇÕES")
    if not cam_conf:
        print("Nenhuma CONFERÊNCIA selecionada.")
        return
    print(f"Planilhão  : {os.path.basename(cam_plan)}")
    print(f"Conferência: {os.path.basename(cam_conf)}")

    wb_tmp = openpyxl.load_workbook(cam_conf, read_only=True)
    aba = escolher_aba(wb_tmp, cam_plan)
    wb_tmp.close()
    print(f"Aba        : {aba.strip()}")

    print("\nLendo planilhão...")
    plan, linhas_plan = ler_planilhao(cam_plan)
    print(f"  {linhas_plan} itens em {len(plan)} notas")

    print("Lendo conferência...")
    conf, linhas_conf, sem_valor = ler_conferencia(cam_conf, aba)
    print(f"  {linhas_conf} linhas com TÍTULO preenchido")
    if sem_valor:
        print("\n  ATENÇÃO: há fórmulas sem valor calculado na conferência.")
        print("  Abra a planilha no Excel, salve (Ctrl+S), feche e rode de novo.\n")

    resultados, ignoradas = comparar(plan, conf)

    agora = datetime.now()
    destino = os.path.join(
        pasta_do_script(),
        f"Relatorio_Conferencia_Retencoes_{aba.strip()}_{agora:%Y%m%d_%H%M}.xlsx")
    info = {"quando": agora.strftime("%d/%m/%Y %H:%M"),
            "planilhao": os.path.basename(cam_plan),
            "conferencia": os.path.basename(cam_conf),
            "aba": aba.strip(), "linhas_plan": linhas_plan, "linhas_conf": linhas_conf}
    gerar_relatorio(resultados, ignoradas, info, destino)

    cont = defaultdict(int)
    for x in resultados:
        cont[x["geral"]] += 1
    print("\nRESULTADO")
    for st in ["OK", "NÃO BATE", "FALTANDO NA CONFERÊNCIA", "FALTANDO NO PLANILHÃO"]:
        print(f"  {st:<26}: {cont[st]}")
    print(f"\nRelatório salvo em:\n  {destino}")

    try:
        os.startfile(destino)   # abre no Excel (Windows)
    except Exception:
        pass


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print(f"\nERRO: arquivo em uso ou sem permissão -> {e}")
        print("Feche o relatório anterior no Excel e rode de novo.")
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"\nERRO: {e}")
    if sys.stdin is not None and sys.stdin.isatty():
        input("\nPressione Enter para fechar...")
