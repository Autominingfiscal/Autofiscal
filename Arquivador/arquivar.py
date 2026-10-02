#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ARQUIVAR DOCUMENTOS DO DIA
==========================================================================
Distribui os arquivos soltos de uma pasta de dia nas subpastas
"NF <nota> R. PORTO", e cria as pastas que faltarem.

Nao envolve planilha nenhuma. Nao precisa do Excel.

Roda inteiro so com o que ja vem no Python. A UNICA dependencia opcional e o
Pillow (pip install --user pillow), e so para uma coisa: ler o numero do
ticket dentro dos PDFs DIGITALIZADOS, que sao imagem pura - a biblioteca
padrao nao decodifica JPEG. Sem Pillow, esses vao pela ordem, como antes.

    TIPO                  EXEMPLO DE NOME                 LIGA PELA...
    XML                   2010030018964.xml               nota (digitos finais)
    Nota fiscal           NF 18964 R. PORTO 19569.pdf     nota E ticket
    Pre-calculo           PC 19569 REGINALDO.txt          ticket
    Ticket de pesagem     TICKET 19569.pdf                ticket
    Ticket assinado       doc12881920260917155023_001.pdf numero LIDO dentro da
                                                          folha; se nao der, a
                                                          ordem da digitalizacao
    Ordem de carregamento ALYSSON ALVES - OC 17-09-2026   motorista lido DENTRO
                                                          do ticket de pesagem

Uso:
    python arquivar.py                      pasta do proprio script
    python arquivar.py "D:\\...\\17.09"       pasta indicada
    python arquivar.py --faixa 18964-18998  cria tambem essa faixa de pastas
    python arquivar.py --desfazer           desfaz a ultima execucao da pasta
    python arquivar.py --simular            so mostra o plano, nao move nada
    python arquivar.py --sim                nao pergunta, executa direto

Codigos de saida: 0 tudo certo | 1 erro | 2 terminou com pendencias
"""

import argparse
import csv
import hashlib
import io
import json
import os
import re
import shutil
import sys
import unicodedata
import zlib
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

VERSAO = "1.0"

ARQ_RELATORIO = "_relatorio_arquivamento.txt"
ARQ_RELATORIO_CSV = "_relatorio_arquivamento.csv"

# Nomes antigos: ficavam na pasta do dia. Hoje o diario do desfazer e o cache
# moram fora dela (veja dir_estado). Continuam listados aqui para serem
# ignorados e apagados se ainda existirem de uma rodada anterior.
ARQ_DIARIO_VELHO = "_arquivamento_desfazer.json"
ARQ_CACHE_VELHO = "_arquivamento_cache.json"

EXT_IGNORADAS = {".bat", ".cmd", ".ps1", ".py", ".pyc"}
ARQ_INTERNOS = {ARQ_RELATORIO, ARQ_RELATORIO_CSV, ARQ_DIARIO_VELHO, ARQ_CACHE_VELHO}

SUFIXO_PADRAO = "R. PORTO"
CONECTORES = {"DA", "DE", "DO", "DOS", "DAS", "E"}


# =====================================================================
#  CAMINHOS LONGOS NO WINDOWS
# =====================================================================
# OneDrive + "2026\09. SETEMBRO\17.09\NF 18964 R. PORTO\..." passa fácil dos
# 260 caracteres que o Windows aceita por padrao. O prefixo \\?\ levanta esse
# limite. Em Linux/Mac a funcao nao faz nada.
def longo(caminho):
    if os.name != "nt":
        return caminho
    c = os.path.abspath(caminho)
    if c.startswith("\\\\?\\"):
        return c
    if c.startswith("\\\\"):
        return "\\\\?\\UNC\\" + c[2:]
    return "\\\\?\\" + c


# =====================================================================
#  CLASSIFICACAO PELO NOME DO ARQUIVO
# =====================================================================
class Doc:
    __slots__ = ("arquivo", "caminho", "tipo", "nota", "ticket", "seq",
                 "motorista", "destinos", "situacao")

    def __init__(self, arquivo, caminho=""):
        self.arquivo = arquivo
        self.caminho = caminho
        self.tipo = "?"
        self.nota = 0
        self.ticket = 0
        self.seq = 0
        self.motorista = ""
        self.destinos = []
        self.situacao = ""


def classificar(nome, caminho=""):
    d = Doc(nome, caminho)
    base, ext = os.path.splitext(nome)
    ext = ext.lower().lstrip(".")

    # XML da nota: so digitos, a nota nos ultimos 5
    if ext == "xml" and re.fullmatch(r"\d{6,}", base):
        d.tipo = "XML"
        d.nota = int(base[-5:])
        return d

    # NOTA FISCAL: "NF <nota> ... <ticket>"
    m = re.match(r"(?i)^NF\s+(\d+)\b", base)
    if m:
        d.tipo = "NOTA FISCAL"
        d.nota = int(m.group(1))
        nums = re.findall(r"\d+", base)
        if len(nums) >= 2:
            d.ticket = int(nums[-1])
        return d

    # PRE-CALCULO: "PC <ticket> <motorista>"
    m = re.match(r"(?i)^PC\s+(\d+)\b", base)
    if m:
        d.tipo = "PRE-CALCULO"
        d.ticket = int(m.group(1))
        return d

    # TICKET ASSINADO ja renomeado numa rodada anterior
    m = re.match(r"(?i)^TICKET\s+ASSINADO\s+(\d+)", base)
    if m:
        d.tipo = "TICKET ASSINADO"
        d.nota = int(m.group(1))
        d.seq = -1
        return d

    # TICKET DE PESAGEM: "TICKET <ticket>"
    m = re.match(r"(?i)^TICKET\s+(\d+)\b", base)
    if m:
        d.tipo = "TICKET"
        d.ticket = int(m.group(1))
        return d

    # ORDEM DE CARREGAMENTO: "<motorista> - OC <data>"
    m = re.match(r"(?i)^(.+?)\s-\sOC\s", base)
    if m:
        d.tipo = "ORDEM CARREGAMENTO"
        d.motorista = m.group(1).strip()
        return d

    # TICKET ASSINADO escaneado: "doc<numeros>_<seq>"
    m = re.match(r"(?i)^doc\d+_(\d+)$", base)
    if m:
        d.tipo = "TICKET ASSINADO"
        d.seq = int(m.group(1))
        return d

    d.situacao = "tipo nao reconhecido - fica onde esta"
    return d


# =====================================================================
#  LEITURA DE PDF (so com a biblioteca padrao)
# =====================================================================
# Le os textos do PDF COM A POSICAO de cada um. A posicao e o que permite achar
# o nome do motorista: ele e a linha logo ACIMA do rotulo "Assinatura do
# Motorista".
#
# Cobre PDF sem compressao (o caso do ticket) e com FlateDecode, e monta o texto
# a partir de Tj, TJ, ' e ", com a posicao vinda de Td, TD, Tm e T*. Isso e mais
# do que o ticket precisa, mas e o que faz a leitura funcionar tambem nos PDFs
# que escrevem letra por letra, como a DANFE.

_RX_NUM = r"[-+]?\d*\.?\d+"


def _descomprimir_streams(bruto):
    """Devolve o conteudo do PDF com os streams comprimidos ja abertos."""
    partes = [bruto.decode("latin-1")]
    for m in re.finditer(rb"stream\r?\n", bruto):
        ini = m.end()
        fim = bruto.find(b"endstream", ini)
        if fim <= ini:
            continue
        dados = bruto[ini:fim]
        if not dados or (dados[0] & 0x0F) != 8:      # nao parece zlib
            continue
        try:
            partes.append(zlib.decompress(dados).decode("latin-1"))
        except zlib.error:
            try:
                partes.append(zlib.decompressobj().decompress(dados).decode("latin-1"))
            except Exception:
                pass
    return "\n".join(partes)


def _texto_de_literal(s):
    """Resolve os escapes de uma string literal de PDF: \\( \\) \\\\ \\n \\251 ..."""
    saida = []
    i = 0
    while i < len(s):
        c = s[i]
        if c == "\\" and i + 1 < len(s):
            p = s[i + 1]
            if p in "()\\":
                saida.append(p); i += 2; continue
            if p == "n": saida.append("\n"); i += 2; continue
            if p == "r": saida.append("\r"); i += 2; continue
            if p == "t": saida.append("\t"); i += 2; continue
            if p.isdigit():
                oct_ = ""
                j = i + 1
                while j < len(s) and len(oct_) < 3 and s[j].isdigit():
                    oct_ += s[j]; j += 1
                try:
                    saida.append(chr(int(oct_, 8)))
                except ValueError:
                    pass
                i = j
                continue
            saida.append(p); i += 2; continue
        saida.append(c)
        i += 1
    return "".join(saida)


def _hex_para_texto(h):
    h = re.sub(r"\s", "", h)
    if len(h) % 2:
        h += "0"
    try:
        b = bytes.fromhex(h)
    except ValueError:
        return ""
    # heuristica simples de UTF-16BE (BOM) usado por alguns geradores
    if b[:2] == b"\xfe\xff":
        try:
            return b[2:].decode("utf-16-be", "ignore")
        except Exception:
            return ""
    return b.decode("latin-1", "ignore")


def pdf_textos(caminho):
    """Lista de (x, y, texto). Lista vazia se nao der para ler."""
    try:
        with open(longo(caminho), "rb") as fh:
            bruto = fh.read()
    except OSError:
        return []

    txt = _descomprimir_streams(bruto)
    itens = []
    x = y = 0.0
    # percorre os operadores de texto na ordem em que aparecem
    padrao = re.compile(
        r"(?P<bt>\bBT\b)"
        r"|(?P<td>(%(n)s)\s+(%(n)s)\s+(Td|TD)\b)"
        r"|(?P<tm>(%(n)s)\s+(%(n)s)\s+(%(n)s)\s+(%(n)s)\s+(%(n)s)\s+(%(n)s)\s+Tm\b)"
        r"|(?P<tstar>T\*)"
        r"|(?P<tj>\((?P<lit>(?:\\.|[^\\()])*)\)\s*(?:Tj|'|\"))"
        r"|(?P<tjhex><(?P<hex>[0-9A-Fa-f\s]*)>\s*Tj)"
        r"|(?P<tjarr>\[(?P<arr>(?:\\.|[^\]\\]|\\\])*)\]\s*TJ)"
        % {"n": _RX_NUM},
        re.S,
    )
    for m in padrao.finditer(txt):
        if m.group("bt"):
            # BT reinicia a matriz de texto. Sem isto os Td de cada bloco iriam
            # se somando e as coordenadas cresceriam ate nao significar mais nada
            # (foi o que aconteceu na primeira versao: y = 23.751 numa pagina A4).
            x = y = 0.0
        elif m.group("td"):
            g = m.group(0).split()
            try:
                dx, dy = float(g[0]), float(g[1])
            except ValueError:
                continue
            # Td/TD sao relativos; sem acompanhar a matriz inteira, somar ja
            # separa as linhas, que e do que precisamos. Tm reposiciona.
            x, y = x + dx, y + dy
        elif m.group("tm"):
            g = m.group(0).split()
            try:
                x, y = float(g[4]), float(g[5])
            except (ValueError, IndexError):
                continue
        elif m.group("tstar"):
            y -= 12.0
        elif m.group("tj"):
            s = _texto_de_literal(m.group("lit"))
            if s.strip():
                itens.append((x, y, s))
        elif m.group("tjhex"):
            s = _hex_para_texto(m.group("hex"))
            if s.strip():
                itens.append((x, y, s))
        elif m.group("tjarr"):
            pedacos = re.findall(r"\((?:\\.|[^\\()])*\)", m.group("arr"), re.S)
            s = "".join(_texto_de_literal(p[1:-1]) for p in pedacos)
            if s.strip():
                itens.append((x, y, s))
    return itens


def _juntar_em_linhas(itens, tol=2.0):
    """Agrupa os textos que estao praticamente na mesma altura numa linha so.
    E o que resolve os PDFs que escrevem um caractere por vez."""
    linhas = []
    for x, y, t in sorted(itens, key=lambda i: (-i[1], i[0])):
        if linhas and abs(linhas[-1][0] - y) <= tol:
            linhas[-1][1].append((x, t))
        else:
            linhas.append([y, [(x, t)]])
    saida = []
    for y, pedacos in linhas:
        pedacos.sort()
        saida.append((pedacos[0][0], y, "".join(t for _, t in pedacos).strip()))
    return saida


def motorista_do_ticket(caminho):
    """Nome completo do motorista, lido acima de 'Assinatura do Motorista'."""
    itens = pdf_textos(caminho)
    if not itens:
        return ""
    for candidatos in (itens, _juntar_em_linhas(itens)):
        for lx, ly, lt in candidatos:
            if re.search(r"(?i)assinatura\s+do\s+motorista", lt):
                acima = [
                    (x, y, t) for (x, y, t) in candidatos
                    if ly + 3 <= y <= ly + 30 and lx - 60 < x < lx + 140 and len(t.strip()) >= 5
                ]
                if acima:
                    acima.sort(key=lambda i: i[1])
                    return acima[0][2].strip()
    return ""


# Modelos dos dez digitos da fonte do ticket, 12 x 18 pixels cada, tirados dos
# proprios campos numericos de um ticket escaneado (CNPJ e data). Sao o que
# permite ler o "N°Ticket" de dentro do PDF digitalizado.
MODELO_LARG, MODELO_ALT = 12, 18
MODELOS_DIGITOS = {
    "0": ("000011110000"
           "001111111100"
           "001111111100"
           "011100001110"
           "011000000110"
           "011000000111"
           "111000000011"
           "111000000011"
           "110000000011"
           "111000000011"
           "111000000011"
           "111000000111"
           "011000000110"
           "011000001110"
           "011100011110"
           "001111111100"
           "001111111000"
           "000011110000"
           ),
    "1": ("000001100000"
           "011111100000"
           "111111100000"
           "111111100000"
           "011011100000"
           "000001100000"
           "000001100000"
           "000001100000"
           "000001100000"
           "000001100000"
           "000001100000"
           "000001100000"
           "000001100000"
           "000001100000"
           "000001100000"
           "111111111111"
           "111111111111"
           "011111111110"
           ),
    "2": ("000011111100"
           "000111111110"
           "001111011111"
           "011100000111"
           "011100000011"
           "011000000011"
           "000000000111"
           "000000001111"
           "000000011110"
           "000000111100"
           "000000111000"
           "000011110000"
           "000011000000"
           "001111000000"
           "011110000010"
           "111111111111"
           "111111111111"
           "011111011111"
           ),
    "3": ("001111111000"
           "011111111100"
           "011110011110"
           "011000000110"
           "000000000110"
           "000000000110"
           "000000001110"
           "000011111100"
           "000011111100"
           "000000111110"
           "000000000110"
           "000000000111"
           "000000000111"
           "000000000111"
           "111000001110"
           "111111111110"
           "011111111000"
           "000111100000"
           ),
    "4": ("000000111100"
           "000000111100"
           "000001111100"
           "000011111100"
           "000011101100"
           "000011001100"
           "000111001100"
           "001110001100"
           "001100001100"
           "001100001100"
           "011000001100"
           "011000001100"
           "111111111111"
           "111111111111"
           "000000011100"
           "000001111110"
           "000011111111"
           "000001111111"
           ),
    "5": ("011000000000"
           "011111111111"
           "011111111111"
           "011100000000"
           "011000000000"
           "011000000000"
           "011101100000"
           "011111111100"
           "011111111110"
           "011110000111"
           "000000000011"
           "000000000011"
           "000000000011"
           "000000000011"
           "110000000011"
           "111100001111"
           "111111111110"
           "011111111100"
           ),
    "6": ("000001111110"
           "000011111111"
           "000111100110"
           "001111000000"
           "011100000000"
           "011000000000"
           "011001110000"
           "111111111000"
           "111111111110"
           "111110001110"
           "011000000110"
           "011000000110"
           "011000000110"
           "011100001110"
           "001110011110"
           "001111111100"
           "000111111000"
           "000000000000"
           ),
    "7": ("111110110111"
           "111111111111"
           "111111111111"
           "111000000111"
           "111000000111"
           "000000000110"
           "000000001110"
           "000000001110"
           "000000001100"
           "000000011100"
           "000000011100"
           "000000111000"
           "000000111000"
           "000000110000"
           "000001110000"
           "000001110000"
           "000001110000"
           "000001100000"
           ),
    "8": ("000111111000"
           "011111111110"
           "111110011111"
           "111000000111"
           "110000000011"
           "110000000011"
           "111000000111"
           "011111111110"
           "011111111110"
           "011111111110"
           "111100001111"
           "110000000011"
           "110000000011"
           "111000000011"
           "111000000111"
           "011111111111"
           "001111111110"
           "000111111000"
           ),
    "9": ("000111111000"
           "011111111100"
           "011111111110"
           "111100001111"
           "111000000111"
           "111000000111"
           "111000000111"
           "111000000111"
           "111100001111"
           "011111111111"
           "011111111111"
           "000011111111"
           "000000001111"
           "000000011110"
           "010001111110"
           "111111111100"
           "111111110000"
           "111111000000"
           ),
}


# =====================================================================
#  LEITURA DO TICKET DENTRO DO PDF DIGITALIZADO
# =====================================================================
# O "doc..._001.pdf" que sai do scanner e imagem pura: nenhuma letra de texto
# dentro dele. Mas o numero do ticket esta impresso, limpo, no alto da folha -
# e a fonte e monoespacada e sempre a mesma. Entao da para ler.
#
# O caminho e este:
#   1. tira do PDF o JPEG que o scanner gravou           (biblioteca padrao)
#   2. abre a imagem                                      (Pillow)
#   3. isola os glifos da faixa de cima por projecao de tinta
#   4. compara cada glifo com os dez modelos guardados aqui em cima
#   5. so aceita o resultado se o numero lido for UM DOS TICKETS DO DIA
#
# O passo 5 e o que torna isso seguro: a lista de tickets do dia ja saiu dos
# nomes das notas fiscais, entao a leitura nao precisa ser perfeita - precisa
# so ser boa o bastante para apontar um ticket conhecido, com folga sobre o
# segundo colocado. Sem essa folga, o arquivo volta para a regra da ordem.
#
# Pillow e a UNICA dependencia do script, e so para isto: a biblioteca padrao
# do Python nao decodifica JPEG. Sem ela, tudo o mais funciona igual.
try:
    from PIL import Image as _Image
except ImportError:
    _Image = None


def _jpegs_do_pdf(caminho):
    """Os JPEGs (DCTDecode) embutidos no PDF, como bytes."""
    try:
        with open(longo(caminho), "rb") as fh:
            d = fh.read()
    except OSError:
        return []
    saida = []
    for m in re.finditer(rb"<<(?:[^<>]|<<(?:[^<>]|<<[^>]*>>)*>>)*>>\s*stream\r?\n", d, re.S):
        if b"DCTDecode" not in m.group(0):
            continue
        ini = m.end()
        fim = d.find(b"endstream", ini)
        if fim > ini:
            saida.append(d[ini:fim].rstrip(b"\r\n"))
    return saida


def _glifos_da_faixa(bw, larg, alt, hmin=10, hmax=26, wmin=4, wmax=22):
    """Separa os glifos por projecao de tinta: primeiro as linhas, depois as
    colunas de cada linha. Funciona porque a fonte e monoespacada e o scan e
    limpo - e e muito mais rapido que sair rastreando componentes."""
    px = bw.load()
    linhas = []
    dentro = False
    for y in range(alt):
        tem = False
        for x in range(0, larg, 2):          # de dois em dois ja basta para achar a linha
            if px[x, y] < 128:
                tem = True
                break
        if tem and not dentro:
            ini = y; dentro = True
        elif not tem and dentro:
            dentro = False
            if hmin <= y - ini <= hmax * 2:
                linhas.append((ini, y - 1))
    if dentro and hmin <= alt - ini <= hmax * 2:
        linhas.append((ini, alt - 1))

    caixas = []
    for y0, y1 in linhas:
        col = []
        dentro = False
        for x in range(larg):
            tem = any(px[x, y] < 128 for y in range(y0, y1 + 1))
            if tem and not dentro:
                ini = x; dentro = True
            elif not tem and dentro:
                dentro = False
                col.append((ini, x - 1))
        if dentro:
            col.append((ini, larg - 1))
        for x0, x1 in col:
            if not (wmin <= x1 - x0 + 1 <= wmax):
                continue
            # aperta a caixa na vertical, para o glifo ficar colado nas bordas
            ys = [y for y in range(y0, y1 + 1) if any(px[x, y] < 128 for x in range(x0, x1 + 1))]
            if not ys:
                continue
            a, b = ys[0], ys[-1]
            if hmin <= b - a + 1 <= hmax:
                caixas.append((x0, a, x1, b, (y0 + y1) // 2))
    return caixas


def _normaliza(bw, c):
    g = bw.crop((c[0], c[1], c[2] + 1, c[3] + 1)).resize((MODELO_LARG, MODELO_ALT), _Image.BILINEAR)
    p = g.load()
    return [1 if p[x, y] < 128 else 0 for y in range(MODELO_ALT) for x in range(MODELO_LARG)]


def _classifica(v):
    """(digito, distancia, distancia do segundo colocado)"""
    r = sorted(((sum(1 for a, b in zip(v, m) if a != int(b)), d)
                for d, m in MODELOS_DIGITOS.items()))
    return r[0][1], r[0][0], r[1][0]


def _tenta_ler(im, candidatos, limiar, margem_minima):
    W, H = im.size
    # o numero fica no alto da folha, no primeiro terco da largura. Os tamanhos
    # sao proporcionais a altura da pagina, para 150, 200 ou 300 dpi darem no mesmo.
    faixa = im.crop((0, int(H * 0.03), int(W * 0.45), int(H * 0.12)))
    bw = faixa.point(lambda p: 0 if p < limiar else 255)
    alturas = (max(6, int(H * 0.0045)), max(12, int(H * 0.0135)))
    larguras = (max(3, int(H * 0.0022)), max(8, int(H * 0.0115)))
    caixas = _glifos_da_faixa(bw, faixa.size[0], faixa.size[1],
                              alturas[0], alturas[1], larguras[0], larguras[1])
    if not caixas:
        return 0
    passo = max(6, int(H * 0.009))          # espaco que separa duas sequencias
    caixas.sort(key=lambda c: (c[4], c[0]))
    seqs, atual = [], []
    for c in caixas:
        if atual and (abs(c[4] - atual[-1][4]) > passo // 2 or c[0] - atual[-1][2] > passo):
            seqs.append(atual); atual = []
        atual.append(c)
    if atual:
        seqs.append(atual)

    for s in seqs:
        if not (4 <= len(s) <= 8):
            continue
        lidos, pior_margem = "", 999
        for c in s:
            d, dist, dist2 = _classifica(_normaliza(bw, c))
            lidos += d
            pior_margem = min(pior_margem, dist2 - dist)
        if not lidos.isdigit() or pior_margem < margem_minima:
            continue
        n = int(lidos)
        if n in candidatos:
            return n
        # o zero da frente as vezes se perde ou sobra
        for alt_n in (int(lidos[1:]) if len(lidos) > 1 else 0, n * 10):
            if alt_n in candidatos:
                return alt_n
    return 0


def ler_ticket_escaneado(caminho, candidatos, margem_minima=8):
    """Numero do ticket lido dentro do PDF digitalizado, ou 0.

    'candidatos' e o conjunto dos tickets do dia. Uma leitura so e aceita se
    cair em cima de um deles - e com folga sobre o segundo colocado. Qualquer
    duvida devolve 0, e o arquivo volta para a regra da ordem.

    Tenta alguns limiares e uma leve correcao de inclinacao, porque a folha
    passa torta no alimentador do scanner e o toner varia de um dia para o
    outro. A primeira tentativa que fecha encerra a busca."""
    if _Image is None or not candidatos:
        return 0
    for bruto in _jpegs_do_pdf(caminho)[:2]:      # a primeira pagina basta
        try:
            im = _Image.open(io.BytesIO(bruto)).convert("L")
        except Exception:
            continue
        if im.size[0] < 400 or im.size[1] < 400:
            continue
        for ang in (0, -1.2, 1.2, -2.2, 2.2):
            giro = im if ang == 0 else im.rotate(ang, resample=_Image.BILINEAR,
                                                 expand=True, fillcolor=255)
            for limiar in (160, 128, 195):
                n = _tenta_ler(giro, candidatos, limiar, margem_minima)
                if n:
                    return n
    return 0



# =====================================================================
#  COMPARACAO DE NOMES DE MOTORISTA
# =====================================================================
def nome_simples(s):
    """Maiusculas, sem acento, sem os conectores."""
    if not s:
        return []
    n = unicodedata.normalize("NFD", s)
    n = "".join(c for c in n if unicodedata.category(c) != "Mn").upper()
    return [p for p in re.findall(r"[A-Z]+", n) if p not in CONECTORES]


def nome_compativel(curto, completo):
    """As palavras do nome curto aparecem no completo NA MESMA ORDEM.

    A ordem nao e frescura: existe um JOSE FRANCISCO DA SILVA e um
    FRANCISCO JOSE DE ALMEIDA na frota. Uma regra que so checasse
    "as duas palavras aparecem" trocaria um pelo outro."""
    pc, pl = nome_simples(curto), nome_simples(completo)
    if not pc or not pl:
        return False
    it = iter(pl)
    return all(any(p == q for q in it) for p in pc)


# =====================================================================
#  APOIO
# =====================================================================
def sha256(caminho, bloco=1 << 20):
    h = hashlib.sha256()
    with open(longo(caminho), "rb") as fh:
        for pedaco in iter(lambda: fh.read(bloco), b""):
            h.update(pedaco)
    return h.hexdigest()


def nome_livre(pasta, nome, reservados=()):
    """Devolve um nome que ainda nao existe na pasta: 'x.pdf' -> 'x (2).pdf'.

    `reservados` sao destinos (os.path.normcase) que outro item do mesmo plano
    ja vai ocupar: contam como existentes."""
    base, ext = os.path.splitext(nome)
    i = 2
    novo = nome

    def ocupado(n):
        c = os.path.join(pasta, n)
        return os.path.normcase(c) in reservados or os.path.exists(longo(c))

    while ocupado(novo):
        novo = "%s (%d)%s" % (base, i, ext)
        i += 1
    return novo


# =====================================================================
#  ONDE FICA O QUE O SCRIPT PRECISA GUARDAR
# =====================================================================
# Na pasta do dia fica SO o relatorio. O diario do desfazer e o cache dos
# motoristas sao coisa de maquina, nao documento do embarque, entao vao para a
# pasta de dados do usuario - e a pasta do dia continua so com os documentos.
def dir_estado():
    if os.name == "nt":
        raiz = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        d = os.path.join(raiz, "arquivar-mvv")
    else:
        d = os.path.join(os.path.expanduser("~"), ".cache", "arquivar-mvv")
    try:
        os.makedirs(longo(d), exist_ok=True)
    except OSError:
        return None
    return d


def _apelido(pasta):
    """Nome de arquivo curto e unico para uma pasta: o nome dela + um resumo do
    caminho inteiro, para duas pastas '17.09' de meses diferentes nao se
    misturarem."""
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", os.path.basename(pasta)) or "pasta"
    resumo = hashlib.sha256(os.path.normcase(os.path.abspath(pasta)).encode("utf-8")).hexdigest()[:10]
    return "%s-%s" % (base[:40], resumo)


def arquivo_estado(pasta, qual):
    d = dir_estado()
    if not d:
        return None
    return os.path.join(d, "%s.%s.json" % (_apelido(pasta), qual))


def limpar_json_antigos(pasta):
    """Tira da pasta do dia os .json que as versoes anteriores gravavam la.
    Devolve os que foram apagados."""
    apagados = []
    for nome in (ARQ_DIARIO_VELHO, ARQ_CACHE_VELHO):
        p = os.path.join(pasta, nome)
        if os.path.exists(longo(p)):
            try:
                os.remove(longo(p))
                apagados.append(nome)
            except OSError:
                pass
    return apagados


def interpretar_notas(txt):
    """'18893-18909,18911-18925, 18930' -> [18893, ..., 18909, 18911, ..., 18930]"""
    saida = []
    for pedaco in re.split(r"[;,]", txt or ""):
        pedaco = pedaco.strip()
        if not pedaco:
            continue
        m = re.fullmatch(r"(\d+)\s*[-a]\s*(\d+)", pedaco)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if b >= a and b - a <= 2000:
                saida.extend(range(a, b + 1))
        elif pedaco.isdigit():
            saida.append(int(pedaco))
    vistas, limpo = set(), []
    for n in saida:
        if n not in vistas:
            vistas.add(n); limpo.append(n)
    return limpo


def escrever_texto(caminho, texto):
    """Texto para o usuario ler: com BOM, que e o que faz o Bloco de Notas e o
    Excel abrirem os acentos certos."""
    with open(longo(caminho), "w", encoding="utf-8-sig", newline="") as fh:
        fh.write(texto)


def escrever_json(caminho, obj):
    """Arquivo que o proprio script le depois: UTF-8 SEM BOM. Com BOM, o
    json.load falha e o cache/desfazer seria descartado em silencio.

    Grava num .tmp e troca de uma vez: interrompido no meio, o arquivo
    anterior continua inteiro."""
    tmp = caminho + ".tmp"
    with open(longo(tmp), "w", encoding="utf-8", newline="") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=1)
    os.replace(longo(tmp), longo(caminho))


def ler_json(caminho):
    with open(longo(caminho), encoding="utf-8-sig") as fh:   # aceita com e sem BOM
        return json.load(fh)


class Cache:
    """Guarda o motorista ja lido de cada ticket. Ler PDF e a parte lenta;
    numa segunda rodada isso sai de graca. A chave inclui tamanho e data do
    arquivo, entao ticket trocado e relido."""

    def __init__(self, pasta, qual="cache"):
        self.caminho = arquivo_estado(pasta, qual)
        self.dados = {}
        if not self.caminho:
            return
        try:
            d = ler_json(self.caminho)
            if isinstance(d, dict) and d.get("versao") == VERSAO:
                self.dados = d.get("motoristas", {})
        except Exception:
            pass

    @staticmethod
    def chave(caminho):
        try:
            st = os.stat(longo(caminho))
            return "%s|%d|%d" % (os.path.basename(caminho), st.st_size, int(st.st_mtime))
        except OSError:
            return os.path.basename(caminho)

    def get(self, caminho):
        return self.dados.get(self.chave(caminho))

    def put(self, caminho, valor):
        self.dados[self.chave(caminho)] = valor

    def salvar(self):
        if not self.caminho:
            return
        try:
            escrever_json(self.caminho, {"versao": VERSAO, "motoristas": self.dados})  # "motoristas" e so o nome do campo
        except OSError:
            pass


# =====================================================================
#  DESFAZER
# =====================================================================
def desfazer(pasta):
    caminho = arquivo_estado(pasta, "desfazer")
    velho = os.path.join(pasta, ARQ_DIARIO_VELHO)      # gravado pelas versoes antigas
    if caminho and os.path.exists(longo(caminho)):
        pass
    elif os.path.exists(longo(velho)):
        caminho = velho
    else:
        print("Nao ha nada para desfazer nesta pasta.")
        print("(o desfazer usa o diario gravado na ultima execucao, em %s)"
              % (dir_estado() or "?"))
        return 1
    diario = ler_json(caminho)

    acoes = diario.get("acoes", [])
    print("Execucao de %s, %d acao(oes)." % (diario.get("quando", "?"), len(acoes)))
    print()
    desfeitas = pulos = erros = 0
    # ao contrario: a ultima acao e a primeira a ser desfeita
    for a in reversed(acoes):
        destino, origem, tipo = a["destino"], a["origem"], a["acao"]
        try:
            if not os.path.exists(longo(destino)):
                print("  pulei (nao esta mais la): %s" % os.path.basename(destino))
                pulos += 1
                continue
            if tipo == "copiar":
                os.remove(longo(destino))
                desfeitas += 1
            else:
                if os.path.exists(longo(origem)):
                    print("  pulei (a origem ja existe): %s" % os.path.basename(origem))
                    pulos += 1
                    continue
                shutil.move(longo(destino), longo(origem))
                desfeitas += 1
        except OSError as e:
            print("  ERRO em %s: %s" % (os.path.basename(destino), e))
            erros += 1

    for p in reversed(diario.get("pastas_criadas", [])):
        try:
            if os.path.isdir(longo(p)) and not os.listdir(longo(p)):
                os.rmdir(longo(p))
        except OSError:
            pass

    print()
    print("Desfeitas: %d   Puladas: %d   Erros: %d" % (desfeitas, pulos, erros))
    if erros == 0:
        try:
            os.remove(longo(caminho))
        except OSError:
            pass
    return 0 if erros == 0 else 2


# =====================================================================
#  ROTINA PRINCIPAL
# =====================================================================
def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Arquiva os documentos do dia nas pastas das notas fiscais.",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pasta", nargs="?", default=None,
                    help="pasta do dia (padrao: a pasta onde este arquivo esta)")
    ap.add_argument("--faixa", default=None, metavar="18964-18998",
                    help="cria tambem as pastas dessa faixa de notas")
    ap.add_argument("--sem-faixa", action="store_true",
                    help="nao pergunta pela faixa extra")
    ap.add_argument("--simular", action="store_true",
                    help="mostra o plano e o relatorio, sem mover nada")
    ap.add_argument("--sim", "--yes", dest="sim", action="store_true",
                    help="nao pede confirmacao")
    ap.add_argument("--desfazer", action="store_true",
                    help="desfaz a ultima execucao feita nesta pasta")
    ap.add_argument("--somente-pastas", action="store_true",
                    help="so cria as pastas das notas, nao move arquivo nenhum")
    ap.add_argument("--sem-pausa", action="store_true",
                    help="nao segura a janela no fim")
    ap.add_argument("--assinados", default=None, metavar="LISTA",
                    help="notas dos tickets assinados, na ordem da digitalizacao "
                         "(ex: 18893-18909,18911-18925)")
    ap.add_argument("--sem-ocr", action="store_true",
                    help="nao le o numero dentro dos tickets digitalizados")
    ap.add_argument("--sem-csv", action="store_true",
                    help="grava so o relatorio .txt, sem o .csv")
    ap.add_argument("--sem-cache", action="store_true",
                    help="ignora o cache e le todos os tickets de novo")
    args = ap.parse_args(argv)

    pasta = args.pasta or os.path.dirname(os.path.abspath(__file__)) or os.getcwd()
    pasta = os.path.abspath(pasta.rstrip("\\/") or ".")

    def fim(codigo):
        if not args.sem_pausa and sys.stdin and sys.stdin.isatty():
            print()
            try:
                input("Terminou. Pressione Enter para fechar esta janela.")
            except (EOFError, KeyboardInterrupt):
                pass
        return codigo

    print()
    print("=" * 63)
    print(" ARQUIVAR DOCUMENTOS DO DIA   (python %d.%d, versao %s)"
          % (sys.version_info[0], sys.version_info[1], VERSAO))
    print("=" * 63)
    print("Pasta: %s" % pasta)
    print()

    if not os.path.isdir(longo(pasta)):
        print("ERRO: a pasta nao existe.")
        return fim(1)

    if args.desfazer:
        return fim(desfazer(pasta))

    # Versoes anteriores deixavam dois .json na pasta do dia. Hoje eles moram
    # fora dela; se ainda estiverem aqui, saem agora.
    if not args.simular:
        limpos = limpar_json_antigos(pasta)
        if limpos:
            print("Tirei da pasta os arquivos de controle das versoes anteriores: %s"
                  % ", ".join(limpos))
            print()

    # ---------------------------------------------------------- leitura ----
    soltos = []
    for nome in sorted(os.listdir(longo(pasta))):
        caminho = os.path.join(pasta, nome)
        if not os.path.isfile(longo(caminho)):
            continue
        if os.path.splitext(nome)[1].lower() in EXT_IGNORADAS:
            continue
        if nome in ARQ_INTERNOS:
            continue
        soltos.append(classificar(nome, caminho))

    # referencias: o que ja esta dentro das subpastas de nota conta tambem.
    # Sem isto, rodar uma segunda vez deixaria as ordens de carregamento orfas,
    # porque os tickets (de onde sai o nome do motorista) ja teriam saido da raiz.
    refs = [(d.arquivo, d.caminho, 0) for d in soltos]
    ja_arquivados = 0
    subpastas = {}
    for nome in sorted(os.listdir(longo(pasta))):
        cam = os.path.join(pasta, nome)
        if not os.path.isdir(longo(cam)):
            continue
        m = re.match(r"(?i)^NF\s+(\d+)\b", nome)
        if not m:
            continue
        n = int(m.group(1))
        subpastas.setdefault(n, []).append(nome)
        try:
            for f in sorted(os.listdir(longo(cam))):
                fc = os.path.join(cam, f)
                if os.path.isfile(longo(fc)):
                    refs.append((f, fc, n))
                    ja_arquivados += 1
        except OSError:
            pass

    # --------------------------------------- notas do dia e ticket -> nota ----
    # As notas do dia saem dos documentos da PROPRIA nota fiscal ja salvos na
    # pasta: o PDF "NF <nota> ..." ou o XML. Um dos dois basta para a nota
    # existir e ganhar pasta. Mas so o PDF traz nota e ticket no mesmo nome,
    # entao so ele monta a tabela ticket -> nota.
    t2n, notas_do_dia, com_pdf, so_xml = {}, set(), set(), set()
    for nome, caminho, nota_pasta in refs:
        d = classificar(nome, caminho)
        if d.tipo == "NOTA FISCAL" and d.nota:
            notas_do_dia.add(d.nota)
            com_pdf.add(d.nota)
            if d.ticket:
                t2n.setdefault(d.ticket, d.nota)
        elif d.tipo == "XML" and d.nota:
            notas_do_dia.add(d.nota)
            so_xml.add(d.nota)
        if d.tipo == "TICKET" and d.ticket and nota_pasta:
            notas_do_dia.add(nota_pasta)
            t2n.setdefault(d.ticket, nota_pasta)
    so_xml -= com_pdf
    ordem_notas = sorted(notas_do_dia)

    if ja_arquivados:
        print("Ja arquivados nas subpastas (usados como referencia): %d" % ja_arquivados)
    print("Notas do dia (PDFs de nota fiscal + XMLs): %d" % len(ordem_notas))
    print("Pares ticket->nota montados .............: %d" % len(t2n))
    if so_xml:
        print("  %d nota(s) so com o XML, sem o PDF da nota fiscal: %s"
              % (len(so_xml), ", ".join(str(n) for n in sorted(so_xml))))
        print("  A pasta e criada do mesmo jeito, mas sem o PDF nao da para saber o")
        print("  ticket delas - o pre-calculo e o ticket de pesagem ficam pendentes.")
    print()

    # ------------------------------------------------- criador de pastas ----
    faixa = args.faixa
    if faixa is None and not args.sem_faixa and sys.stdin and sys.stdin.isatty():
        print("CRIADOR DE PASTAS")
        print("  Vou garantir a pasta das %d notas fiscais ja salvas aqui." % len(ordem_notas))
        try:
            faixa = input("  Criar tambem uma faixa extra? (ex: 18964-18998, Enter pula): ")
        except (EOFError, KeyboardInterrupt):
            faixa = ""
    a_criar = list(ordem_notas)
    if faixa:
        m = re.fullmatch(r"\s*(\d+)\s*[-a]\s*(\d+)\s*", faixa)
        if m:
            ini, f = int(m.group(1)), int(m.group(2))
            if f >= ini and f - ini <= 2000:
                a_criar += [n for n in range(ini, f + 1) if n not in notas_do_dia]
            else:
                print("  Faixa invalida ou grande demais; ignorei.")
        else:
            print("  Nao entendi a faixa; ignorei. Formato: 18964-18998")

    pastas_criadas = []
    for n in sorted(set(a_criar)):
        if n in subpastas:
            continue
        nova = os.path.join(pasta, "NF %d %s" % (n, SUFIXO_PADRAO))
        if not args.simular:
            try:
                os.makedirs(longo(nova), exist_ok=True)
            except OSError as e:
                print("  nao consegui criar %s: %s" % (os.path.basename(nova), e))
                continue
        pastas_criadas.append(nova)
        subpastas.setdefault(n, []).append(os.path.basename(nova))
    print("  Pastas criadas agora: %d   (ja existiam: %d)"
          % (len(pastas_criadas), len(set(a_criar)) - len(pastas_criadas)))
    print()

    if args.somente_pastas:
        print("--somente-pastas: parei aqui, nenhum arquivo foi movido.")
        return fim(0)

    def pasta_da_nota(n):
        """Quando a nota tem mais de uma pasta, a 'R. PORTO' e a que recebe."""
        nomes = subpastas.get(n) or []
        for nm in nomes:
            if SUFIXO_PADRAO.upper() in nm.upper():
                return os.path.join(pasta, nm)
        if nomes:
            return os.path.join(pasta, nomes[0])
        return os.path.join(pasta, "NF %d %s" % (n, SUFIXO_PADRAO))

    # ------------------------------------------------- tickets assinados ----
    # O ticket assinado e o unico documento cujo nome (doc..._001) nao diz nada:
    # ele e atribuido pela ORDEM da digitalizacao. Para isso a quantidade de
    # digitalizados tem que bater com a quantidade de notas que eles cobrem -
    # senao, um unico furo no meio empurra todos os seguintes para a nota errada.
    #
    # A dificuldade e saber QUAIS notas a leva cobre. Sao testadas duas listas,
    # nesta ordem, e vale a primeira que bater com a quantidade de digitalizados:
    #
    #   1) todas as notas do dia que ainda nao tem ticket assinado.
    #      E o caso normal, e tambem o de escanear em duas levas.
    #
    #   2) so as notas que tem documento SOLTO nesta rodada.
    #      E o caso da nota que ja foi inteiramente arquivada num dia anterior:
    #      ela continua sendo nota do dia, mas nao faz parte desta leva de
    #      digitalizacao. Sem esta segunda tentativa, uma nota ja arquivada
    #      deixava a conta 32 para 33 e nenhum digitalizado era arquivado.
    scans_soltos = sorted([d for d in soltos if d.tipo == "TICKET ASSINADO" and d.seq >= 0],
                          key=lambda d: d.seq)

    notas_com_scan = set()
    for nome, _c, nota_pasta in refs:
        if nota_pasta and re.match(r"(?i)^TICKET ASSINADO\s", nome):
            notas_com_scan.add(nota_pasta)
    for d in soltos:
        if d.tipo == "TICKET ASSINADO" and d.seq == -1 and d.nota:
            notas_com_scan.add(d.nota)

    # notas com algum documento solto agora
    notas_soltas = set()
    for d in soltos:
        if d.tipo in ("XML", "NOTA FISCAL") and d.nota:
            notas_soltas.add(d.nota)
        elif d.tipo in ("PRE-CALCULO", "TICKET") and d.ticket in t2n:
            notas_soltas.add(t2n[d.ticket])

    lista_a = [n for n in ordem_notas if n not in notas_com_scan]
    lista_b = [n for n in lista_a if n in notas_soltas]

    if args.assinados:
        notas_sem_scan = interpretar_notas(args.assinados)
        criterio_scan = "lista informada em --assinados"
    elif len(scans_soltos) == len(lista_a) or not scans_soltos:
        notas_sem_scan = lista_a
        criterio_scan = "notas do dia sem ticket assinado"
    elif len(scans_soltos) == len(lista_b):
        notas_sem_scan = lista_b
        criterio_scan = "notas com documento solto nesta rodada"
        fora = [n for n in lista_a if n not in notas_soltas]
        print("Tickets assinados: %d digitalizados para %d notas do dia sem ticket assinado."
              % (len(scans_soltos), len(lista_a)))
        print("  A conta fecha com as %d notas que tem documento solto agora." % len(lista_b))
        print("  Fica(m) de fora: %s - ja estava(m) arquivada(s) antes desta rodada."
              % ", ".join(str(n) for n in fora))
        print()
    else:
        notas_sem_scan = lista_a
        criterio_scan = "nenhuma lista bateu"

    # --- ler o numero do ticket DENTRO de cada digitalizado -------------------
    # Isto vale mais do que qualquer regra de ordem: o numero esta impresso na
    # folha. So o que nao for lido com folga cai na regra da ordem.
    lidos_ocr = 0
    if scans_soltos and not args.sem_ocr and _Image is not None and t2n:
        cache_ocr = Cache(pasta, "escaneados")
        candidatos = set(t2n)

        def ler_scan(d):
            if not args.sem_cache:
                v = cache_ocr.get(d.caminho)
                if v is not None:
                    return d, int(v), True
            return d, ler_ticket_escaneado(d.caminho, candidatos), False

        with ThreadPoolExecutor(max_workers=min(8, (os.cpu_count() or 2) * 4)) as ex:
            for d, tk, do_cache in ex.map(ler_scan, scans_soltos):
                if not do_cache:
                    cache_ocr.put(d.caminho, tk)
                if tk and tk in t2n:
                    d.ticket = tk
                    d.nota = t2n[tk]
                    d.seq = -2          # -2 = resolvido por leitura, nao pela ordem
                    lidos_ocr += 1
        if not args.simular:
            cache_ocr.salvar()

        print("Tickets assinados: li o numero dentro de %d de %d digitalizados."
              % (lidos_ocr, len(scans_soltos)))
        if lidos_ocr < len(scans_soltos):
            print("  Os outros %d vao pela ordem da digitalizacao."
                  % (len(scans_soltos) - lidos_ocr))
        print()

        # os que foram lidos saem da fila da ordem, junto com as notas deles
        tomadas = {d.nota for d in scans_soltos if d.seq == -2}
        notas_sem_scan = [n for n in notas_sem_scan if n not in tomadas]
        scans_soltos = [d for d in scans_soltos if d.seq != -2]
        if lidos_ocr:
            criterio_scan = "%d lidos do PDF; resto: %s" % (lidos_ocr, criterio_scan)
    elif scans_soltos and not args.sem_ocr and _Image is None:
        print("Tickets assinados: o Pillow nao esta instalado, entao nao da para ler o")
        print("  numero dentro do PDF. Vao pela ordem da digitalizacao.")
        print("  Para ligar a leitura:  pip install --user pillow")
        print()

    # ---------------------------------------------------------- roteamento ----
    for d in soltos:
        if d.tipo in ("XML", "NOTA FISCAL"):
            if d.nota in notas_do_dia:
                d.destinos = [d.nota]
            else:
                d.situacao = "nota %d nao aparece em nenhum documento de nota fiscal desta pasta" % d.nota

        elif d.tipo in ("PRE-CALCULO", "TICKET"):
            if d.ticket in t2n:
                d.nota = t2n[d.ticket]
                d.destinos = [d.nota]
            else:
                d.situacao = "ticket %d sem nota fiscal correspondente nesta pasta" % d.ticket

        elif d.tipo == "TICKET ASSINADO":
            if d.seq == -2:
                # o numero do ticket foi lido dentro do PDF: nao ha o que supor
                d.destinos = [d.nota]
            elif d.seq == -1:
                if d.nota in notas_do_dia:
                    d.destinos = [d.nota]
                else:
                    d.situacao = "nota %d nao aparece nos documentos desta pasta" % d.nota
            elif len(scans_soltos) != len(notas_sem_scan):
                d.situacao = ("%d digitalizados para %d notas candidatas: a ordem nao "
                              "fecha, nao vou adivinhar. Use --assinados com a lista "
                              "certa (ex: --assinados 18893-18909,18911-18925)"
                              % (len(scans_soltos), len(notas_sem_scan)))
            else:
                i = scans_soltos.index(d)
                d.nota = notas_sem_scan[i]
                d.destinos = [d.nota]

    # ------------------------------------ ordem de carregamento -> notas ----
    # O nome da OC so tem motorista e data. A ponte ate a nota vem de DENTRO do
    # ticket de pesagem, que traz o nome completo do motorista:
    #     OC (motorista) -> ticket -> nota
    ocs = [d for d in soltos if d.tipo == "ORDEM CARREGAMENTO"]
    diag_oc = "Sem ordens de carregamento nesta pasta."
    if ocs:
        print("Lendo os tickets para achar o motorista de cada um...")
        tickets = [(nome, cam, npasta) for (nome, cam, npasta) in refs
                   if classificar(nome).tipo == "TICKET"
                   and os.path.splitext(nome)[1].lower() == ".pdf"]
        cache = Cache(pasta)

        def ler(t):
            nome, cam, _n = t
            if not args.sem_cache:
                v = cache.get(cam)
                if v is not None:
                    return nome, cam, v, True
            return nome, cam, motorista_do_ticket(cam), False

        lidos = 0
        sem_nome = 0
        sem_nota = 0
        do_cache = 0
        notas_do_motorista = {}
        with ThreadPoolExecutor(max_workers=min(8, (os.cpu_count() or 2) * 4)) as ex:
            for nome, cam, nomemot, veio_cache in ex.map(ler, tickets):
                if veio_cache:
                    do_cache += 1
                else:
                    cache.put(cam, nomemot)
                if not nomemot:
                    sem_nome += 1
                    continue
                lidos += 1
                tk = classificar(nome).ticket
                n = t2n.get(tk)
                if not n:
                    sem_nota += 1
                    continue
                notas_do_motorista.setdefault(nomemot, []).append(n)
        if not args.simular:
            cache.salvar()

        diag_oc = ("Tickets lidos: %d   motoristas identificados: %d   "
                   "sem nome no PDF: %d   sem nota: %d   do cache: %d"
                   % (lidos, len(notas_do_motorista), sem_nome, sem_nota, do_cache))
        print("  motoristas identificados: %d   tickets lidos: %d   sem nome: %d   "
              "sem nota: %d   (do cache: %d)"
              % (len(notas_do_motorista), lidos, sem_nome, sem_nota, do_cache))
        if lidos == 0:
            print("  Nenhum ticket foi lido. As ordens de carregamento dependem do nome")
            print("  do motorista, que esta dentro do TICKET <numero>.pdf.")
        print()

        for d in ocs:
            achados = [m for m in notas_do_motorista if nome_compativel(d.motorista, m)]
            if len(achados) == 1:
                d.destinos = sorted(set(notas_do_motorista[achados[0]]))
            elif len(achados) > 1:
                d.situacao = ("o nome '%s' bate com mais de um motorista (%s): nao vou chutar"
                              % (d.motorista, "; ".join(achados)))
            else:
                d.situacao = "nao achei o motorista '%s' em nenhum ticket desta pasta" % d.motorista

    # --------------------------------------------------------------- plano ----
    plano = []
    # destinos ja tomados por itens anteriores DESTE plano. Sem isto, dois
    # digitalizados lidos com o mesmo numero iriam para o mesmo nome, e o
    # shutil.move do Windows sobrescreve o primeiro sem avisar.
    reservados = set()
    for d in soltos:
        if d.destinos:
            for nf in d.destinos:
                pdest = pasta_da_nota(nf)
                nome_final = d.arquivo
                if d.tipo == "TICKET ASSINADO" and d.seq >= 0:
                    # arquivado pela ORDEM: leva a marca de conferir
                    nome_final = "TICKET ASSINADO %d (conferir)%s" % (
                        nf, os.path.splitext(d.arquivo)[1])
                elif d.tipo == "TICKET ASSINADO" and d.seq == -2:
                    # o numero foi LIDO dentro da folha: nao ha o que conferir
                    nome_final = "TICKET ASSINADO %d%s" % (
                        nf, os.path.splitext(d.arquivo)[1])
                alvo = os.path.join(pdest, nome_final)
                acao = "mover"
                if os.path.normcase(alvo) in reservados:
                    nome_final = nome_livre(pdest, nome_final, reservados)
                    acao = "mover como '%s' (outro arquivo desta rodada ia para o mesmo nome)" % nome_final
                elif os.path.exists(longo(alvo)):
                    try:
                        igual = sha256(alvo) == sha256(d.caminho)
                    except OSError:
                        igual = False
                    if igual:
                        acao = "JA ESTA LA - identico, nao mexo"
                    else:
                        nome_final = nome_livre(pdest, nome_final, reservados)
                        acao = "mover como '%s' (ja havia outro diferente)" % nome_final
                if acao.startswith("mover"):
                    reservados.add(os.path.normcase(os.path.join(pdest, nome_final)))
                plano.append({
                    "Arquivo": d.arquivo, "Tipo": d.tipo, "Nota": str(nf),
                    "Ticket": str(d.ticket) if d.ticket else "",
                    "Destino": os.path.basename(pdest) + os.sep + nome_final,
                    "Acao": acao,
                    "_pasta": pdest, "_nome": nome_final, "_origem": d.caminho,
                })
        else:
            plano.append({
                "Arquivo": d.arquivo, "Tipo": d.tipo, "Nota": "",
                "Ticket": str(d.ticket) if d.ticket else "",
                "Destino": "", "Acao": "PENDENTE: " + (d.situacao or "sem destino"),
                "_pasta": "", "_nome": "", "_origem": "",
            })

    mover = [p for p in plano if p["Acao"].startswith("mover")]
    pend = [p for p in plano if p["Acao"].startswith("PENDENTE")]
    ja = [p for p in plano if p["Acao"].startswith("JA ESTA LA")]

    print("RESUMO POR TIPO")
    for tipo in sorted({p["Tipo"] for p in plano}):
        g = [p for p in plano if p["Tipo"] == tipo]
        print("  %-20s total %3d   mover %3d   ja existe %3d   pendente %3d" % (
            tipo, len(g),
            sum(1 for p in g if p["Acao"].startswith("mover")),
            sum(1 for p in g if p["Acao"].startswith("JA ESTA LA")),
            sum(1 for p in g if p["Acao"].startswith("PENDENTE"))))
    print()
    if pend:
        print("PENDENTES (ficam onde estao)")
        for p in pend:
            print("  %-44s %s" % (p["Arquivo"][:44], p["Acao"]))
        print()

    # ------------------------------------------------------------ relatorio ----
    cab = [
        "RELATORIO DE ARQUIVAMENTO - " + datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "Pasta: %s" % pasta,
        "Python %s   plataforma: %s   versao do script: %s"
        % (sys.version.split()[0], sys.platform, VERSAO),
        "Soltos na raiz: %d   ja arquivados nas subpastas: %d" % (len(soltos), ja_arquivados),
        "Notas do dia: %d   pares ticket->nota: %d   so com XML: %d"
        % (len(ordem_notas), len(t2n), len(so_xml)),
        "Tickets assinados: %d lidos do PDF, %d pela ordem   criterio: %s"
        % (lidos_ocr, len(scans_soltos), criterio_scan),
        diag_oc,
        "",
    ]
    colunas = ["Arquivo", "Tipo", "Nota", "Ticket", "Destino", "Acao"]
    larg = [max(len(c), *(len(p[c]) for p in plano)) if plano else len(c) for c in colunas]
    linhas = ["  ".join(c.ljust(w) for c, w in zip(colunas, larg)).rstrip(),
              "  ".join("-" * w for w in larg).rstrip()]
    for p in plano:
        linhas.append("  ".join(p[c].ljust(w) for c, w in zip(colunas, larg)).rstrip())
    if not args.simular:
        escrever_texto(os.path.join(pasta, ARQ_RELATORIO), "\n".join(cab + linhas) + "\n")
        if args.sem_csv:
            print("Relatorio detalhado: %s" % ARQ_RELATORIO)
        else:
            with open(longo(os.path.join(pasta, ARQ_RELATORIO_CSV)), "w",
                      encoding="utf-8-sig", newline="") as fh:
                w = csv.writer(fh, delimiter=";")
                w.writerow(colunas)
                for p in plano:
                    w.writerow([p[c] for c in colunas])
            print("Relatorio detalhado: %s e %s" % (ARQ_RELATORIO, ARQ_RELATORIO_CSV))
        print()

    if args.simular:
        print("--simular: nada foi movido.")
        print("\n".join(cab + linhas))
        return fim(2 if pend else 0)

    if not mover:
        print("Nao ha nada para mover.")
        return fim(2 if pend else 0)

    print("Vou mover %d arquivo(s). Nada foi movido ainda." % len(mover))
    if not args.sim:
        try:
            r = input("Confirmar? (S = sim, qualquer outra tecla cancela): ")
        except (EOFError, KeyboardInterrupt):
            r = ""
        if (r or "").strip().upper() != "S":
            print("Cancelado. Nada foi movido.")
            return fim(0)

    # ------------------------------------------------------------ execucao ----
    # Agrupado por arquivo de origem: quando um documento vai para mais de uma
    # pasta (a OC das duas viagens do mesmo motorista no dia), ele e COPIADO nas
    # primeiras e MOVIDO na ultima - as duas pastas ficam completas e nada sobra
    # solto na raiz.
    porta_origem = {}
    for p in mover:
        porta_origem.setdefault(p["_origem"], []).append(p)

    # O diario do desfazer e regravado a CADA arquivo movido: se a rodada for
    # interrompida no meio (Ctrl+C, botao Parar do painel, queda de rede), o
    # que ja saiu do lugar continua podendo voltar com --desfazer.
    acoes, ok, erros = [], 0, 0
    diario = {"versao": VERSAO, "quando": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
              "pasta": pasta, "acoes": acoes, "pastas_criadas": pastas_criadas}
    alvo_diario = arquivo_estado(pasta, "desfazer")

    def gravar_diario():
        nonlocal alvo_diario
        if not alvo_diario:
            return
        try:
            escrever_json(alvo_diario, diario)
        except OSError as e:
            print("  ATENCAO: nao consegui gravar o diario do desfazer (%s)" % e)
            alvo_diario = None

    gravar_diario()
    for origem, itens in porta_origem.items():
        for i, p in enumerate(itens):
            ultimo = (i == len(itens) - 1)
            alvo = os.path.join(p["_pasta"], p["_nome"])
            try:
                os.makedirs(longo(p["_pasta"]), exist_ok=True)
                if os.path.exists(longo(alvo)):
                    # o shutil.move do Windows sobrescreveria sem avisar
                    raise OSError("ja existe um arquivo com esse nome no destino; nada foi sobrescrito")
                tam = os.path.getsize(longo(origem))
                if ultimo:
                    shutil.move(longo(origem), longo(alvo))
                else:
                    shutil.copy2(longo(origem), longo(alvo))
                # registra ANTES de conferir: se o arquivo saiu do lugar, o
                # desfazer precisa saber, mesmo que a conferencia falhe
                acoes.append({"origem": origem, "destino": alvo,
                              "acao": "mover" if ultimo else "copiar"})
                gravar_diario()
                if not os.path.exists(longo(alvo)) or os.path.getsize(longo(alvo)) != tam:
                    raise OSError("o arquivo nao chegou inteiro ao destino")
                ok += 1
            except OSError as e:
                print("  ERRO em %s: %s" % (p["Arquivo"], e))
                p["Acao"] = "ERRO: %s" % e
                erros += 1

    # relatorio de novo, agora com os erros registrados
    linhas = ["  ".join(c.ljust(w) for c, w in zip(colunas, larg)).rstrip(),
              "  ".join("-" * w for w in larg).rstrip()]
    for p in plano:
        linhas.append("  ".join(p[c].ljust(w) for c, w in zip(colunas, larg)).rstrip())
    escrever_texto(os.path.join(pasta, ARQ_RELATORIO), "\n".join(cab + linhas) + "\n")

    print()
    print("Arquivados: %d   Erros: %d   Pendentes: %d   Ja estavam la: %d"
          % (ok, erros, len(pend), len(ja)))
    if ok:
        print("Para voltar atras: python \"%s\" --desfazer" % os.path.basename(__file__))
    print()
    return fim(2 if (pend or erros) else 0)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nInterrompido.")
        sys.exit(1)
