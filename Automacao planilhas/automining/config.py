"""Configuracao em arquivo texto (config.ini), no lugar da aba CONFIG do VBA.

O arquivo fica ao lado dos scripts. E so abrir no Bloco de Notas para mudar
preco do frete, CFOP, ano do embarque ou os caminhos da rede.
"""

import configparser
import os
import re
from pathlib import Path

from .planilha import ErroDeUso, como_numero

NOME = "config.ini"

MODELO = """\
; ---------------------------------------------------------------------------
;  AUTOMACAO MVV - configuracao
;
;  Use SEMPRE o caminho de rede (\\\\servidor\\pasta\\...), nunca a letra de
;  unidade mapeada: letra mapeada muda de maquina para maquina.
; ---------------------------------------------------------------------------

; CAMINHOS PARA MAIS DE UM USUARIO: pode usar %USERPROFILE% ou %OneDrive%
;   (ex.: %OneDrive%\\Documentos\\pasta\\arquivo.xlsx). Caminho com o nome de
;   OUTRO usuario (C:\\Users\\fulano\\...) e ajustado sozinho para quem estiver
;   usando - nao precisa mais editar este arquivo ao trocar de usuario.

[arquivos]
; Planilha de expedicao (EXPED_CONCENTRADO), de onde saem as notas
expedicao = C:\\caminho\\EXPED_CONCENTRADO - 7 embarque.xlsx

; Planilha de remessa, a que tem a aba "Remessas Porto"
remessa = C:\\caminho\\REMESSA FORM. LOTE MVV EMBARQUE 007 2026.xlsx

; Planilha de faturamento
faturamento = C:\\caminho\\FATURAMENTO 2026.xlsx

; Pasta dos PDFs das notas
pdfs = C:\\caminho\\PDFs

; SIM desce em todas as subpastas dos PDFs; NAO le so a pasta indicada
incluir_subpastas = SIM


[remessa]
; As abas da expedicao sao "dd.mm"; o ano vem daqui
ano = 2026
; VL. UNITARIO padrao, usado so nas notas NOVAS. Depois e so digitar o
; valor certo na coluna G da planilha: o script mantem o que estiver la.
valor_unitario = 14.80
cfop = 550404
operacao = R. PORTO
transportadora_padrao = RODOGRANEL LOGÍSTICA E SERVIÇOS LTDA


[faturamento]
uf = AL
cfop = 550404
operacao = R.PORTO
descricao = REMESSA FORMACAO DE LOTE PARA EXPORTACAO
"""


# =============================================================================
# Caminhos que funcionam para qualquer usuario do Windows
# =============================================================================
def caminho_do_usuario(texto, base=None):
    """Ajusta um caminho do config para o usuario que esta rodando o script.

    - aceita variaveis do Windows: %USERPROFILE%, %OneDrive%, %USERNAME%, ~
    - caminho relativo (ex.: ..\\Planilhas\\x.xlsx) vale a partir de `base`
    - caminho de OUTRO usuario (C:\\Users\\fulano\\...) que nao existe aqui e
      trocado pela pasta do usuario atual; a parte "OneDrive..." vira a
      OneDrive dele (mesmo que o nome seja "OneDrive - Empresa").
    Assim ninguem precisa editar o config quando outra pessoa usa o script.
    """
    import glob as _glob

    if not texto:
        return texto
    t = str(texto).strip().strip('"')
    t = re.sub(r"%([^%]+)%", lambda m: os.environ.get(m.group(1), m.group(0)), t)
    t = os.path.expanduser(t)
    if base and not os.path.isabs(t) and not t.startswith(("\\\\", "//")):
        t = os.path.normpath(os.path.join(base, t))

    def existe(c):
        return bool(_glob.glob(c)) if any(x in c for x in "*?") else os.path.exists(c)

    if existe(t):
        return t
    m = re.match(r"^[A-Za-z]:[\\/]+Users[\\/]+[^\\/]+[\\/]*(.*)$", t, re.I)
    if not m:
        return t
    resto = m.group(1)
    candidatos = []
    partes = re.split(r"[\\/]+", resto, maxsplit=1)
    for var in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        od = os.environ.get(var)
        if od and partes[0].lower().startswith("onedrive"):
            candidatos.append(os.path.join(od, partes[1]) if len(partes) > 1 else od)
    candidatos.append(os.path.join(os.path.expanduser("~"), resto))
    for c in candidatos:
        if existe(c):
            return c
    return t


class Config:
    def __init__(self, dados, caminho):
        self._d = dados
        self.caminho = caminho

    # -- arquivos --
    @property
    def expedicao(self):
        return self._arquivo("arquivos", "expedicao")

    @property
    def remessa(self):
        return self._arquivo("arquivos", "remessa")

    @property
    def faturamento(self):
        return self._arquivo("arquivos", "faturamento")

    @property
    def pdfs(self):
        caminho = Path(caminho_do_usuario(
            self._d.get("arquivos", "pdfs", fallback="").strip(), self.caminho.parent))
        if not str(caminho) or not caminho.is_dir():
            raise ErroDeUso(
                f"Nao consegui abrir a pasta dos PDFs:\n  {caminho}\n\n"
                f"Confira a linha 'pdfs' em {self.caminho.name}."
            )
        return caminho

    @property
    def incluir_subpastas(self):
        v = self._d.get("arquivos", "incluir_subpastas", fallback="SIM")
        return v.strip().upper() != "NAO"

    # -- remessa --
    @property
    def ano(self):
        return self._d.getint("remessa", "ano")

    @property
    def valor_unitario(self):
        # aceita 14,80 / 14.80 / R$ 14,80
        bruto = self._d.get("remessa", "valor_unitario", fallback="")
        valor = como_numero(bruto)
        if valor is None:
            raise ErroDeUso(
                f"Nao entendi o valor_unitario '{bruto}' em {self.caminho.name}.\n"
                "Escreva so o numero, por exemplo: valor_unitario = 14,80"
            )
        return valor

    @property
    def cfop_remessa(self):
        return self._d.getfloat("remessa", "cfop")

    @property
    def operacao_remessa(self):
        return self._d.get("remessa", "operacao").strip()

    @property
    def transportadora_padrao(self):
        return self._d.get("remessa", "transportadora_padrao").strip()

    # -- faturamento --
    @property
    def uf(self):
        return self._d.get("faturamento", "uf").strip()

    @property
    def cfop_faturamento(self):
        return self._d.getfloat("faturamento", "cfop")

    @property
    def operacao_faturamento(self):
        return self._d.get("faturamento", "operacao").strip()

    @property
    def descricao(self):
        return self._d.get("faturamento", "descricao").strip()

    # -- apoio --
    def _arquivo(self, secao, chave):
        bruto = self._d.get(secao, chave, fallback="").strip()
        caminho = Path(caminho_do_usuario(bruto, self.caminho.parent)) if bruto else Path()
        if not bruto or not caminho.is_file():
            raise ErroDeUso(
                f"Nao achei a planilha de {chave}:\n  {bruto or '(em branco)'}\n\n"
                f"Confira a linha '{chave}' em {self.caminho.name}."
            )
        return caminho


def carregar(pasta=None):
    pasta = Path(pasta) if pasta else Path(__file__).resolve().parent.parent
    caminho = pasta / NOME
    if not caminho.is_file():
        caminho.write_text(MODELO, encoding="utf-8")
        raise ErroDeUso(
            f"Criei o arquivo {NOME} em:\n  {pasta}\n\n"
            "Abra no Bloco de Notas, aponte os caminhos das planilhas e da "
            "pasta dos PDFs, salve e rode de novo."
        )
    dados = configparser.ConfigParser(inline_comment_prefixes=(";",), interpolation=None)
    dados.read(caminho, encoding="utf-8")
    return Config(dados, caminho)
