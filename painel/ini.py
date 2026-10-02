"""Edicao de arquivos .ini sem perder comentarios, ordem nem acentos.

O configparser do Python apaga os comentarios quando grava. Aqui o arquivo e
lido linha a linha e, ao salvar, so a parte depois do "=" das chaves
alteradas e trocada.
"""

import re

LINHA_CHAVE = re.compile(r"^(\s*)([^;#\[\s=][^=]*?)(\s*=\s*)(.*?)\s*$")
LINHA_SECAO = re.compile(r"^\s*\[([^\]]+)\]\s*$")
ENFEITE = re.compile(r"^[\s=\-_*#.]*$")


class Campo:
    def __init__(self, secao, chave, valor, indice, ajuda):
        self.secao = secao
        self.chave = chave
        self.valor = valor
        self.original = valor
        self.indice = indice
        self.ajuda = ajuda


class DocumentoIni:
    def __init__(self, caminho):
        self.caminho = caminho
        bruto = open(caminho, "rb").read()
        self.bom = bruto.startswith(b"\xef\xbb\xbf")
        if self.bom:
            bruto = bruto[3:]
        try:
            texto = bruto.decode("utf-8")
            self.codificacao = "utf-8"
        except UnicodeDecodeError:
            texto = bruto.decode("cp1252")
            self.codificacao = "cp1252"
        self.quebra = "\r\n" if "\r\n" in texto else "\n"
        self.final_com_quebra = texto.endswith(("\n", "\r\n"))
        self.linhas = texto.splitlines()
        self.campos = []
        self._ler()

    def _ler(self):
        secao = ""
        ajuda = []
        for i, linha in enumerate(self.linhas):
            s = linha.strip()
            m = LINHA_SECAO.match(linha)
            if m:
                secao = m.group(1).strip()
                ajuda = []
                continue
            if not s:
                ajuda = []
                continue
            if s[0] in ";#":
                texto = s.lstrip(";#").strip()
                if texto and not ENFEITE.match(texto):
                    ajuda.append(texto)
                continue
            m = LINHA_CHAVE.match(linha)
            if m:
                self.campos.append(Campo(secao, m.group(2).strip(), m.group(4), i, " ".join(ajuda)))
            ajuda = []

    def secoes(self):
        vistas = []
        for c in self.campos:
            if c.secao not in vistas:
                vistas.append(c.secao)
        return vistas

    def alterados(self):
        return [c for c in self.campos if c.valor != c.original]

    def salvar(self):
        for c in self.alterados():
            m = LINHA_CHAVE.match(self.linhas[c.indice])
            self.linhas[c.indice] = m.group(1) + m.group(2) + m.group(3) + c.valor.strip()
        texto = self.quebra.join(self.linhas)
        if self.final_com_quebra:
            texto += self.quebra
        dados = texto.encode(self.codificacao)
        if self.bom:
            dados = b"\xef\xbb\xbf" + dados
        with open(self.caminho, "wb") as f:
            f.write(dados)
        for c in self.campos:
            c.valor = c.valor.strip()
            c.original = c.valor
