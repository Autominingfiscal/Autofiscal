"""Pecas de baixo nivel para tirar texto de PDF so com o Python.

Cada ferramenta monta a sua leitura em cima destas pecas, porque os
documentos sao diferentes: a chave da DANFE pode estar em qualquer lugar, o
ticket da balanca precisa da POSICAO de cada texto, etc. O que todas repetem -
achar os streams, descompactar, resolver os escapes das strings - fica aqui.
"""

import re
import zlib

_ESCAPES = {ord("n"): b"\n", ord("r"): b"\r", ord("t"): b"\t", ord("b"): b"\b",
            ord("f"): b"\f"}
_OCTAL = re.compile(rb"[0-7]{1,3}")


def ler_bytes(caminho):
    """Conteudo do arquivo, ou b'' se nao der para abrir."""
    try:
        with open(caminho, "rb") as f:
            return f.read()
    except OSError:
        return b""


def streams(dados):
    """O conteudo cru de cada par stream ... endstream do arquivo."""
    pos = 0
    while True:
        ini = dados.find(b"stream", pos)
        if ini < 0:
            return
        if dados[ini - 3:ini] == b"end":         # "endstream" solto: nao abre stream
            pos = ini + len(b"stream")
            continue
        corpo = ini + len(b"stream")
        if dados[corpo:corpo + 2] == b"\r\n":
            corpo += 2
        elif dados[corpo:corpo + 1] in (b"\n", b"\r"):
            corpo += 1
        fim = dados.find(b"endstream", corpo)
        if fim < 0:
            return
        yield dados[corpo:fim]
        pos = fim + len(b"endstream")


def parece_zlib(bruto):
    """O primeiro byte de todo stream FlateDecode diz 'metodo 8 (deflate)'."""
    return bool(bruto) and (bruto[0] & 0x0F) == 8


def descompactar(bruto):
    """Stream FlateDecode aberto; b'' se nao for zlib. Truncado: o que der."""
    if not parece_zlib(bruto):
        return b""
    try:
        return zlib.decompress(bruto)
    except zlib.error:
        pass
    try:
        return zlib.decompressobj().decompress(bruto)
    except zlib.error:
        return b""


def conteudos(dados):
    """Cada stream ja pronto para procurar texto: descompactado quando e zlib,
    cru quando o gerador do PDF nao comprimiu."""
    for bruto in streams(dados):
        yield descompactar(bruto) or bruto


def desescapar(raw):
    """Bytes de uma string literal de PDF, com os escapes resolvidos.

    \\n \\r \\t \\b \\f, \\( \\) \\\\, octal (\\101 -> 'A') e barra no fim da linha
    (continua na linha de baixo). Quem chama decide a codificacao do resultado.
    """
    out, i, n = bytearray(), 0, len(raw)
    while i < n:
        c = raw[i]
        if c != 0x5C or i + 1 >= n:              # 0x5C = barra invertida
            out.append(c)
            i += 1
            continue
        prox = raw[i + 1]
        m = _OCTAL.match(raw, i + 1)
        if m:
            out.append(int(m.group(0), 8) & 0xFF)
            i = m.end()
        elif prox in _ESCAPES:
            out += _ESCAPES[prox]
            i += 2
        elif prox in (0x0D, 0x0A):               # quebra de linha escapada: some
            i += 2
            if prox == 0x0D and raw[i:i + 1] == b"\n":
                i += 1
        else:                                    # \( \) \\ e qualquer outro: o caractere
            out.append(prox)
            i += 2
    return bytes(out)
