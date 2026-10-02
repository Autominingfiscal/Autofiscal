"""Extracao da chave de acesso de dentro do PDF da nota.

Nao depende de biblioteca nenhuma: usa o zlib, que ja vem com o Python.
A logica e a mesma do modulo VBA - descompactar os streams do PDF, juntar o
texto e procurar a sequencia de 44 digitos, gravando so o que passar nas duas
conferencias.
"""

import re
import zlib

ROTULO = "CHAVE DE ACESSO"


# ------------------------------------------------------------ texto do PDF --
def extrair_texto(caminho):
    """Junta o texto dos streams do PDF. Devolve '' se nao der para ler."""
    try:
        with open(caminho, "rb") as f:
            dados = f.read()
    except OSError:
        return ""

    pedacos = []
    for bruto in _streams(dados):
        conteudo = _descompactar(bruto)
        if conteudo:
            pedacos.append(_texto_do_conteudo(conteudo))
    return " ".join(p for p in pedacos if p)


def _streams(dados):
    """O conteudo de cada par stream/endstream do arquivo."""
    pos = 0
    while True:
        ini = dados.find(b"stream", pos)
        if ini < 0:
            return
        corpo = ini + len(b"stream")
        if dados[corpo : corpo + 2] == b"\r\n":
            corpo += 2
        elif dados[corpo : corpo + 1] in (b"\n", b"\r"):
            corpo += 1

        fim = dados.find(b"endstream", corpo)
        if fim < 0:
            return
        # todo stream vai para o zlib, com ou sem /FlateDecode: o que nao for
        # zlib volta vazio em _descompactar
        yield dados[corpo:fim]
        pos = fim + len(b"endstream")


def _descompactar(bruto):
    try:
        return zlib.decompress(bruto)
    except zlib.error:
        pass
    # stream truncado: aproveita o que der
    try:
        d = zlib.decompressobj()
        return d.decompress(bruto)
    except zlib.error:
        return b""


_TEXTO = re.compile(rb"\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]+>")


def _texto_do_conteudo(conteudo):
    """Pega o conteudo dos operadores de texto do stream."""
    if b"Tj" not in conteudo and b"TJ" not in conteudo:
        return ""
    saida = []
    for m in _TEXTO.finditer(conteudo):
        bruto = m.group(0)
        if bruto.startswith(b"<"):
            hexa = re.sub(rb"\s", b"", bruto[1:-1])
            if len(hexa) % 2:
                hexa += b"0"
            try:
                bytes_ = bytes.fromhex(hexa.decode("ascii"))
            except ValueError:
                continue
            # hexadecimal costuma vir em UTF-16BE nos PDFs
            if len(bytes_) >= 2 and bytes_[0] == 0:
                saida.append(bytes_.decode("utf-16-be", "ignore"))
            else:
                saida.append(bytes_.decode("latin-1", "ignore"))
        else:
            saida.append(_literal(bruto[1:-1]))
    return " ".join(saida)


_ESCAPES = {b"n": "\n", b"r": "\r", b"t": "\t", b"b": "\b", b"f": "\f"}


def _literal(bruto):
    saida = []
    i = 0
    while i < len(bruto):
        ch = bruto[i : i + 1]
        if ch == b"\\" and i + 1 < len(bruto):
            prox = bruto[i + 1 : i + 2]
            if prox in _ESCAPES:
                saida.append(_ESCAPES[prox])
                i += 2
                continue
            if prox.isdigit():
                oct_ = bruto[i + 1 : i + 4]
                oct_ = oct_[: len(oct_) - len(oct_.lstrip(b"01234567")) or 3]
                try:
                    saida.append(chr(int(oct_, 8)))
                    i += 1 + len(oct_)
                    continue
                except ValueError:
                    pass
            saida.append(prox.decode("latin-1", "ignore"))
            i += 2
            continue
        saida.append(ch.decode("latin-1", "ignore"))
        i += 1
    return "".join(saida)


# ----------------------------------------------------------- a chave em si --
def dv_confere(chave):
    """Digito verificador da chave da NF-e: modulo 11, pesos 2..9 da direita."""
    soma = 0
    peso = 2
    for ch in reversed(chave[:43]):
        soma += int(ch) * peso
        peso = peso + 1 if peso < 9 else 2
    dv = 11 - (soma % 11)
    if dv >= 10:
        dv = 0
    return dv == int(chave[43])


def nota_da_chave(chave):
    """O numero da nota vem embutido na chave, nas posicoes 26 a 34."""
    return int(chave[25:34])


def aceita(chave, nf_esperada=None):
    """As duas conferencias que autorizam gravar a chave.

    Devolve (ok, motivo).
    """
    if len(chave) != 44 or not chave.isdigit():
        return False, "nao tem 44 digitos"
    if not dv_confere(chave):
        return False, "digito verificador nao confere"
    if nf_esperada:
        dentro = nota_da_chave(chave)
        if dentro != int(nf_esperada):
            return False, f"a nota dentro da chave e {dentro}, esperava {int(nf_esperada)}"
    return True, ""


def formatar(chave):
    """11 grupos de 4 digitos, como sempre foi digitado na planilha."""
    return " ".join(chave[i : i + 4] for i in range(0, 44, 4))


_SEQ = re.compile(r"[\d\s]{44,}")


def _candidatas(texto):
    """Toda sequencia de 44 digitos do documento, tolerando espacos no meio."""
    vistas = set()
    for m in _SEQ.finditer(texto):
        digitos = re.sub(r"\s", "", m.group(0))
        for i in range(len(digitos) - 43):
            cand = digitos[i : i + 44]
            if cand not in vistas:
                vistas.add(cand)
                yield cand


def chave_do_pdf(caminho, nf_esperada=None):
    """Devolve (chave, motivo). chave='' quando nao deu para extrair.

    Caminho normal: logo depois do rotulo "CHAVE DE ACESSO".
    Reserva: qualquer sequencia de 44 digitos que passe nas conferencias.
    """
    texto = extrair_texto(caminho)
    if not texto:
        return "", "nao achei texto no PDF (provavelmente e digitalizado)"

    recusada = None

    pos = texto.upper().find(ROTULO)
    if pos >= 0:
        depois = texto[pos + len(ROTULO) : pos + len(ROTULO) + 200]
        digitos = re.sub(r"[^\d]", "", re.match(r"[^\d]*([\d\s]*)", depois).group(1))
        if len(digitos) >= 44:
            cand = digitos[:44]
            ok, motivo = aceita(cand, nf_esperada)
            if ok:
                return cand, ""
            recusada = (cand, motivo)

    for cand in _candidatas(texto):
        ok, motivo = aceita(cand, nf_esperada)
        if ok:
            return cand, ""
        if recusada is None:
            recusada = (cand, motivo)

    if recusada:
        return "", f"chave encontrada mas recusada ({recusada[1]}): {recusada[0]}"
    return "", "nao achei chave valida no PDF"
