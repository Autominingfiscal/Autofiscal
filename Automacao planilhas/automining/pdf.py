"""Extracao da chave de acesso de dentro do PDF da nota.

Nao depende de biblioteca nenhuma. A logica e a mesma do modulo VBA -
descompactar os streams do PDF, juntar o texto e procurar a sequencia de 44
digitos, gravando so o que passar nas duas conferencias. As pecas de baixo
nivel (streams, zlib, escapes) e a chave da NF-e moram em comum/.
"""

import re

from comum import pdf as pdfbase
from comum.nfe import dv_confere, formatar, nota_da_chave  # noqa: F401  (usados por chaves.py)

ROTULO = "CHAVE DE ACESSO"


# ------------------------------------------------------------ texto do PDF --
def extrair_texto(caminho):
    """Junta o texto dos streams do PDF. Devolve '' se nao der para ler."""
    pedacos = [_texto_do_conteudo(c) for c in pdfbase.conteudos(pdfbase.ler_bytes(caminho))]
    return " ".join(p for p in pedacos if p)


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
            saida.append(pdfbase.desescapar(bruto[1:-1]).decode("latin-1"))
    return " ".join(saida)


# ----------------------------------------------------------- a chave em si --
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
