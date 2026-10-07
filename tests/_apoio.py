"""Apoio dos testes: deixa os modulos de cada ferramenta importaveis."""

import os
import sys
import tempfile
import zlib

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

for sub in ("", "Automacao planilhas", "Arquivador", "Lançador", "Autoiss",
            "Conferencia planilhão", "Dashboard Local", "Manutencao",
            "Inspetor de janelas", "Email do faturamento", "Dashboard Fluig"):
    p = os.path.join(RAIZ, sub)
    if p not in sys.path:
        sys.path.insert(0, p)


def pasta_temporaria(caso):
    """Pasta temporaria apagada no fim do teste."""
    d = tempfile.TemporaryDirectory()
    caso.addCleanup(d.cleanup)
    return d.name


def chave_valida(nota=19007):
    """Chave de acesso de 44 digitos (modelo 55) com o digito verificador certo."""
    # cUF+AAMM+CNPJ (20) | modelo 55 | serie | nNF (9) | tpEmis | cNF (8)
    base = "26260900000000000000" + "55" + "001" + "%09d" % nota + "1" + "00000000"
    assert len(base) == 43
    soma, peso = 0, 2
    for ch in reversed(base):
        soma += int(ch) * peso
        peso = peso + 1 if peso < 9 else 2
    dv = 11 - soma % 11
    return base + str(0 if dv >= 10 else dv)


def pdf_com_texto(caminho, *linhas, comprimido=True):
    """Grava um PDF minimo com as linhas de texto (operador Tj)."""
    conteudo = b"BT /F1 10 Tf " + b" ".join(
        b"(" + t.encode("latin-1") + b") Tj" for t in linhas) + b" ET"
    if comprimido:
        conteudo = zlib.compress(conteudo)
        dic = b"<< /Length %d /Filter /FlateDecode >>" % len(conteudo)
    else:
        dic = b"<< /Length %d >>" % len(conteudo)
    with open(caminho, "wb") as f:
        f.write(b"%PDF-1.4\n1 0 obj\n" + dic + b"\nstream\n" + conteudo
                + b"\nendstream\nendobj\n%%EOF\n")
