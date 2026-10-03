"""Pasta comum: o codigo que todas as ferramentas usam."""

import os
import unittest
import zlib
from unittest import mock

import _apoio

from comum import arquivos, nfe, numeros, pdf, pesagem, texto
from comum.caminhos import caminho_do_usuario, existe


class TestNumeros(unittest.TestCase):
    def test_texto_brasileiro(self):
        casos = {
            "R$ 1.234,56": 1234.56, "14,80": 14.8, "1.234": 1234.0, "14.80": 14.8,
            "(10,00)": -10.0, "-5,5": -5.5, "1 234,5": 1234.5, "\u00a0R$\u00a07,00": 7.0,
            "1.234.567": 1234567.0,
        }
        for entrada, esperado in casos.items():
            self.assertEqual(numeros.numero_br(entrada), esperado, entrada)

    def test_numero_de_verdade(self):
        self.assertEqual(numeros.numero_br(7), 7.0)
        self.assertEqual(numeros.numero_br(2.5), 2.5)
        self.assertIsNone(numeros.numero_br(float("nan")))

    def test_o_que_nao_e_numero(self):
        for v in (None, True, False, "", " - ", "--", "=A1*2", "#REF!", "abc", "nan", "1,2,3"):
            self.assertIsNone(numeros.numero_br(v), repr(v))


class TestTexto(unittest.TestCase):
    def test_norm(self):
        self.assertEqual(texto.norm("Placa do\nCavalo"), "PLACADOCAVALO")
        self.assertEqual(texto.norm("PESO_SAÍDA"), "PESOSAIDA")
        self.assertEqual(texto.norm("ABC-1D23"), "ABC1D23")
        self.assertEqual(texto.norm(None), "")

    def test_sem_acento(self):
        self.assertEqual(texto.sem_acento("Líquido Ção"), "Liquido Cao")


class TestNfe(unittest.TestCase):
    def test_chave(self):
        chave = _apoio.chave_valida(19007)
        self.assertTrue(nfe.dv_confere(chave))
        self.assertFalse(nfe.dv_confere(chave[:-1] + str((int(chave[-1]) + 1) % 10)))
        self.assertFalse(nfe.dv_confere(chave[:43]))
        self.assertFalse(nfe.dv_confere("x" * 44))
        self.assertEqual(nfe.nota_da_chave(chave), 19007)
        self.assertEqual(nfe.modelo_da_chave(chave), "55")
        self.assertEqual(nfe.formatar(chave).replace(" ", ""), chave)
        self.assertEqual(nfe.so_digitos("26.26 09-0"), "2626090")


class TestPesagem(unittest.TestCase):
    ALVOS = {"TRUCADO": 48500, "LS 4 EIXOS": 58500, "VANDERLEIA": 53000}

    def test_peso_alvo(self):
        self.assertEqual(pesagem.peso_alvo(self.ALVOS, "ls-4 eixos"), 58500)
        self.assertEqual(pesagem.peso_alvo(self.ALVOS, "VANDERLÉIA"), 53000)
        self.assertEqual(pesagem.peso_alvo(self.ALVOS, "CARRETA LS 4 EIXOS"), 58500)
        self.assertIsNone(pesagem.peso_alvo(self.ALVOS, "BITREM"))
        self.assertIsNone(pesagem.peso_alvo(self.ALVOS, None))

    def test_faixas(self):
        f = pesagem.faixa_de_aceite
        self.assertEqual(f(58500 + 81, 58500, 80, 1000), pesagem.VERMELHO)
        self.assertEqual(f(58500 + 80, 58500, 80, 1000), pesagem.AMARELO)
        self.assertEqual(f(58500, 58500, 80, 1000), pesagem.DENTRO)
        self.assertEqual(f(58500 - 1000, 58500, 80, 1000), pesagem.DENTRO)
        self.assertEqual(f(58500 - 1001, 58500, 80, 1000), pesagem.VERDE)


class TestPdf(unittest.TestCase):
    def test_desescapar(self):
        self.assertEqual(pdf.desescapar(rb"a\(b\)\\c\101\n\9"), b"a(b)\\cA\n9")
        self.assertEqual(pdf.desescapar(b"quebra\\\r\ncontinua"), b"quebracontinua")
        self.assertEqual(pdf.desescapar(b"fim\\"), b"fim\\")

    def test_streams_e_conteudos(self):
        cru = b"BT (cru) Tj ET"
        comp = zlib.compress(b"BT (comprimido) Tj ET")
        dados = (b"1 0 obj << >> stream\n" + cru + b"\nendstream endobj "
                 b"2 0 obj << /Filter /FlateDecode >> stream\r\n" + comp + b"endstream")
        self.assertEqual(len(list(pdf.streams(dados))), 2)
        textos = b" ".join(pdf.conteudos(dados))
        self.assertIn(b"(cru)", textos)
        self.assertIn(b"(comprimido)", textos)

    def test_descompactar_lixo(self):
        self.assertEqual(pdf.descompactar(b"\xff\xd8 jpeg"), b"")
        self.assertEqual(pdf.descompactar(b""), b"")

    def test_ler_bytes_sem_arquivo(self):
        self.assertEqual(pdf.ler_bytes(os.path.join(_apoio.pasta_temporaria(self), "x.pdf")), b"")


class TestArquivos(unittest.TestCase):
    def test_utf8_bom_e_ansi(self):
        pasta = _apoio.pasta_temporaria(self)
        for nome, bruto in (("a.ini", "\ufeff[g]\nnome = AÇÃO\n".encode("utf-8")),
                            ("b.ini", "[g]\nnome = AÇÃO\n".encode("cp1252"))):
            caminho = os.path.join(pasta, nome)
            with open(caminho, "wb") as f:
                f.write(bruto)
            self.assertEqual(arquivos.ler_ini(caminho)["g"]["nome"], "AÇÃO", nome)

    def test_percentual_e_comentario_no_fim(self):
        caminho = os.path.join(_apoio.pasta_temporaria(self), "c.ini")
        with open(caminho, "w", encoding="utf-8") as f:
            f.write("[g]\npasta = %OneDrive%\\x ; comentario\n")
        cp = arquivos.ler_ini(caminho, inline_comment_prefixes=(";",))
        self.assertEqual(cp["g"]["pasta"], "%OneDrive%\\x")


class TestCaminhos(unittest.TestCase):
    def test_variavel_relativo_e_vazio(self):
        pasta = _apoio.pasta_temporaria(self)
        with mock.patch.dict(os.environ, {"MINHAPASTA": pasta}):
            self.assertEqual(caminho_do_usuario('"%MINHAPASTA%"'), pasta)
        self.assertEqual(caminho_do_usuario("sub\\x.ini", base=pasta),
                         os.path.normpath(os.path.join(pasta, "sub", "x.ini")))
        self.assertEqual(caminho_do_usuario(""), "")

    def test_pasta_de_outro_usuario_vira_a_deste(self):
        casa = _apoio.pasta_temporaria(self)
        onedrive = os.path.join(casa, "OneDrive - Empresa")
        os.makedirs(os.path.join(onedrive, "Docs"))
        open(os.path.join(onedrive, "Docs", "EXPED_7.xlsx"), "w").close()
        with mock.patch.dict(os.environ, {"USERPROFILE": casa, "HOME": casa, "OneDrive": onedrive}):
            self.assertEqual(caminho_do_usuario(r"C:\Users\fulano\OneDrive\Docs"),
                             os.path.join(onedrive, "Docs"))
            achado = caminho_do_usuario(r"C:\Users\fulano\OneDrive\Docs\EXPED_*.xlsx")
            self.assertTrue(existe(achado), achado)

    def test_existe_com_coringa(self):
        pasta = _apoio.pasta_temporaria(self)
        open(os.path.join(pasta, "EXPED_7.xlsx"), "w").close()
        self.assertTrue(existe(os.path.join(pasta, "EXPED_*.xlsx")))
        self.assertFalse(existe(os.path.join(pasta, "NADA_*.xlsx")))
        self.assertFalse(existe(""))


if __name__ == "__main__":
    unittest.main()
