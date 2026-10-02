"""Arquivador: classificacao pelo nome e a rodada completa com desfazer."""

import contextlib
import io
import os
import unittest
from unittest import mock

import _apoio

import arquivar


def rodar(*argv):
    """Roda o main do arquivador sem poluir a saida dos testes."""
    with contextlib.redirect_stdout(io.StringIO()) as saida:
        codigo = arquivar.main(list(argv))
    return codigo, saida.getvalue()


class TestClassificar(unittest.TestCase):
    def test_tipos(self):
        casos = {
            "2010030018964.xml": ("XML", 18964, 0),
            "NF 18964 R. PORTO 19569.pdf": ("NOTA FISCAL", 18964, 19569),
            "PC 19569 REGINALDO.txt": ("PRE-CALCULO", 0, 19569),
            "TICKET 19569.pdf": ("TICKET", 0, 19569),
            "TICKET ASSINADO 18964.pdf": ("TICKET ASSINADO", 18964, 0),
        }
        for nome, (tipo, nota, ticket) in casos.items():
            d = arquivar.classificar(nome)
            self.assertEqual((d.tipo, d.nota, d.ticket), (tipo, nota, ticket), nome)

    def test_ordem_de_carregamento_e_scan(self):
        d = arquivar.classificar("ALYSSON ALVES - OC 17-09-2026.pdf")
        self.assertEqual((d.tipo, d.motorista), ("ORDEM CARREGAMENTO", "ALYSSON ALVES"))
        d = arquivar.classificar("doc12881920260917155023_003.pdf")
        self.assertEqual((d.tipo, d.seq), ("TICKET ASSINADO", 3))

    def test_desconhecido_fica(self):
        self.assertEqual(arquivar.classificar("foto.jpg").tipo, "?")


class TestApoio(unittest.TestCase):
    def test_interpretar_notas(self):
        self.assertEqual(arquivar.interpretar_notas("10-12, 15;11 , 20a21"),
                         [10, 11, 12, 15, 20, 21])

    def test_nome_compativel_respeita_ordem(self):
        self.assertTrue(arquivar.nome_compativel("JOSE SILVA", "JOSÉ FRANCISCO DA SILVA"))
        self.assertFalse(arquivar.nome_compativel("FRANCISCO JOSE", "JOSE FRANCISCO DA SILVA"))

    def test_nome_livre_considera_reservados(self):
        pasta = _apoio.pasta_temporaria(self)
        open(os.path.join(pasta, "x.pdf"), "w").close()
        self.assertEqual(arquivar.nome_livre(pasta, "x.pdf"), "x (2).pdf")
        reservado = {os.path.normcase(os.path.join(pasta, "x (2).pdf"))}
        self.assertEqual(arquivar.nome_livre(pasta, "x.pdf", reservado), "x (3).pdf")


class TestRodada(unittest.TestCase):
    def setUp(self):
        self.pasta = _apoio.pasta_temporaria(self)
        estado = _apoio.pasta_temporaria(self)
        # diario e cache vao para uma pasta temporaria, nao para o AppData real
        p = mock.patch.dict(os.environ, {"LOCALAPPDATA": estado})
        p.start()
        self.addCleanup(p.stop)
        for nome, conteudo in {
            "NF 18964 R. PORTO 19569.pdf": "nf1",
            "2010030018964.xml": "xml1",
            "PC 19569 JOAO.txt": "pc1",
            "TICKET 19569.pdf": "tk1",
            "NF 18965 R. PORTO 19570.pdf": "nf2",
            "sem padrao.docx": "?",
        }.items():
            with open(os.path.join(self.pasta, nome), "w") as f:
                f.write(conteudo)

    def arquivos(self):
        saida = []
        for raiz, _, fs in os.walk(self.pasta):
            for f in fs:
                if not f.startswith("_relatorio"):
                    saida.append(os.path.relpath(os.path.join(raiz, f), self.pasta))
        return sorted(saida)

    def test_simular_nao_move(self):
        antes = self.arquivos()
        codigo, _ = rodar(self.pasta, "--simular", "--sem-faixa", "--sem-pausa")
        self.assertEqual(codigo, 2)          # 2 = tem pendencia (o .docx)
        self.assertEqual(self.arquivos(), antes)

    def test_move_e_desfaz(self):
        antes = self.arquivos()
        codigo, _ = rodar(self.pasta, "--sim", "--sem-faixa", "--sem-pausa")
        self.assertEqual(codigo, 2)
        self.assertEqual(self.arquivos(), sorted([
            os.path.join("NF 18964 R. PORTO", "2010030018964.xml"),
            os.path.join("NF 18964 R. PORTO", "NF 18964 R. PORTO 19569.pdf"),
            os.path.join("NF 18964 R. PORTO", "PC 19569 JOAO.txt"),
            os.path.join("NF 18964 R. PORTO", "TICKET 19569.pdf"),
            os.path.join("NF 18965 R. PORTO", "NF 18965 R. PORTO 19570.pdf"),
            "sem padrao.docx",
        ]))
        codigo, _ = rodar(self.pasta, "--desfazer", "--sem-pausa")
        self.assertEqual(codigo, 0)
        self.assertEqual(self.arquivos(), antes)

    def test_nao_sobrescreve_arquivo_diferente(self):
        destino = os.path.join(self.pasta, "NF 18964 R. PORTO")
        os.makedirs(destino)
        with open(os.path.join(destino, "TICKET 19569.pdf"), "w") as f:
            f.write("outro conteudo")
        rodar(self.pasta, "--sim", "--sem-faixa", "--sem-pausa")
        with open(os.path.join(destino, "TICKET 19569.pdf")) as f:
            self.assertEqual(f.read(), "outro conteudo")
        with open(os.path.join(destino, "TICKET 19569 (2).pdf")) as f:
            self.assertEqual(f.read(), "tk1")


if __name__ == "__main__":
    unittest.main()
