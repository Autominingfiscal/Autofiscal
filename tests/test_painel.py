"""Painel: ferramenta.json, marcadores, edicao do .ini e caminhos."""

import os
import unittest
from unittest import mock

import _apoio

from painel.caminhos import caminho_do_usuario
from painel.ferramentas import Acao, Ferramenta, carregar_todas
from painel.ini import DocumentoIni


class TestFerramentasDoProjeto(unittest.TestCase):
    """Confere os ferramenta.json de verdade: um erro aqui some o botao do painel."""

    def test_todas_carregam_e_os_scripts_existem(self):
        ferramentas = carregar_todas(_apoio.RAIZ)
        self.assertGreater(len(ferramentas), 0)
        for f in ferramentas:
            self.assertEqual(f.erro, "", f.id)
            self.assertTrue(f.acoes, f.id)
            for a in f.acoes:
                self.assertTrue(os.path.isfile(a.caminho_script), a.caminho_script)
            if f.ajuda:
                self.assertTrue(os.path.isfile(f.caminho_ajuda), f.caminho_ajuda)


class TestMarcadores(unittest.TestCase):
    def acao(self, args):
        f = Ferramenta.__new__(Ferramenta)
        f.pasta = "."
        return Acao({"args": args}, f)

    def test_troca_e_grupo_opcional(self):
        a = self.acao(["{pasta: Escolha a pasta}", ["--faixa", "{texto?:Faixa}"], "--sim"])
        m = a.marcadores()
        self.assertEqual([x[1:3] for x in m], [("pasta", False), ("texto", True)])
        self.assertEqual(a.montar_args({m[0][0]: "C:\\x", m[1][0]: ""}), ["C:\\x", "--sim"])
        self.assertEqual(a.montar_args({m[0][0]: "C:\\x", m[1][0]: "1-5"}),
                         ["C:\\x", "--faixa", "1-5", "--sim"])

    def test_filtro_de_arquivo(self):
        m = self.acao(["{arquivo:PDF da nota|*.pdf *.PDF}"]).marcadores()
        self.assertEqual(m[0][1:], ("arquivo", False, "PDF da nota", "*.pdf *.PDF"))


class TestDocumentoIni(unittest.TestCase):
    def test_salva_so_o_alterado_e_mantem_o_resto(self):
        caminho = os.path.join(_apoio.pasta_temporaria(self), "c.ini")
        original = ("; comentário com acento\r\n[geral]\r\n; ajuda da pasta\r\n"
                    "pasta = C:\\velha\r\nativo = SIM\r\n")
        with open(caminho, "wb") as f:
            f.write(b"\xef\xbb\xbf" + original.encode("utf-8"))
        doc = DocumentoIni(caminho)
        campo = next(c for c in doc.campos if c.chave == "pasta")
        self.assertEqual(campo.ajuda, "ajuda da pasta")
        campo.valor = "C:\\nova"
        doc.salvar()
        with open(caminho, "rb") as f:
            gravado = f.read()
        self.assertEqual(gravado, b"\xef\xbb\xbf" + original.replace(
            "C:\\velha", "C:\\nova").encode("utf-8"))

    def test_arquivo_em_cp1252(self):
        caminho = os.path.join(_apoio.pasta_temporaria(self), "c.ini")
        with open(caminho, "wb") as f:
            f.write("[a]\nnome = SERVIÇOS\n".encode("cp1252"))
        doc = DocumentoIni(caminho)
        self.assertEqual(doc.codificacao, "cp1252")
        self.assertEqual(doc.campos[0].valor, "SERVIÇOS")


class TestCaminhoDoUsuario(unittest.TestCase):
    def test_variavel_e_relativo(self):
        pasta = _apoio.pasta_temporaria(self)
        with mock.patch.dict(os.environ, {"MINHAPASTA": pasta}):
            self.assertEqual(caminho_do_usuario("%MINHAPASTA%"), pasta)
        self.assertEqual(caminho_do_usuario("sub\\x.ini", base=pasta),
                         os.path.normpath(os.path.join(pasta, "sub", "x.ini")))

    def test_pasta_de_outro_usuario_vira_a_deste(self):
        casa = _apoio.pasta_temporaria(self)
        os.makedirs(os.path.join(casa, "OneDrive - Empresa", "Docs"))
        with mock.patch.dict(os.environ, {"USERPROFILE": casa, "HOME": casa,
                                          "OneDrive": os.path.join(casa, "OneDrive - Empresa")}):
            self.assertEqual(caminho_do_usuario(r"C:\Users\fulano\OneDrive\Docs"),
                             os.path.join(casa, "OneDrive - Empresa", "Docs"))

    def test_vazio(self):
        self.assertEqual(caminho_do_usuario(""), "")


if __name__ == "__main__":
    unittest.main()


class TestInstalado(unittest.TestCase):
    """Instalado (.exe) nao existe python.exe: o executor roda os scripts."""

    def test_pedido_de_script(self):
        from painel.executor import pedido_de_script
        self.assertEqual(pedido_de_script(["a.exe", "--rodar-script", "x.py", "1", "2"]),
                         ("x.py", ["1", "2"]))
        self.assertIsNone(pedido_de_script(["a.exe"]))
        self.assertIsNone(pedido_de_script(["a.exe", "--outra", "x.py"]))

    def test_comando_pelo_codigo_fonte(self):
        from painel.execucao import montar_comando
        with mock.patch("painel.execucao.instalado", return_value=False):
            cmd = montar_comando("s.py", ["a"])
        self.assertEqual(cmd[1:], ["-u", "s.py", "a"])

    def test_comando_instalado_usa_o_executor(self):
        from painel.execucao import montar_comando
        with mock.patch("painel.execucao.instalado", return_value=True), \
                mock.patch("painel.execucao.caminho_executor", return_value="E.exe"):
            self.assertEqual(montar_comando("s.py", ["a"]), ["E.exe", "--rodar-script", "s.py", "a"])
            with self.assertRaises(OSError):
                montar_comando("s.py", [], comando=["-m", "pip"])

    def test_rodar_script_devolve_o_codigo_de_saida(self):
        from painel.executor import rodar_script
        pasta = _apoio.pasta_temporaria(self)
        script = os.path.join(pasta, "s.py")
        with open(script, "w", encoding="utf-8") as f:
            f.write("import sys\nassert sys.argv[1:] == ['x']\nsys.exit(2)\n")
        with mock.patch("sys.argv", []), mock.patch("sys.path", list(os.sys.path)):
            self.assertEqual(rodar_script(script, ["x"]), 2)
