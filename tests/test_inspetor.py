"""Inspetor de janelas: o que ele conclui a partir do que o Windows enxerga."""

import os
import unittest
from collections import Counter

import _apoio  # noqa: F401

import inspecionar


def contagem(total, digitacao, botoes=0):
    return {"total": total, "digitacao": digitacao, "botoes": botoes, "classes": Counter()}


class TestTipoDePrograma(unittest.TestCase):
    def test_reconhece_janela_remota_e_progress(self):
        self.assertEqual(inspecionar.tipo_de_programa(r"C:\Citrix\wfica32.exe")[0], "remoto")
        self.assertEqual(inspecionar.tipo_de_programa(r"C:\Windows\System32\MSTSC.EXE")[0],
                         "remoto")
        self.assertEqual(inspecionar.tipo_de_programa(r"C:\dlc\bin\prowin32.exe")[0], "progress")
        self.assertEqual(inspecionar.tipo_de_programa(r"C:\x\notepad.exe")[0], "local")
        self.assertEqual(inspecionar.tipo_de_programa("")[0], "local")


class TestConclusao(unittest.TestCase):
    def conclusao(self, tipo, w32, uia):
        resumo = inspecionar.montar_resumo("Tela", "x.exe", tipo, "Citrix",
                                           {"win32": ([], w32), "uia": ([], uia)})
        return "\n".join(resumo)

    def test_campos_visiveis(self):
        self.assertIn("enxerga os campos", self.conclusao("local", contagem(20, 3), contagem(0, 0)))
        self.assertIn("enxerga os campos", self.conclusao("progress", contagem(1, 0), contagem(9, 2)))

    def test_remota_vence_mesmo_com_controles(self):
        self.assertIn("OUTRO computador", self.conclusao("remoto", contagem(5, 1), contagem(5, 1)))

    def test_sem_campos(self):
        self.assertIn("nenhum campo", self.conclusao("local", contagem(12, 0), contagem(8, 0)))
        self.assertIn("nao enxerga nada", self.conclusao("local", contagem(1, 0), contagem(1, 0)))


class TestNaoGravaConteudo(unittest.TestCase):
    """O .txt e feito para ser enviado: o que foi digitado nao pode ir junto."""

    def test_campo_de_digitacao_mostra_so_o_tamanho(self):
        class Campo:
            def class_name(self):
                return "Edit"

            def window_text(self):
                return "CLIENTE SIGILOSO LTDA"

            def is_enabled(self):
                return True

            def control_id(self):
                return 7

            def rectangle(self):
                raise RuntimeError

        n = contagem(0, 0)
        linha = inspecionar._descrever_win32(Campo(), n)
        self.assertNotIn("SIGILOSO", linha)
        self.assertIn("21 caractere(s)", linha)
        self.assertEqual(n["digitacao"], 1)


@unittest.skipUnless(os.name == "nt", "so no Windows")
class TestListar(unittest.TestCase):
    def test_lista_sem_erro(self):
        import contextlib
        import io
        with contextlib.redirect_stdout(io.StringIO()) as saida:
            self.assertEqual(inspecionar.main(["--listar"]), 0)
        self.assertIn("JANELAS ABERTAS", saida.getvalue())


if __name__ == "__main__":
    unittest.main()
