"""Autoiss (ISSQN) e Conferencia do planilhao: numeros, nomes e casamento de notas."""

import unittest

import _apoio  # noqa: F401

import conferencia_retencoes as plan
import conferir_issqn as iss


class TestAutoiss(unittest.TestCase):
    def test_valor_num(self):
        self.assertEqual(iss.valor_num("R$ 1.234,56"), (1234.56, True))
        self.assertEqual(iss.valor_num(" - "), (0.0, False))
        self.assertEqual(iss.valor_num("#REF!"), (0.0, False))
        self.assertEqual(iss.valor_num(None), (0.0, False))

    def test_numeros(self):
        self.assertEqual(iss.numero_principal("000123-1"), "123")
        self.assertEqual(iss.numeros_combinam("123", "123"), 2)
        self.assertEqual(iss.numeros_combinam("202600000000123", "123"), 1)
        self.assertEqual(iss.numeros_combinam("4123", "123"), 0)

    def test_nomes(self):
        a = iss.nome_limpo("Acme Serviços Ltda - ME")
        self.assertEqual(a, "ACME SERVICOS")
        self.assertEqual(iss.comparar_nomes(a, iss.nome_limpo("ACME SERVICOS EIRELI")), 2)
        self.assertEqual(iss.comparar_nomes(a, iss.nome_limpo("Outra Empresa SA")), 0)

    def test_casar_prefere_mesmo_numero_e_fornecedor(self):
        nota = {"nf": "123", "razao_n": "ACME SERVICOS", "valor": 50.0}
        conf = [
            {"tit": "123", "forn_n": "OUTRA EMPRESA", "valor": 50.0, "usado": False},
            {"tit": "123", "forn_n": "ACME SERVICOS", "valor": 50.0, "usado": False},
        ]
        melhor, tipo, outro = iss.casar(nota, conf)
        self.assertIs(melhor, conf[1])
        self.assertEqual(tipo, 2)
        self.assertIs(outro, conf[0])


class TestConferenciaPlanilhao(unittest.TestCase):
    def test_para_numero(self):
        self.assertEqual(plan.para_numero("R$ 1.234,56"), 1234.56)
        self.assertEqual(plan.para_numero("(10,00)"), -10.0)
        self.assertEqual(plan.para_numero(" - "), 0.0)
        self.assertEqual(plan.para_numero(12), 12.0)

    def test_chave_nf(self):
        for v in (123, 123.0, "000123", "NF 123", "123/1"):
            self.assertEqual(plan.chave_nf(v), "123", v)
        self.assertIsNone(plan.chave_nf(None))
        self.assertIsNone(plan.chave_nf(float("nan")))


if __name__ == "__main__":
    unittest.main()
