"""Dashboard Fluig: regras do .ini, leitura da exportacao e o HTML gerado."""

import datetime as dt
import json
import os
import re
import time
import unittest
from unittest import mock

import _apoio

import openpyxl

import dashboard_fluig as df

CFG = df.ler_config()
CAB = list(df.COLUNAS_PADRAO.values())


def planilha(pasta, linhas, nome="Resultado da consulta de solicitações.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(CAB)
    for l in linhas:
        ws.append([l.get(campo) for campo in df.COLUNAS_PADRAO])
    caminho = os.path.join(pasta, nome)
    wb.save(caminho)
    return caminho


class TestRegras(unittest.TestCase):
    def test_categorias_pelo_comeco_da_palavra(self):
        casos = {
            ("Serviço Fornecimento Energia (ONS)", "NF-e", ""): "Fornecimento de energia (ONS)",
            ("Serviço de construção civil", "NFS-e", ""): "Serviço",      # CONSTRUCAO nao e ONS
            ("CTE Rodogranel", "CT-e", ""): "CTE Rodogranel",
            ("", "CT-e", ""): "CTE Rodogranel",
            ("Produtos Gerais", "NF-e", ""): "Produtos gerais",
            ("Serviço", "NFS-e", ""): "Serviço",
            ("Locação", "", ""): "Outros",
        }
        for textos, esperado in casos.items():
            self.assertEqual(df.categoria_de(textos, CFG), esperado, textos)

    def test_situacao_e_reprovada(self):
        self.assertEqual(df.situacao_de("Finalizada", CFG), "F")
        self.assertEqual(df.situacao_de("CANCELADA", CFG), "C")
        self.assertEqual(df.situacao_de("Aberta - em andamento", CFG), "A")
        self.assertEqual(df.situacao_de("Suspensa", CFG), "?")
        for v in ("Reprovado", "reprovada", "Não", "N", "false"):
            self.assertTrue(df.reprovada(v, CFG), v)
        for v in ("Aprovado", "Sim", "", None, "Nota"):
            self.assertFalse(df.reprovada(v, CFG), v)

    def test_datas_e_valores(self):
        self.assertEqual(df.para_data("02/10/2026 14:05:09"), dt.datetime(2026, 10, 2, 14, 5, 9))
        self.assertEqual(df.para_data("2/1/2026"), dt.datetime(2026, 1, 2))
        self.assertEqual(df.para_data("2026-10-02T14:05"), dt.datetime(2026, 10, 2, 14, 5))
        self.assertEqual(df.para_data(dt.datetime(2026, 1, 1, 8)), dt.datetime(2026, 1, 1, 8))
        self.assertIsNone(df.para_data("31/02/2026"))
        self.assertIsNone(df.para_data(""))
        self.assertEqual(df.para_valor("R$ 1.234,56"), 1234.56)
        self.assertEqual(df.para_valor(10), 10.0)
        self.assertIsNone(df.para_valor(None))


class TestLeitura(unittest.TestCase):
    def test_registros_sla_e_linha_repetida(self):
        pasta = _apoio.pasta_temporaria(self)
        caminho = planilha(pasta, [
            {"solicitacao": 101, "situacao": "Finalizada", "inicio": "01/10/2026 08:00",
             "fim": "03/10/2026 08:00", "tecnico": "Ana", "aprovacao": "Aprovado",
             "fornecedor": "Fornecedor A", "valor": "R$ 1.000,00", "tipo_nota": "Serviço"},
            {"solicitacao": 102, "situacao": "Aberta", "inicio": "05/10/2026 09:00",
             "cnpj": "00.000.000/0001-00", "valor": 50, "tipo_documento": "CT-e"},
            {"solicitacao": 102, "tecnico": "Bruno"},                  # repete e completa
            {"solicitacao": 103, "situacao": "Finalizada"},            # sem data: fica de fora
        ])
        cab, linhas = df.ler_linhas(caminho)
        regs, avisos, diag = df.montar_registros(cab, linhas, CFG)
        self.assertEqual([r["n"] for r in regs], ["101", "102"])
        r1, r2 = regs
        self.assertEqual((r1["s"], r1["h"], r1["f"], r1["v"]), ("F", 48.0, "2026-10-03", 1000.0))
        self.assertEqual((r2["s"], r2["h"], r2["f"], r2["t"]), ("A", None, None, "Bruno"))
        self.assertEqual(r2["fo"], "00.000.000/0001-00")              # sem razao social: o CNPJ
        self.assertEqual(r2["k"], "CTE Rodogranel")
        self.assertTrue(any("repetiam" in a for a in avisos))
        self.assertTrue(any("sem data de inicio" in a for a in avisos))
        self.assertEqual(diag["situacao"]["Finalizada"], 2)

    def test_coluna_obrigatoria_faltando(self):
        with self.assertRaises(ValueError):
            df.montar_registros(["Outra coluna"], [("x",)], CFG)

    def test_planilha_mais_recente(self):
        pasta = _apoio.pasta_temporaria(self)
        velha = planilha(pasta, [], "Resultado da consulta de solicitações (1).xlsx")
        nova = planilha(pasta, [], "Resultado da consulta de solicitações (2).xlsx")
        agora = time.time()
        os.utime(velha, (agora - 100, agora - 100))
        os.utime(nova, (agora, agora))
        self.assertEqual(df.achar_planilha(os.path.join(pasta, "Resultado da consulta*.xlsx")), nova)
        self.assertIsNone(df.achar_planilha(os.path.join(pasta, "nao existe*.xlsx")))


class TestHtml(unittest.TestCase):
    def test_dados_embutidos_e_nome_malicioso_nao_quebra_o_script(self):
        regs = [{"n": "1", "i": "2026-10-01", "f": "2026-10-02", "s": "F", "t": "Ana",
                 "r": 0, "fo": "</script><b>x", "v": 10.0, "k": "Serviço", "h": 24.0}]
        html = df.gerar_html(regs, "x.xlsx", CFG, agora=dt.datetime(2026, 10, 7, 9, 0))
        self.assertNotIn("/*__DADOS__*/", html)
        self.assertEqual(html.count("</script>"), 1)                 # so o fechamento verdadeiro
        bloco = re.search(r"const DADOS = (.*?);\n", html).group(1)
        dados = json.loads(bloco)
        self.assertEqual(dados["fornecedores"], ["</script><b>x"])
        self.assertEqual(dados["hoje"], "2026-10-07")
        self.assertEqual(dados["linhas"][0][:4], ["1", "2026-10-01", "2026-10-02", "F"])

    def test_main_gera_o_arquivo(self):
        pasta = _apoio.pasta_temporaria(self)
        caminho = planilha(pasta, [{"solicitacao": 1, "situacao": "Aberta", "inicio": "01/10/2026"}])
        with mock.patch.object(df, "SAIDA", pasta), \
                mock.patch("sys.stdout"):
            self.assertEqual(df.main([caminho, "--nao-abrir"]), 0)
        self.assertTrue(os.path.isfile(os.path.join(pasta, "Dashboard Fluig.html")))


if __name__ == "__main__":
    unittest.main()
