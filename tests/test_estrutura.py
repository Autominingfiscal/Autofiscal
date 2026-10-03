"""Regras da estrutura, valendo tambem para as ferramentas que ainda vao chegar.

1. O que esta na pasta comum nao e copiado de novo dentro de uma ferramenta.
2. Toda ferramenta que usa a pasta comum tem o bloco que acha a pasta comum.
3. Toda ferramenta, exportada, roda sozinha - sem o resto do Autofiscal.
4. A Manutencao cria ferramentas novas que ja aparecem no painel.
"""

import ast
import contextlib
import io
import json
import os
import shutil
import subprocess
import sys
import unittest
from unittest import mock

import _apoio

import manutencao
from painel.ferramentas import carregar_todas

COMUM = os.path.join(_apoio.RAIZ, "comum")
MARCA_BLOCO = 'pasta "comum" do Autofiscal'
FORA = {"comum", "tests", "distribuicao", ".git", "__pycache__", "Claude outputs"}


def arquivos_py_das_ferramentas():
    for raiz, pastas, arquivos in os.walk(_apoio.RAIZ):
        pastas[:] = [p for p in pastas if p not in FORA and not p.startswith(".")]
        for nome in arquivos:
            if nome.endswith((".py", ".pyw")):
                yield os.path.join(raiz, nome)


def funcoes_da_comum():
    nomes = set()
    for nome in os.listdir(COMUM):
        if nome.endswith(".py") and not nome.startswith("_"):
            with open(os.path.join(COMUM, nome), encoding="utf-8") as f:
                arvore = ast.parse(f.read())
            nomes |= {n.name for n in arvore.body
                      if isinstance(n, ast.FunctionDef) and not n.name.startswith("_")}
    return nomes


def funcoes_definidas(caminho):
    with open(caminho, encoding="utf-8") as f:
        arvore = ast.parse(f.read())
    return {n.name for n in ast.walk(arvore) if isinstance(n, ast.FunctionDef)}


class TestNadaRepetido(unittest.TestCase):
    # nomes iguais que fazem outra coisa, de proposito (nao sao copia da comum)
    PERMITIDOS = {
        # a ferramenta guarda um atalho com o mesmo nome que recebe o cfg dela
        ("Lançador/etapa3_lancador.py", "peso_alvo"),
        ("Dashboard Local/painel_local.py", "peso_alvo"),
    }

    def test_nenhuma_ferramenta_redefine_funcao_da_comum(self):
        da_comum = funcoes_da_comum()
        self.assertIn("caminho_do_usuario", da_comum)
        achados = []
        for caminho in arquivos_py_das_ferramentas():
            rel = os.path.relpath(caminho, _apoio.RAIZ).replace("\\", "/")
            for nome in funcoes_definidas(caminho) & da_comum:
                if (rel, nome) not in self.PERMITIDOS:
                    achados.append(f"{rel}: def {nome}")
        self.assertEqual(achados, [], "Use a funcao da pasta comum em vez de copiar:\n"
                         + "\n".join(achados))


class TestBlocoDaComum(unittest.TestCase):
    def test_quem_usa_comum_tem_o_bloco(self):
        faltando = []
        for f in carregar_todas(_apoio.RAIZ):
            codigo_da_pasta = ""
            for raiz, pastas, arquivos in os.walk(f.pasta):
                pastas[:] = [p for p in pastas if p != "__pycache__"]
                for nome in arquivos:
                    if nome.endswith(".py"):
                        with open(os.path.join(raiz, nome), encoding="utf-8") as arq:
                            codigo_da_pasta += arq.read()
            if "comum" not in codigo_da_pasta:
                continue
            for a in f.acoes:
                with open(a.caminho_script, encoding="utf-8") as arq:
                    if MARCA_BLOCO not in arq.read() and "Manutencao" not in f.pasta:
                        faltando.append(os.path.relpath(a.caminho_script, _apoio.RAIZ))
        self.assertEqual(sorted(set(faltando)), [])


class TestExportar(unittest.TestCase):
    """Cada ferramenta exportada abre sem a pasta Autofiscal por perto."""

    def test_cada_ferramenta_roda_sozinha(self):
        destino = _apoio.pasta_temporaria(self)
        ambiente = {k: v for k, v in os.environ.items() if k != "PYTHONPATH"}
        for f in carregar_todas(_apoio.RAIZ):
            if f.id == "Manutencao":
                continue
            with contextlib.redirect_stdout(io.StringIO()):
                res = manutencao.exportar(f.pasta, destino)
            self.assertTrue(os.path.isfile(os.path.join(res["pasta"], "comum", "caminhos.py")))
            for script in sorted({a.script for a in f.acoes}):
                caminho = os.path.join(res["pasta"], script)
                # carrega o script como modulo (sem rodar o main) num Python novo
                cmd = [sys.executable, "-c",
                       "import runpy, sys; runpy.run_path(sys.argv[1], run_name='teste')",
                       caminho]
                r = subprocess.run(cmd, cwd=res["pasta"], env=ambiente,
                                   capture_output=True, text=True, timeout=60)
                self.assertEqual(r.returncode, 0, f"{f.id}/{script}:\n{r.stdout}\n{r.stderr}")

    def test_atualizar_mantem_configuracao_de_la(self):
        destino = _apoio.pasta_temporaria(self)
        ferramenta = os.path.join(_apoio.RAIZ, "Automacao planilhas")
        with contextlib.redirect_stdout(io.StringIO()):
            res = manutencao.exportar(ferramenta, destino)
        ini = os.path.join(res["pasta"], "config.ini")
        with open(ini, "w", encoding="utf-8") as f:
            f.write("; configuracao do outro PC\n")
        with open(os.path.join(res["pasta"], "Backup.xlsx"), "w") as f:
            f.write("dado")
        with contextlib.redirect_stdout(io.StringIO()):
            res2 = manutencao.exportar(ferramenta, destino)
        with open(ini, encoding="utf-8") as f:
            self.assertEqual(f.read(), "; configuracao do outro PC\n")
        self.assertIn("config.ini", res2["config_mantida"])
        self.assertTrue(os.path.isfile(os.path.join(res2["pasta"], "automining", "pdf.py")))

    def test_nao_exporta_dados_nem_pasta_que_nao_e_ferramenta(self):
        destino = _apoio.pasta_temporaria(self)
        with self.assertRaises(ValueError):
            manutencao.exportar(COMUM, destino)
        origem = os.path.join(_apoio.pasta_temporaria(self), "Teste")
        os.makedirs(os.path.join(origem, "Backup"))
        for nome in ("ferramenta.json", "x.py", "planilha.xlsx", "nota.pdf", "Backup/b.py"):
            with open(os.path.join(origem, nome), "w") as f:
                f.write("{}")
        with contextlib.redirect_stdout(io.StringIO()):
            res = manutencao.exportar(origem, destino)
        copiados = sorted(os.listdir(res["pasta"]))
        self.assertEqual(copiados, ["comum", "ferramenta.json", "x.py"])


class TestNovaFerramenta(unittest.TestCase):
    def test_cria_a_partir_do_modelo_e_aparece_no_painel(self):
        raiz = _apoio.pasta_temporaria(self)
        for pasta in ("_modelo_ferramenta", "comum"):
            shutil.copytree(os.path.join(_apoio.RAIZ, pasta), os.path.join(raiz, pasta),
                            ignore=shutil.ignore_patterns("__pycache__"))
        with mock.patch.object(manutencao, "RAIZ", raiz), \
                mock.patch.object(manutencao, "MODELO", os.path.join(raiz, "_modelo_ferramenta")), \
                contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(manutencao.nova("Conferência de Fretes"), 0)
            self.assertEqual(manutencao.nova("Conferência de Fretes"), 1)     # ja existe
            self.assertEqual(manutencao.nova("_escondida"), 1)
            self.assertEqual(manutencao.nova('Nome/com:proibidos?'), 0)
        nomes = {f.nome: f for f in carregar_todas(raiz)}
        self.assertIn("Conferência de Fretes", nomes)
        self.assertIn("Nome/com:proibidos?", nomes)
        nova = nomes["Conferência de Fretes"]
        self.assertEqual(nova.erro, "")
        with open(os.path.join(nova.pasta, "ferramenta.json"), encoding="utf-8") as f:
            self.assertEqual(json.load(f)["ordem"], 10)
        # o exemplo do modelo ja usa a pasta comum e roda
        r = subprocess.run([sys.executable, os.path.join(nova.pasta, "exemplo.py"), raiz],
                           input="R$ 1.234,56\n", capture_output=True, text=True, timeout=60)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("Valor lido: 1234.56", r.stdout)


if __name__ == "__main__":
    unittest.main()
