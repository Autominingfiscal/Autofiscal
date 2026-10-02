"""Funcoes repetidas de proposito em mais de uma ferramenta.

Cada ferramenta precisa funcionar sozinha (copiada para outro PC sem o resto
da pasta Autofiscal), entao algumas funcoes existem em mais de um arquivo.
Estes testes rodam os MESMOS casos em todas as copias: se uma for corrigida
e as outras esquecidas, falha aqui.
"""

import datetime as dt
import os
import unittest
from unittest import mock

import _apoio

import conferir_issqn
import etapa3_lancador
import painel_local
from automining import config as automining_config
from painel import caminhos as painel_caminhos

COPIAS_CAMINHO = {
    "painel/caminhos.py": painel_caminhos.caminho_do_usuario,
    "automining/config.py": automining_config.caminho_do_usuario,
    "Lançador/etapa3_lancador.py": etapa3_lancador.caminho_do_usuario,
    "Dashboard Local/painel_local.py": painel_local.caminho_do_usuario,
    "Autoiss/conferir_issqn.py": conferir_issqn.caminho_do_usuario,
}


class TestCaminhoDoUsuarioIgualEmTodas(unittest.TestCase):
    def test_mesmos_resultados(self):
        casa = _apoio.pasta_temporaria(self)
        onedrive = os.path.join(casa, "OneDrive - Empresa")
        os.makedirs(os.path.join(onedrive, "Docs"))
        open(os.path.join(onedrive, "Docs", "EXPED_7.xlsx"), "w").close()
        ambiente = {"USERPROFILE": casa, "HOME": casa, "OneDrive": onedrive,
                    "MINHAPASTA": casa}
        casos = [
            ("", None),
            ("%MINHAPASTA%", None),
            ('"%MINHAPASTA%\\OneDrive - Empresa"', None),
            ("Docs\\EXPED_7.xlsx", onedrive),
            (r"C:\Users\fulano\OneDrive\Docs", None),
            (r"C:\Users\fulano\OneDrive\Docs\EXPED_*.xlsx", None),
            (r"C:\Users\fulano\Documents\nao existe", None),
        ]
        with mock.patch.dict(os.environ, ambiente):
            for texto, base in casos:
                esperado = COPIAS_CAMINHO["painel/caminhos.py"](texto, base)
                for nome, funcao in COPIAS_CAMINHO.items():
                    self.assertEqual(funcao(texto, base), esperado, f"{nome}: {texto!r}")


class TestPesoAlvoIgualNoLancadorENoDashboard(unittest.TestCase):
    """O Dashboard pinta as mesmas faixas que o Lancador: a regra tem que bater."""

    def test_mesmo_alvo(self):
        alvos = {"TRUCADO": 48500.0, "LS 4 EIXOS": 58500.0, "VANDERLEIA": 53000.0}
        cfg_lan = {"alvos": alvos}
        cfg_dash = {"alvos": {painel_local.norm(k): v for k, v in alvos.items()}}
        for modelo in ("TRUCADO", "ls 4 eixos", "LS-4 EIXOS", "VANDERLÉIA", "BITREM", "", None):
            self.assertEqual(etapa3_lancador.peso_alvo(cfg_lan, modelo),
                             painel_local.peso_alvo(cfg_dash, modelo), modelo)

    def test_norm_igual(self):
        for t in ("Nº_NOTA", "PESO_SAÍDA", "Placa do\nCavalo", None):
            self.assertEqual(etapa3_lancador.norm(t), painel_local.norm(t))


class TestHoraIgualNoLancadorENoDashboard(unittest.TestCase):
    def test_hora_gravada_pelo_lancador_e_lida_pelo_dashboard(self):
        momento = dt.datetime(2026, 9, 18, 7, 38, 16)
        self.assertAlmostEqual(painel_local.minutos(etapa3_lancador.hora_excel_s(momento)),
                               7 * 60 + 38 + 16 / 60)


if __name__ == "__main__":
    unittest.main()
