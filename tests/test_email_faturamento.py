"""E-mails do faturamento: quais arquivos vao em cada e-mail e como o e-mail e montado."""

import contextlib
import io
import os
import unittest
from unittest import mock

import _apoio

import enviar_emails as em


def criar(pasta, *nomes):
    for nome in nomes:
        caminho = os.path.join(pasta, *nome.split("/"))
        os.makedirs(os.path.dirname(caminho), exist_ok=True)
        with open(caminho, "wb") as f:
            f.write(b"%PDF")


class FakeEmail:
    def __init__(self):
        self.To = self.CC = self.Subject = ""
        self.HTMLBody = "<html><body lang=PT-BR><p>Fulano - Fiscal</p></body></html>"  # assinatura
        self.anexos, self.enviado, self.aberto = [], False, False
        self.Attachments = mock.Mock(Add=self.anexos.append)
        self.GetInspector = object()

    def Send(self):
        self.enviado = True

    def Display(self, modal=False):
        self.aberto = True


class FakeOutlook:
    def __init__(self):
        self.criados = []

    def CreateItem(self, tipo):
        self.criados.append(FakeEmail())
        return self.criados[-1]


CFG = {"xml": {"para": "a@x.com", "cc": ""}, "documentos": {"para": "b@x.com", "cc": "c@x.com"},
       "limite_mb": 20.0}


class TestTipoDoArquivo(unittest.TestCase):
    def test_nomes(self):
        casos = {
            "2610050018964.xml": "XML", "NFe2626-procNFe.XML": "XML",
            "NF 18964 R. PORTO 19569.pdf": "NF", "nf 18964.PDF": "NF",
            "TICKET 19569.pdf": "TICKET", "RELATÓRIO DE PESAGEM.pdf": "RELATORIO",
            "RELATORIO.PDF": "RELATORIO",
            "TICKET ASSINADO 18964.pdf": None,      # o escaneado nao vai
            "PC 19569 JOAO.txt": None, "NF 18964.txt": None, "doc123_001.pdf": None,
        }
        for nome, tipo in casos.items():
            self.assertEqual(em.tipo_do_arquivo(nome), tipo, nome)


class TestPlano(unittest.TestCase):
    def setUp(self):
        self.pasta = _apoio.pasta_temporaria(self)

    def plano(self, cfg=CFG):
        return em.planejar(self.pasta, cfg, "05/10/2026")

    def test_separa_os_anexos_e_ignora_backup(self):
        criar(self.pasta, "1.xml", "Backup/1.xml", "NF 1 R. PORTO/NF 1 R. PORTO 9.pdf",
              "NF 1 R. PORTO/TICKET 9.pdf", "NF 1 R. PORTO/TICKET ASSINADO 1.pdf", "RELATORIO.pdf")
        e1, e2 = self.plano()
        self.assertEqual([os.path.basename(a) for a in e1["anexos"]], ["1.xml"])
        self.assertEqual([os.path.basename(a) for a in e2["anexos"]],
                         ["NF 1 R. PORTO 9.pdf", "TICKET 9.pdf", "RELATORIO.pdf"])
        self.assertEqual((e1["problemas"], e2["problemas"]), ([], []))
        self.assertEqual(e1["assunto"], "FATURAMENTO / RELATÓRIO DE PESAGENS - 05/10/2026")
        self.assertIn("faturamento de hoje (05/10/2026).", e1["texto"])
        self.assertIn("emitidos na data de hoje (05/10/2026).", e2["texto"])

    def test_problemas(self):
        criar(self.pasta, "NF 1.pdf")
        sem_para = dict(CFG, xml={"para": "", "cc": ""})
        e1, e2 = self.plano(sem_para)
        self.assertIn("nenhum XML na pasta", e1["problemas"])
        self.assertIn("ninguem em 'para' no email_config.ini", e1["problemas"])
        self.assertTrue(any("TICKET" in p for p in e2["problemas"]))
        self.assertTrue(any("RELATORIO" in p for p in e2["problemas"]))

    def test_limite_de_tamanho(self):
        criar(self.pasta, "1.xml")
        e1, _ = self.plano(dict(CFG, limite_mb=0.000001))
        self.assertTrue(any("acima do limite" in p for p in e1["problemas"]))


class TestMontarEmail(unittest.TestCase):
    def test_texto_entra_antes_da_assinatura(self):
        outlook = FakeOutlook()
        m = em.montar_email(outlook, CFG["documentos"], "Assunto", "Boa tarde, prezados.\n\nAt.te,",
                            ["C:\\x\\NF 1.pdf"], _apoio.pasta_temporaria(self))
        self.assertEqual((m.To, m.CC, m.Subject), ("b@x.com", "c@x.com", "Assunto"))
        corpo = m.HTMLBody
        self.assertLess(corpo.index("Boa tarde, prezados."), corpo.index("Fulano - Fiscal"))
        self.assertLess(corpo.index("<body lang=PT-BR>"), corpo.index("Boa tarde"))
        self.assertEqual(m.anexos, ["C:\\x\\NF 1.pdf"])
        self.assertFalse(m.enviado)

    def test_abrir_nunca_envia_e_enviar_so_o_que_esta_certo(self):
        pasta = _apoio.pasta_temporaria(self)
        criar(pasta, "1.xml", "NF 1.pdf")                 # e-mail 2 sem ticket e sem relatorio
        for modo in ("abrir", "enviar"):
            outlook = FakeOutlook()
            with mock.patch.object(em, "abrir_outlook", return_value=outlook), \
                    mock.patch.object(em, "ler_config", return_value=CFG), \
                    contextlib.redirect_stdout(io.StringIO()):
                codigo = em.main([pasta, "--modo", modo])
            if modo == "abrir":
                self.assertEqual(len(outlook.criados), 2)        # abre os dois para conferir
                self.assertTrue(all(m.aberto and not m.enviado for m in outlook.criados))
            else:
                self.assertEqual(len(outlook.criados), 1)        # so o de XML esta completo
                self.assertTrue(outlook.criados[0].enviado)
            self.assertEqual(codigo, 0 if modo == "abrir" else 2)

    def test_simular_nao_abre_o_outlook(self):
        pasta = _apoio.pasta_temporaria(self)
        with mock.patch.object(em, "abrir_outlook") as abrir, \
                mock.patch.object(em, "ler_config", return_value=CFG), \
                contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(em.main([pasta, "--modo", "simular"]), 0)
        abrir.assert_not_called()


if __name__ == "__main__":
    unittest.main()
