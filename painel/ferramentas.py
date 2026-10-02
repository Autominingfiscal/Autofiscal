"""Leitura dos arquivos ferramenta.json de cada subpasta."""

import importlib.util
import json
import os
import re

ARQUIVO = "ferramenta.json"

# {tipo:titulo|filtro}  ou  {tipo?:titulo}  (com ? = opcional)
MARCADOR = re.compile(r"\{(pasta|arquivo|texto)(\?)?:([^}|]*)(?:\|([^}]*))?\}")


class Acao:
    def __init__(self, dados, ferramenta):
        self.ferramenta = ferramenta
        self.rotulo = dados.get("rotulo", "Executar")
        self.descricao = dados.get("descricao", "")
        self.script = dados.get("script", "")
        self.args = dados.get("args", [])
        self.continuo = bool(dados.get("continuo", False))
        self.confirmar = dados.get("confirmar", "")
        self.principal = bool(dados.get("principal", False))

    @property
    def caminho_script(self):
        return os.path.join(self.ferramenta.pasta, self.script)

    def marcadores(self):
        """Lista (marcador, tipo, opcional, titulo, filtro) de tudo que precisa perguntar.

        `marcador` e o texto exato do marcador no ferramenta.json: e a chave que
        montar_args usa para trocar pela resposta."""
        achados = []
        for item in self.args:
            for texto in (item if isinstance(item, list) else [item]):
                for m in MARCADOR.finditer(str(texto)):
                    achados.append((m.group(0), m.group(1), bool(m.group(2)), m.group(3).strip(),
                                    m.group(4) or ""))
        return achados

    def montar_args(self, respostas):
        """Troca os marcadores pelas respostas. Grupo com opcional vazio sai inteiro."""
        saida = []
        for item in self.args:
            grupo = item if isinstance(item, list) else [item]
            montado, descartar = [], False
            for texto in grupo:
                def troca(m):
                    return respostas.get(m.group(0), "")
                novo = MARCADOR.sub(troca, str(texto))
                if MARCADOR.search(str(texto)) and not novo.strip():
                    descartar = True
                montado.append(novo)
            if not descartar:
                saida.extend(montado)
        return saida


class Ferramenta:
    def __init__(self, pasta):
        self.pasta = pasta
        self.id = os.path.basename(pasta)
        self.erro = ""
        dados = {}
        try:
            with open(os.path.join(pasta, ARQUIVO), encoding="utf-8-sig") as f:
                dados = json.load(f)
        except (OSError, ValueError) as e:
            self.erro = f"Nao consegui ler o {ARQUIVO}: {e}"
        self.nome = dados.get("nome", self.id)
        self.descricao = dados.get("descricao", "")
        self.ordem = dados.get("ordem", 100)
        self.config = dados.get("config") or ""
        self.campos = dados.get("campos", {})
        self.ajuda = dados.get("ajuda") or ""
        self.requer = dados.get("requer", [])
        self.acoes = [Acao(a, self) for a in dados.get("acoes", [])]

    @property
    def caminho_config(self):
        return os.path.join(self.pasta, self.config) if self.config else ""

    @property
    def caminho_ajuda(self):
        return os.path.join(self.pasta, self.ajuda) if self.ajuda else ""

    def faltando(self):
        """Bibliotecas que a ferramenta usa e nao estao instaladas."""
        falta = []
        for r in self.requer:
            modulo = r.get("modulo") if isinstance(r, dict) else r
            pacote = r.get("pacote", modulo) if isinstance(r, dict) else r
            try:
                achou = importlib.util.find_spec(modulo) is not None
            except (ImportError, ValueError):
                achou = False
            if not achou:
                falta.append(pacote)
        return falta


def carregar_todas(raiz):
    raiz = os.path.abspath(raiz)
    ferramentas = []
    for nome in sorted(os.listdir(raiz)):
        pasta = os.path.join(raiz, nome)
        if nome.startswith(("_", ".")) or not os.path.isdir(pasta):
            continue
        if os.path.isfile(os.path.join(pasta, ARQUIVO)):
            ferramentas.append(Ferramenta(pasta))
    ferramentas.sort(key=lambda f: (f.ordem, f.nome.lower()))
    return ferramentas
