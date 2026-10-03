# -*- coding: utf-8 -*-
"""MANUTENCAO DO AUTOFISCAL - criar, exportar e testar ferramentas.

    python manutencao.py nova "Nome da ferramenta"
        Cria a pasta da ferramenta a partir de _modelo_ferramenta, ja com o
        ferramenta.json preenchido. Depois e so clicar em Recarregar ferramentas.

    python manutencao.py exportar "<pasta da ferramenta>" "<destino>"
        Copia a ferramenta + a pasta comum para <destino>\\<ferramenta>, pronta
        para rodar em outro PC sem o resto do Autofiscal. Rodar de novo ATUALIZA
        a copia: o codigo e substituido, mas os .ini/.json de configuracao que ja
        existirem la sao mantidos.

    python manutencao.py testes
        Roda os testes automaticos de todas as ferramentas.
"""

import argparse
import json
import os
import re
import shutil
import sys
import unittest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODELO = os.path.join(RAIZ, "_modelo_ferramenta")
COMUM = os.path.join(RAIZ, "comum")

# o que e CODIGO (sempre substituido ao exportar) e o que e CONFIGURACAO
# (copiada so se ainda nao existir no destino). O resto - planilhas, PDFs,
# logs, relatorios, backups - e dado do trabalho e nao vai.
EXT_CODIGO = {".py", ".pyw", ".bat", ".cmd", ".md"}
EXT_CONFIG = {".ini", ".json"}
ARQ_CODIGO = {"ferramenta.json"}
PASTAS_FORA = {"__pycache__", "Backup", ".git", ".pytest_cache"}

_PROIBIDOS = re.compile(r'[\\/:*?"<>|]')


# =============================================================================
# Nova ferramenta
# =============================================================================
def nova(nome):
    nome = " ".join((nome or "").split())
    pasta_nome = _PROIBIDOS.sub("", nome).strip(" .")
    if not pasta_nome:
        print("ERRO: digite um nome para a ferramenta.")
        return 1
    if pasta_nome.startswith("_"):
        print("ERRO: o nome nao pode comecar com _ (o painel ignora essas pastas).")
        return 1
    destino = os.path.join(RAIZ, pasta_nome)
    if os.path.exists(destino):
        print(f"ERRO: ja existe a pasta {pasta_nome}. Escolha outro nome.")
        return 1

    ordem = _proxima_ordem()            # antes de copiar: a copia do modelo nao conta
    shutil.copytree(MODELO, destino, ignore=shutil.ignore_patterns(*PASTAS_FORA))
    arq = os.path.join(destino, "ferramenta.json")
    with open(arq, encoding="utf-8") as f:
        dados = json.load(f)
    dados["nome"] = nome
    dados["descricao"] = "Descreva aqui, em uma frase, o que a ferramenta faz."
    dados["ordem"] = ordem
    with open(arq, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
        f.write("\n")

    print(f"Criei a ferramenta: {destino}")
    print()
    print("Proximos passos:")
    print("  1. No painel, clique em 'Recarregar ferramentas' (embaixo, a esquerda).")
    print("  2. Troque o exemplo.py pelo script da ferramenta (o bloco 'comum' do")
    print("     comeco deve ficar) e ajuste os botoes no ferramenta.json.")
    print("  3. Veja o passo a passo em 'COMO ADICIONAR UMA FERRAMENTA.md'.")
    return 0


def _proxima_ordem():
    """Uma posicao depois da ultima ferramenta (Manutencao fica sempre no fim)."""
    maior = 0
    for nome in os.listdir(RAIZ):
        arq = os.path.join(RAIZ, nome, "ferramenta.json")
        if nome.startswith(("_", ".")) or not os.path.isfile(arq):
            continue
        try:
            with open(arq, encoding="utf-8-sig") as f:
                ordem = int(json.load(f).get("ordem", 0))
        except (OSError, ValueError, TypeError):
            continue
        if ordem < 900:
            maior = max(maior, ordem)
    return maior + 10


# =============================================================================
# Exportar / atualizar uma ferramenta em outro lugar
# =============================================================================
def exportar(pasta_ferramenta, destino, relato=print):
    """Copia a ferramenta e a pasta comum para destino/<nome da ferramenta>.

    Devolve {"pasta": ..., "codigo": n, "config_nova": n, "config_mantida": [..]}.
    """
    origem = os.path.abspath(pasta_ferramenta)
    if not os.path.isfile(os.path.join(origem, "ferramenta.json")):
        raise ValueError(f"{origem} nao e uma ferramenta (falta o ferramenta.json).")
    if os.path.normcase(origem) in (os.path.normcase(COMUM), os.path.normcase(RAIZ)):
        raise ValueError("Escolha a pasta de UMA ferramenta.")
    alvo = os.path.join(os.path.abspath(destino), os.path.basename(origem))
    if os.path.normcase(alvo) == os.path.normcase(origem):
        raise ValueError("O destino e a propria ferramenta. Escolha outra pasta.")

    res = {"pasta": alvo, "codigo": 0, "config_nova": 0, "config_mantida": []}
    _copiar_arvore(origem, alvo, res)
    _copiar_arvore(COMUM, os.path.join(alvo, "comum"), res, so_codigo=True)
    relato(f"Ferramenta exportada para: {alvo}")
    relato(f"  arquivos de codigo copiados/atualizados: {res['codigo']}")
    if res["config_nova"]:
        relato(f"  configuracoes copiadas (nao existiam la): {res['config_nova']}")
    if res["config_mantida"]:
        relato("  configuracoes que JA existiam la e foram mantidas:")
        for nome in res["config_mantida"]:
            relato(f"    {nome}")
    return res


def _copiar_arvore(origem, alvo, res, so_codigo=False):
    for raiz, pastas, arquivos in os.walk(origem):
        pastas[:] = [p for p in pastas if p not in PASTAS_FORA and not p.startswith(".")]
        rel = os.path.relpath(raiz, origem)
        destino_dir = os.path.normpath(os.path.join(alvo, rel))
        for nome in arquivos:
            ext = os.path.splitext(nome)[1].lower()
            de = os.path.join(raiz, nome)
            para = os.path.join(destino_dir, nome)
            if nome in ARQ_CODIGO or ext in EXT_CODIGO:
                os.makedirs(destino_dir, exist_ok=True)
                shutil.copy2(de, para)
                res["codigo"] += 1
            elif ext in EXT_CONFIG and not so_codigo:
                if os.path.exists(para):
                    res["config_mantida"].append(os.path.relpath(para, alvo))
                else:
                    os.makedirs(destino_dir, exist_ok=True)
                    shutil.copy2(de, para)
                    res["config_nova"] += 1


# =============================================================================
# Testes
# =============================================================================
def testes():
    pasta = os.path.join(RAIZ, "tests")
    suite = unittest.defaultTestLoader.discover(pasta, top_level_dir=pasta)
    resultado = unittest.TextTestRunner(stream=sys.stdout, verbosity=1).run(suite)
    print()
    if resultado.wasSuccessful():
        print(f"Tudo certo: {resultado.testsRun} teste(s) passaram.")
        return 0
    print(f"ERRO: {len(resultado.failures) + len(resultado.errors)} teste(s) falharam "
          f"de {resultado.testsRun}. Veja os detalhes acima.")
    return 1


def main(argv=None):
    ap = argparse.ArgumentParser(description="Manutencao do Autofiscal")
    sub = ap.add_subparsers(dest="comando", required=True)
    p = sub.add_parser("nova", help="cria uma ferramenta a partir do modelo")
    p.add_argument("nome")
    p = sub.add_parser("exportar", help="copia uma ferramenta para rodar fora do Autofiscal")
    p.add_argument("ferramenta")
    p.add_argument("destino")
    sub.add_parser("testes", help="roda os testes automaticos")
    args = ap.parse_args(argv)

    if args.comando == "nova":
        return nova(args.nome)
    if args.comando == "exportar":
        try:
            exportar(args.ferramenta, args.destino)
        except (ValueError, OSError) as e:
            print(f"ERRO: {e}")
            return 1
        return 0
    return testes()


if __name__ == "__main__":
    codigo = main()
    if sys.stdin is not None and sys.stdin.isatty():
        try:
            input("Pressione Enter para fechar...")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(codigo)
