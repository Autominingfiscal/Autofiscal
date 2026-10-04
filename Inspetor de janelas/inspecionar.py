# -*- coding: utf-8 -*-
"""INSPETOR DE JANELAS - o que o Windows enxerga dentro de uma janela.

Serve para saber se um programa (ex.: Datasul) pode ser automatizado pelos
campos da tela. Nao clica nem digita nada: so le.

    python inspecionar.py              espera 8 s para voce clicar na janela e a inspeciona
    python inspecionar.py --segundos 15
    python inspecionar.py --listar     lista as janelas abertas e o programa de cada uma

O resultado vai para Inspecoes\\inspecao_<data>.txt, nesta pasta. O CONTEUDO
dos campos de digitacao nao e gravado (so quantos caracteres tem), para o
arquivo poder ser enviado sem dados de clientes ou notas.

Codigos de saida: 0 inspecionou | 1 erro
"""

import argparse
import ctypes
import os
import sys
import time
from collections import Counter
from ctypes import wintypes
from datetime import datetime

PASTA = os.path.dirname(os.path.abspath(__file__))
SAIDA = os.path.join(PASTA, "Inspecoes")

LIMITE_ELEMENTOS = 3000
LIMITE_PROFUNDIDADE = 30

# programa que mostra uma janela vinda de OUTRO computador: os campos nao existem aqui
REMOTOS = {
    "wfica32.exe": "Citrix", "cdviewer.exe": "Citrix", "selfservice.exe": "Citrix",
    "mstsc.exe": "Area de Trabalho Remota", "msrdc.exe": "Area de Trabalho Remota (RemoteApp)",
    "msrdcw.exe": "Area de Trabalho Remota", "vmware-view.exe": "VMware Horizon",
    "anydesk.exe": "AnyDesk", "teamviewer.exe": "TeamViewer",
}
PROGRESS = {"prowin32.exe", "prowin.exe", "_progres.exe"}     # telas classicas do Datasul

# campos com dado digitado: so o tamanho vai para o arquivo
CLASSES_DE_DIGITACAO = ("edit", "richedit", "combobox", "fill-in")
TIPOS_DE_DIGITACAO = {"Edit", "ComboBox", "Document"}


# ------------------------------------------------------------------ janelas ----
def _texto_janela(hwnd):
    n = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(n + 1)
    ctypes.windll.user32.GetWindowTextW(hwnd, buf, n + 1)
    return buf.value


def _classe_janela(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def programa_da_janela(hwnd):
    """Caminho do .exe dono da janela ("" se o Windows nao deixar ler)."""
    pid = wintypes.DWORD()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    proc = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid.value)  # QUERY_LIMITED_INFORMATION
    if not proc:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(1024)
        tam = wintypes.DWORD(1024)
        if ctypes.windll.kernel32.QueryFullProcessImageNameW(proc, 0, buf, ctypes.byref(tam)):
            return buf.value
        return ""
    finally:
        ctypes.windll.kernel32.CloseHandle(proc)


def janelas_abertas():
    achadas = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def cada(hwnd, _):
        if ctypes.windll.user32.IsWindowVisible(hwnd) and _texto_janela(hwnd).strip():
            achadas.append(hwnd)
        return True

    ctypes.windll.user32.EnumWindows(cada, 0)
    return achadas


def tipo_de_programa(exe):
    nome = os.path.basename(exe).lower()
    if nome in REMOTOS:
        return "remoto", REMOTOS[nome]
    if nome in PROGRESS:
        return "progress", "Progress OpenEdge (telas classicas do Datasul)"
    return "local", ""


# ---------------------------------------------------------------- arvores ----
def _curto(texto, limite=70):
    texto = " ".join(str(texto or "").split())
    return texto if len(texto) <= limite else texto[:limite - 1] + "…"


def _retangulo(w):
    try:
        r = w.rectangle()
        return f"({r.left},{r.top} {r.width()}x{r.height()})"
    except Exception:
        return ""


def _percorrer(raiz, descrever, linhas, contagem):
    """Desce pela arvore de controles, com limite de tamanho e de profundidade."""
    pilha = [(raiz, 0)]
    while pilha and contagem["total"] < LIMITE_ELEMENTOS:
        w, nivel = pilha.pop()
        contagem["total"] += 1
        try:
            linhas.append("  " * nivel + descrever(w, contagem))
        except Exception as e:
            linhas.append("  " * nivel + f"(nao consegui ler este controle: {e})")
        if nivel >= LIMITE_PROFUNDIDADE:
            continue
        try:
            filhos = w.children()
        except Exception:
            filhos = []
        pilha.extend((f, nivel + 1) for f in reversed(filhos))
    if contagem["total"] >= LIMITE_ELEMENTOS:
        linhas.append(f"... parei em {LIMITE_ELEMENTOS} controles")


def _descrever_win32(w, contagem):
    classe = w.class_name()
    contagem["classes"][classe] += 1
    digitacao = any(c in classe.lower() for c in CLASSES_DE_DIGITACAO)
    texto = w.window_text()
    if digitacao:
        contagem["digitacao"] += 1
        mostrado = f"[campo de digitacao: {len(texto)} caractere(s)]"
    else:
        mostrado = repr(_curto(texto))
    estado = "" if w.is_enabled() else "  (desabilitado)"
    return f"{classe}  id={w.control_id()}  {mostrado}  {_retangulo(w)}{estado}"


def _descrever_uia(w, contagem):
    info = w.element_info
    tipo = info.control_type or "?"
    contagem["classes"][tipo] += 1
    if tipo in TIPOS_DE_DIGITACAO:
        contagem["digitacao"] += 1
    if tipo == "Button":
        contagem["botoes"] += 1
    partes = [tipo]
    if info.name:
        partes.append(repr(_curto(info.name)))
    if info.automation_id:
        partes.append(f"auto_id={info.automation_id!r}")
    if info.class_name:
        partes.append(f"classe={info.class_name}")
    partes.append(_retangulo(w))
    if not w.is_enabled():
        partes.append("(desabilitado)")
    return "  ".join(partes)


def arvore(hwnd, backend):
    from pywinauto import Desktop
    descrever = _descrever_win32 if backend == "win32" else _descrever_uia
    contagem = {"total": 0, "digitacao": 0, "botoes": 0, "classes": Counter()}
    linhas = []
    raiz = Desktop(backend=backend).window(handle=hwnd).wrapper_object()
    _percorrer(raiz, descrever, linhas, contagem)
    return linhas, contagem


# ------------------------------------------------------------------ acoes ----
def listar():
    print("JANELAS ABERTAS\n")
    for hwnd in janelas_abertas():
        exe = programa_da_janela(hwnd)
        tipo, explicacao = tipo_de_programa(exe)
        marca = {"remoto": "  <- REMOTA: " + explicacao,
                 "progress": "  <- " + explicacao}.get(tipo, "")
        print(f"- {_curto(_texto_janela(hwnd), 60)}")
        print(f"    programa: {os.path.basename(exe) or '?'}   classe: {_classe_janela(hwnd)}{marca}")
    return 0


def inspecionar(segundos, abrir=True):
    try:
        import pywinauto  # noqa: F401
    except ImportError:
        print("ERRO: falta a biblioteca pywinauto.  pip install --user pywinauto")
        return 1

    print("Clique na janela que voce quer inspecionar (ex.: a tela de emissao do Datasul).")
    for falta in range(segundos, 0, -1):
        print(f"  inspecionando em {falta}...", flush=True)
        time.sleep(1)
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    titulo = _texto_janela(hwnd)
    exe = programa_da_janela(hwnd)
    if titulo == "Autofiscal" or os.path.basename(exe).lower().startswith("autofiscal"):
        print("ATENÇÃO: a janela da frente ainda era o Autofiscal. Rode de novo e clique")
        print("na outra janela antes de terminar a contagem.")
        return 1
    tipo, explicacao = tipo_de_programa(exe)

    print(f"\nJanela: {_curto(titulo)}")
    print(f"Programa: {exe or '?'}")
    print("Lendo os controles (pode levar alguns segundos)...", flush=True)

    resultados = {}
    for backend in ("win32", "uia"):
        try:
            resultados[backend] = arvore(hwnd, backend)
        except Exception as e:
            resultados[backend] = ([f"(falhou: {type(e).__name__}: {e})"],
                                   {"total": 0, "digitacao": 0, "botoes": 0, "classes": Counter()})

    resumo = montar_resumo(titulo, exe, tipo, explicacao, resultados)
    print()
    print("\n".join(resumo))

    os.makedirs(SAIDA, exist_ok=True)
    caminho = os.path.join(SAIDA, datetime.now().strftime("inspecao_%Y%m%d_%H%M%S.txt"))
    with open(caminho, "w", encoding="utf-8") as f:
        f.write("INSPECAO DE JANELA - " + datetime.now().strftime("%d/%m/%Y %H:%M:%S") + "\n")
        f.write("O conteudo dos campos de digitacao NAO foi gravado, so o tamanho.\n\n")
        f.write("\n".join(resumo) + "\n")
        for backend, nome in (("win32", "WIN32 (controles classicos do Windows)"),
                              ("uia", "UI AUTOMATION (acessibilidade)")):
            f.write("\n" + "=" * 70 + f"\n{nome}\n" + "=" * 70 + "\n")
            f.write("\n".join(resultados[backend][0]) + "\n")
    print(f"\nArquivo com todos os detalhes: {caminho}")
    if abrir:
        try:
            os.startfile(caminho)
        except OSError:
            pass
    return 0


def montar_resumo(titulo, exe, tipo, explicacao, resultados):
    w32, uia = resultados["win32"][1], resultados["uia"][1]
    linhas = ["RESUMO",
              f"  janela ........: {_curto(titulo)}",
              f"  programa ......: {os.path.basename(exe) or '?'}"
              + (f"  ({explicacao})" if explicacao else ""),
              f"  win32 .........: {w32['total']} controle(s), {w32['digitacao']} de digitacao",
              f"  ui automation .: {uia['total']} controle(s), {uia['digitacao']} de digitacao, "
              f"{uia['botoes']} botao(oes)"]
    mais = ", ".join(f"{c} x{n}" for c, n in w32["classes"].most_common(6))
    if mais:
        linhas.append(f"  classes win32 .: {mais}")
    linhas.append("")
    if tipo == "remoto":
        linhas.append("CONCLUSAO: a janela vem de OUTRO computador (" + explicacao + "). Daqui so")
        linhas.append("da para ver a imagem: automatizar seria por imagem e posicao na tela.")
    elif max(w32["digitacao"], uia["digitacao"]) > 0:
        linhas.append("CONCLUSAO: o Windows enxerga os campos de digitacao desta tela. Da para")
        linhas.append("automatizar pelos campos (pywinauto), sem depender da posicao na tela.")
    elif max(w32["total"], uia["total"]) > 1:
        linhas.append("ATENÇÃO: o Windows enxerga a janela, mas nenhum campo de digitacao. Abra a")
        linhas.append("tela com os campos (nao o menu) e inspecione de novo. Se continuar assim,")
        linhas.append("os campos sao desenhados pelo programa e a automacao teria que ser por imagem.")
    else:
        linhas.append("ATENÇÃO: o Windows nao enxerga nada dentro da janela: a automacao teria")
        linhas.append("que ser por imagem e posicao na tela.")
    return linhas


def main(argv=None):
    ap = argparse.ArgumentParser(description="Inspetor de janelas")
    ap.add_argument("--listar", action="store_true", help="lista as janelas abertas")
    ap.add_argument("--segundos", type=int, default=8, help="tempo para clicar na janela")
    ap.add_argument("--nao-abrir", action="store_true", help="nao abre o .txt no fim")
    args = ap.parse_args(argv)
    try:
        # titulo com simbolo que o console nao tem (ex.: emoji na aba do navegador)
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    if os.name != "nt":
        print("ERRO: o inspetor so funciona no Windows.")
        return 1
    if args.listar:
        return listar()
    return inspecionar(max(1, args.segundos), abrir=not args.nao_abrir)


if __name__ == "__main__":
    codigo = main()
    if sys.stdin is not None and sys.stdin.isatty():
        try:
            input("Pressione Enter para fechar...")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(codigo)
