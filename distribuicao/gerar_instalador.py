# -*- coding: utf-8 -*-
"""Gera o Autofiscal.exe e o instalador "Autofiscal-<versao>-instalador.exe".

    python distribuicao\\gerar_instalador.py            (ou "Gerar instalador.bat")
    python distribuicao\\gerar_instalador.py --sem-instalador   (so a pasta com o .exe)

Precisa, so no computador que GERA o instalador:
    pip install pyinstaller openpyxl pypdf pillow pywin32
    Inno Setup 6  (https://jrsoftware.org/isinfo.php)

Quem INSTALA nao precisa de Python nem de biblioteca nenhuma.

Passos:
  1. PyInstaller gera a pasta Autofiscal com Autofiscal.exe (a janela),
     "Autofiscal Executor.exe" (roda os scripts) e _internal\\ (Python e
     bibliotecas), em %LOCALAPPDATA%\\Autofiscal-build\\dist (fora do OneDrive).
  2. As pastas das ferramentas e a pasta comum sao copiadas ao lado do .exe,
     como codigo-fonte: continuam editaveis e a Manutencao continua funcionando.
  3. O Inno Setup empacota tudo num instalador unico, em distribuicao\\saida\\.
"""

import argparse
import ast
import fnmatch
import importlib.util
import os
import shutil
import subprocess
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
# rascunho do build FORA do OneDrive: ele trava os arquivos enquanto sincroniza
# (PermissionError no meio do PyInstaller) e subiria dezenas de MB a cada build
TRABALHO = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
                        "Autofiscal-build")
BUILD = os.path.join(TRABALHO, "build")
DIST = os.path.join(TRABALHO, "dist")
SAIDA = os.path.join(AQUI, "saida")
PROGRAMA = os.path.join(DIST, "Autofiscal")
ICONE = os.path.join(RAIZ, "painel", "autofiscal.ico")    # o mesmo da janela do painel

# o que nao vai para o computador de quem instala
FORA_RAIZ = {"tests", "painel", "distribuicao", ".git", ".gitignore", ".gitattributes",
             "Autofiscal.pyw", "Claude outputs", "README.md"}
FORA_EXT = {".bat", ".cmd"}          # chamam python.exe, que nao existe na instalacao

BIBLIOTECAS = ["openpyxl", "pypdf", "PIL", "win32com"]


def versao():
    sys.path.insert(0, RAIZ)
    from painel import VERSAO
    return VERSAO


def arquivos_do_projeto():
    """Arquivos do projeto, sem o que o .gitignore deixa de fora (planilhas, logs, backups).

    Pergunta ao git. Sem git (projeto baixado em ZIP do GitHub, ou git nao
    instalado), le o .gitignore e percorre a pasta."""
    try:
        saida = subprocess.run(["git", "-c", "core.quotepath=off", "ls-files", "-z", "--cached",
                                "--others", "--exclude-standard"],
                               cwd=RAIZ, capture_output=True, check=True).stdout
        return [p for p in saida.decode("utf-8").split("\0") if p]
    except (OSError, subprocess.CalledProcessError):
        print("(sem git nesta pasta: usando o .gitignore para escolher os arquivos)")
        return arquivos_sem_git()


def regras_do_gitignore():
    regras = []
    try:
        with open(os.path.join(RAIZ, ".gitignore"), encoding="utf-8") as f:
            linhas = f.read().splitlines()
    except OSError:
        return regras
    for linha in linhas:
        linha = linha.strip()
        if not linha or linha.startswith(("#", "!")):
            continue
        so_pasta = linha.endswith("/")
        linha = linha.strip("/")
        # com / no meio vale para o caminho a partir da raiz; sem, para o nome em qualquer pasta
        regras.append((linha, so_pasta, "/" in linha))
    return regras


def ignorado(rel, eh_pasta, regras):
    nome = rel.rsplit("/", 1)[-1]
    for padrao, so_pasta, pelo_caminho in regras:
        if so_pasta and not eh_pasta:
            continue
        if fnmatch.fnmatchcase(rel if pelo_caminho else nome, padrao):
            return True
    return False


def arquivos_sem_git():
    regras = regras_do_gitignore() + [(".git", True, False)]
    achados = []
    for raiz, pastas, arquivos in os.walk(RAIZ):
        base = os.path.relpath(raiz, RAIZ).replace("\\", "/")
        base = "" if base == "." else base + "/"
        pastas[:] = sorted(p for p in pastas if not ignorado(base + p, True, regras))
        achados += [base + a for a in sorted(arquivos) if not ignorado(base + a, False, regras)]
    return achados


def vai_junto(rel):
    partes = rel.split("/")
    if partes[0] in FORA_RAIZ:
        return False
    if len(partes) == 1 and not rel.endswith(".md"):     # na raiz, so a documentacao
        return False
    return os.path.splitext(rel)[1].lower() not in FORA_EXT


def imports_das_ferramentas(arquivos):
    """Modulos que os scripts das ferramentas importam e que nao sao do projeto.

    O PyInstaller so enxerga o que o painel importa; os scripts das ferramentas
    rodam depois, a partir do disco, e precisam achar a biblioteca dentro do .exe."""
    locais = set()
    for rel in arquivos:
        partes = rel.split("/")
        locais.add(os.path.splitext(partes[-1])[0])
        locais.update(partes[:-1])
    achados = set()
    for rel in arquivos:
        if not rel.endswith((".py", ".pyw")):
            continue
        with open(os.path.join(RAIZ, rel), encoding="utf-8") as f:
            arvore = ast.parse(f.read())
        for no in ast.walk(arvore):
            if isinstance(no, ast.Import):
                achados.update(a.name for a in no.names)
            elif isinstance(no, ast.ImportFrom) and no.module and not no.level:
                achados.add(no.module)
    externos = set()
    for nome in achados:
        if nome.split(".")[0] in locais:
            continue
        try:
            if importlib.util.find_spec(nome) is not None:
                externos.add(nome)
        except (ImportError, ValueError):
            pass                      # import opcional (dentro de try) que nao existe aqui
    return sorted(externos)


def conferir_bibliotecas():
    falta = [b for b in BIBLIOTECAS if importlib.util.find_spec(b) is None]
    if importlib.util.find_spec("PyInstaller") is None:
        falta.append("pyinstaller")
    if falta:
        sys.exit("Falta instalar neste computador: " + ", ".join(falta)
                 + "\n  pip install pyinstaller openpyxl pypdf pillow pywin32")


def gerar_info_versao(ver):
    """Versao que aparece em Propriedades > Detalhes do .exe."""
    numeros = [int(n) for n in ver.split(".")] + [0, 0, 0, 0]
    tupla = tuple(numeros[:4])
    os.makedirs(BUILD, exist_ok=True)
    caminho = os.path.join(BUILD, "versao.txt")
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={tupla}, prodvers={tupla}),
  kids=[
    StringFileInfo([StringTable('041604B0', [
      StringStruct('ProductName', 'Autofiscal'),
      StringStruct('FileDescription', 'Autofiscal - ferramentas fiscais'),
      StringStruct('ProductVersion', '{ver}'),
      StringStruct('FileVersion', '{ver}'),
      StringStruct('LegalCopyright', 'Uso interno')])]),
    VarFileInfo([VarStruct('Translation', [0x0416, 1200])])
  ]
)
""")
    return caminho


def rodar_pyinstaller(ver, arquivos):
    ocultos = imports_das_ferramentas(arquivos)
    print(f"Bibliotecas usadas pelas ferramentas: {len(ocultos)}")
    env = dict(os.environ,
               AUTOFISCAL_RAIZ=RAIZ,
               AUTOFISCAL_OCULTOS=os.pathsep.join(ocultos),
               AUTOFISCAL_ICONE=ICONE,
               AUTOFISCAL_INFO_VERSAO=gerar_info_versao(ver))
    subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
                    "--distpath", DIST, "--workpath", os.path.join(BUILD, "pyinstaller"),
                    os.path.join(AQUI, "Autofiscal.spec")], check=True, env=env)


def copiar_ferramentas(arquivos):
    n = 0
    for rel in arquivos:
        if not vai_junto(rel):
            continue
        destino = os.path.join(PROGRAMA, *rel.split("/"))
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        shutil.copy2(os.path.join(RAIZ, *rel.split("/")), destino)
        n += 1
    print(f"{n} arquivo(s) das ferramentas copiados para {PROGRAMA}")


def achar_iscc():
    candidatos = [
        shutil.which("iscc"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Inno Setup 6", "ISCC.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Inno Setup 6", "ISCC.exe"),
        os.path.join(os.environ.get("ProgramFiles", ""), "Inno Setup 6", "ISCC.exe"),
    ]
    return next((c for c in candidatos if c and os.path.isfile(c)), None)


def rodar_inno(ver):
    iscc = achar_iscc()
    if not iscc:
        print("\nInno Setup nao encontrado: o instalador NAO foi gerado.")
        print("A pasta pronta para copiar esta em:", PROGRAMA)
        print("Para gerar o instalador, instale o Inno Setup 6 (winget install JRSoftware.InnoSetup).")
        return 1
    subprocess.run([iscc, f"/DVersao={ver}", f"/DPrograma={PROGRAMA}", f"/DSaida={SAIDA}",
                    f"/DIcone={ICONE}",
                    os.path.join(AQUI, "instalador.iss")], check=True)
    print(f"\nInstalador pronto: {os.path.join(SAIDA, f'Autofiscal-{ver}-instalador.exe')}")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Gera o instalador do Autofiscal")
    ap.add_argument("--sem-instalador", action="store_true",
                    help="so gera a pasta com o .exe, sem o Inno Setup")
    args = ap.parse_args(argv)

    conferir_bibliotecas()
    ver = versao()
    print(f"Autofiscal {ver}\n")
    arquivos = [a for a in arquivos_do_projeto() if os.path.isfile(os.path.join(RAIZ, a))]
    shutil.rmtree(PROGRAMA, ignore_errors=True)
    rodar_pyinstaller(ver, arquivos)
    copiar_ferramentas(arquivos)
    if args.sem_instalador:
        print("\nPronto (sem instalador):", PROGRAMA)
        return 0
    return rodar_inno(ver)


if __name__ == "__main__":
    sys.exit(main())
