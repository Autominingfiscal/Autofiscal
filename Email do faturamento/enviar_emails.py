# -*- coding: utf-8 -*-
"""E-MAILS DO FATURAMENTO - dois e-mails no Outlook com os documentos do dia.

    1. "FATURAMENTO / RELATORIO DE PESAGENS - <data>" com todos os XML da pasta
    2. o mesmo titulo com as NF, os TICKET e o RELATORIO.pdf da pasta

    python enviar_emails.py "<pasta do dia>" --modo simular    so lista os anexos
    python enviar_emails.py "<pasta do dia>" --modo preparar   NOVO OUTLOOK (padrao do painel)
    python enviar_emails.py "<pasta do dia>" --modo abrir      Outlook CLASSICO: abre para conferir
    python enviar_emails.py "<pasta do dia>" --modo enviar     Outlook CLASSICO: envia direto

--modo preparar: o "novo Outlook" nao aceita automacao. Para cada e-mail,
abre um e-mail novo ja com destinatarios, assunto e texto (link mailto:) e
uma pasta so com os anexos dele: e so arrastar os arquivos e enviar. Os
anexos nao vao pelo link: nenhum programa de e-mail aceita isso.

A pasta e lida com as subpastas (ex.: as "NF ... R. PORTO" do Arquivador).
Quem recebe fica no email_config.ini, nesta pasta.

Codigos de saida: 0 tudo certo | 1 erro | 2 algum e-mail ficou sem enviar
"""

# --- pasta "comum" do Autofiscal ---------------------------------------------
# Fica na pasta Autofiscal, logo acima desta ferramenta, ou dentro dela quando a
# ferramenta foi exportada para outro PC (Manutencao > Exportar ferramenta).
import os
import sys

for _pasta in (os.path.dirname(os.path.abspath(__file__)),
               os.path.dirname(os.path.dirname(os.path.abspath(__file__)))):
    if os.path.isdir(os.path.join(_pasta, "comum")):
        sys.path.insert(0, _pasta)
        break
else:
    sys.exit("Nao achei a pasta 'comum' do Autofiscal, nem nesta pasta nem na de cima.\n"
             "Para usar a ferramenta fora da pasta Autofiscal, copie-a pela ferramenta\n"
             "Manutencao > Exportar ferramenta, que leva a pasta 'comum' junto.")
# -----------------------------------------------------------------------------

import argparse
import html
import re
import shutil
import tempfile
import time
from datetime import datetime
from urllib.parse import quote

from comum.arquivos import ler_ini
from comum.caminhos import caminho_do_usuario
from comum.texto import sem_acento

PASTA = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(PASTA, "email_config.ini")
PASTAS_FORA = {"Backup", "__pycache__"}

ASSUNTO = "FATURAMENTO / RELATÓRIO DE PESAGENS - {data}"

CORPO_XML = """Boa tarde, prezados.

Espero que estejam bem e seguros.

Conforme solicitado, segue em anexo arquivos XML do faturamento de hoje ({data}).

At.te,"""

CORPO_DOCUMENTOS = """Boa tarde, prezados.

Espero que estejam bem e seguros.

Conforme solicitado, segue em anexo NF's, ticket's e relatório de pesagem emitidos na data de hoje ({data}).

At.te,"""


# ----------------------------------------------------------------- anexos ----
def tipo_do_arquivo(nome):
    """XML | NF | TICKET | RELATORIO | None, pelo nome do arquivo.

    Mesmos nomes do Arquivador: 'NF 18964 R. PORTO 19569.pdf', 'TICKET 19569.pdf'.
    'TICKET ASSINADO ...' (o escaneado) nao entra: so o ticket de pesagem."""
    base, ext = os.path.splitext(nome)
    ext = ext.lower()
    if ext == ".xml":
        return "XML"
    if ext != ".pdf":
        return None
    b = sem_acento(base).upper().strip()
    if re.match(r"NF\s*\d", b):
        return "NF"
    if re.match(r"TICKET\s*\d", b):
        return "TICKET"
    if b.startswith("RELATORIO"):
        return "RELATORIO"
    return None


def coletar(pasta):
    """{tipo: [caminhos]} da pasta e das subpastas. Arquivo com o mesmo nome em
    duas subpastas entra uma vez so."""
    achados = {"XML": [], "NF": [], "TICKET": [], "RELATORIO": []}
    vistos = set()
    for raiz, pastas, arquivos in os.walk(pasta):
        pastas[:] = sorted(p for p in pastas if p not in PASTAS_FORA and not p.startswith((".", "_")))
        for nome in sorted(arquivos):
            tipo = tipo_do_arquivo(nome)
            if tipo and nome.lower() not in vistos:
                vistos.add(nome.lower())
                achados[tipo].append(os.path.join(raiz, nome))
    return achados


def megabytes(caminhos):
    return sum(os.path.getsize(c) for c in caminhos) / (1024 * 1024)


# ----------------------------------------------------------------- config ----
def ler_config():
    if not os.path.isfile(CONFIG):
        raise ValueError(f"Nao achei o {os.path.basename(CONFIG)} nesta pasta.")
    cp = ler_ini(CONFIG, inline_comment_prefixes=(";",))

    def lista(secao, campo):
        texto = cp.get(secao, campo, fallback="") or ""
        return "; ".join(e.strip() for e in re.split(r"[;,]", texto) if e.strip())

    try:
        limite = float(cp.get("geral", "limite_mb", fallback="20").replace(",", "."))
    except ValueError:
        limite = 20.0
    return {
        "xml": {"para": lista("email_xml", "para"), "cc": lista("email_xml", "cc")},
        "documentos": {"para": lista("email_documentos", "para"),
                       "cc": lista("email_documentos", "cc")},
        "limite_mb": limite,
    }


# ---------------------------------------------------------------- outlook ----
def abrir_outlook():
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("falta a biblioteca pywin32 (pip install --user pywin32)")
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        raise RuntimeError("nao consegui abrir o Outlook. Ele precisa ser o Outlook CLASSICO "
                           f"(o 'novo Outlook' nao aceita automacao). Detalhe: {e}")


def corpo_html(texto):
    """Texto -> HTML na fonte padrao do Outlook, uma linha por paragrafo."""
    linhas = []
    for linha in texto.split("\n"):
        conteudo = html.escape(linha) if linha.strip() else "&nbsp;"
        linhas.append(f'<p style="margin:0">{conteudo}</p>')
    return ('<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">'
            + "".join(linhas) + "</div>")


def para_anexar(caminho, temporaria):
    """O Outlook nao anexa caminho com mais de ~255 caracteres (comum no OneDrive):
    nesse caso anexa uma copia feita numa pasta temporaria."""
    if len(caminho) < 250:
        return caminho
    origem = "\\\\?\\" + os.path.abspath(caminho) if os.name == "nt" else caminho
    destino = os.path.join(temporaria, os.path.basename(caminho))
    shutil.copy2(origem, destino)
    return destino


def montar_email(outlook, dest, assunto, texto, anexos, temporaria):
    m = outlook.CreateItem(0)                 # 0 = olMailItem
    m.To = dest["para"]
    m.CC = dest["cc"]
    m.Subject = assunto
    m.GetInspector                            # carrega a assinatura padrao no corpo
    atual = m.HTMLBody or ""
    corpo = corpo_html(texto)
    abre_body = re.search(r"<body[^>]*>", atual, re.I)
    m.HTMLBody = (atual[:abre_body.end()] + corpo + atual[abre_body.end():]) if abre_body \
        else corpo + atual
    for caminho in anexos:
        m.Attachments.Add(para_anexar(caminho, temporaria))
    return m


# ----------------------------------------------------------- novo outlook ----
PASTA_ANEXOS = os.path.join(tempfile.gettempdir(), "Autofiscal - anexos dos e-mails")
DIAS_GUARDADOS = 7


def link_mailto(dest, assunto, texto):
    """mailto: com para, cc, assunto e corpo. Abre um e-mail novo no programa de
    e-mail padrao do Windows (no PC da empresa, o novo Outlook)."""
    para = ",".join(e.strip() for e in dest["para"].split(";") if e.strip())
    campos = []
    cc = ",".join(e.strip() for e in dest["cc"].split(";") if e.strip())
    if cc:
        campos.append("cc=" + quote(cc, safe="@,"))
    campos.append("subject=" + quote(assunto, safe=""))
    campos.append("body=" + quote(texto.replace("\n", "\r\n"), safe=""))
    return "mailto:" + quote(para, safe="@,") + "?" + "&".join(campos)


def _limpar_anexos_velhos(base):
    limite = time.time() - DIAS_GUARDADOS * 86400
    try:
        for nome in os.listdir(base):
            caminho = os.path.join(base, nome)
            if os.path.isdir(caminho) and os.path.getmtime(caminho) < limite:
                shutil.rmtree(caminho, ignore_errors=True)
    except OSError:
        pass


def separar_anexos(email, data, base=PASTA_ANEXOS):
    """Copia os anexos de um e-mail para uma pasta so deles e devolve a pasta.
    Assim e so selecionar tudo (Ctrl+A) e arrastar para o e-mail."""
    os.makedirs(base, exist_ok=True)
    _limpar_anexos_velhos(base)
    nome = f"{data.replace('/', '-')} - {email['nome']}"
    pasta = os.path.join(base, re.sub(r'[\\/:*?"<>|]', "", nome))
    shutil.rmtree(pasta, ignore_errors=True)              # rodou de novo: so os anexos de agora
    os.makedirs(pasta)
    for caminho in email["anexos"]:
        origem = "\\\\?\\" + os.path.abspath(caminho) if os.name == "nt" else caminho
        shutil.copy2(origem, os.path.join(pasta, os.path.basename(caminho)))
    return pasta


def abrir_no_windows(alvo):
    os.startfile(alvo)                                    # pasta no Explorer / mailto no e-mail


def preparar(emails, data):
    """Um e-mail por vez: abre a pasta dos anexos e o e-mail novo, e espera."""
    prontos = [e for e in emails if e["anexos"]]
    for n, e in enumerate(prontos, start=1):
        pasta = separar_anexos(e, data)
        print(f"\n>>> {e['nome']}: {len(e['anexos'])} anexo(s)")
        abrir_no_windows(pasta)
        abrir_no_windows(link_mailto(e["dest"], e["assunto"], e["texto"]))
        print("    1. Um e-mail novo abriu com destinatarios, assunto e texto.")
        print(f"    2. Na pasta que abriu ({os.path.basename(pasta)}), selecione tudo (Ctrl+A)")
        print("       e arraste os arquivos para dentro do e-mail.")
        print("    3. Confira e clique em Enviar.")
        if n < len(prontos):
            try:
                input("    Depois de enviar, responda qualquer coisa aqui para preparar o proximo: ")
            except EOFError:
                pass
    return prontos


# ------------------------------------------------------------------ plano ----
def planejar(pasta, cfg, data):
    achados = coletar(pasta)
    assunto = ASSUNTO.format(data=data)
    emails = [
        {"nome": "E-mail 1 (XML)", "dest": cfg["xml"], "assunto": assunto,
         "texto": CORPO_XML.format(data=data), "anexos": achados["XML"], "problemas": [],
         "avisos": []},
        {"nome": "E-mail 2 (NF, tickets e relatorio)", "dest": cfg["documentos"],
         "assunto": assunto, "texto": CORPO_DOCUMENTOS.format(data=data),
         "anexos": achados["NF"] + achados["TICKET"] + achados["RELATORIO"], "problemas": [],
         "avisos": []},
    ]
    e1, e2 = emails
    if not achados["XML"]:
        e1["problemas"].append("nenhum XML na pasta")
    if not achados["NF"]:
        e2["problemas"].append("nenhuma NF (PDF com nome 'NF <numero> ...')")
    if not achados["TICKET"]:
        e2["problemas"].append("nenhum TICKET (PDF com nome 'TICKET <numero>')")
    if not achados["RELATORIO"]:
        e2["problemas"].append("nao achei o RELATORIO.pdf")
    elif len(achados["RELATORIO"]) > 1:
        e2["avisos"].append(f"{len(achados['RELATORIO'])} relatorios na pasta: vao todos")
    if achados["NF"] and achados["TICKET"] and len(achados["NF"]) != len(achados["TICKET"]):
        e2["avisos"].append(f"{len(achados['NF'])} NF e {len(achados['TICKET'])} tickets: "
                            "quantidades diferentes")
    for e in emails:
        if not e["dest"]["para"]:
            e["problemas"].append("ninguem em 'para' no email_config.ini")
        tamanho = megabytes(e["anexos"]) if e["anexos"] else 0
        e["mb"] = tamanho
        if tamanho > cfg["limite_mb"]:
            e["problemas"].append(f"anexos somam {tamanho:.1f} MB, acima do limite de "
                                  f"{cfg['limite_mb']:g} MB (o servidor pode recusar)")
    return emails


def mostrar_plano(emails, pasta):
    for e in emails:
        print("=" * 70)
        print(f"{e['nome']}  ->  {len(e['anexos'])} anexo(s), {e['mb']:.1f} MB")
        print(f"  Para: {e['dest']['para'] or '(ninguem)'}")
        if e["dest"]["cc"]:
            print(f"  Cc:   {e['dest']['cc']}")
        print(f"  Assunto: {e['assunto']}")
        for caminho in e["anexos"]:
            print(f"    - {os.path.relpath(caminho, pasta)}")
        for a in e["avisos"]:
            print(f"  ATENÇÃO: {a}")
        for p in e["problemas"]:
            print(f"  ERRO: {p}")
    print("=" * 70)


def main(argv=None):
    ap = argparse.ArgumentParser(description="E-mails do faturamento no Outlook")
    ap.add_argument("pasta", help="pasta com os XML, as NF, os tickets e o relatorio")
    ap.add_argument("--modo", choices=("simular", "preparar", "abrir", "enviar"), default="simular")
    ap.add_argument("--data", help="data no texto (dd/mm/aaaa). Padrao: hoje")
    args = ap.parse_args(argv)

    pasta = caminho_do_usuario(args.pasta)
    if not os.path.isdir(pasta):
        print(f"ERRO: a pasta nao existe: {pasta}")
        return 1
    data = args.data or datetime.now().strftime("%d/%m/%Y")
    try:
        cfg = ler_config()
    except (ValueError, OSError) as e:
        print(f"ERRO: {e}")
        return 1

    print(f"Pasta: {pasta}")
    print(f"Data no e-mail: {data}\n")
    emails = planejar(pasta, cfg, data)
    mostrar_plano(emails, pasta)
    if args.modo == "simular":
        print("\nSimulacao: nenhum e-mail foi criado.")
        return 0
    if args.modo == "preparar":
        prontos = preparar(emails, data)
        if not prontos:
            print("\nERRO: nenhum e-mail tem anexo. Veja os erros acima.")
            return 1
        sem = [e for e in emails if e not in prontos]
        for e in sem:
            print(f"ATENÇÃO: {e['nome']} NAO foi preparado (sem anexos).")
        print("\nPronto. Os e-mails so saem quando voce clica em Enviar no Outlook.")
        return 2 if sem else 0

    # abrir: monta o que tiver anexo, mesmo com problema (quem confere decide);
    # enviar: so o e-mail sem nenhum problema
    if args.modo == "enviar":
        prontos = [e for e in emails if not e["problemas"]]
    else:
        prontos = [e for e in emails if e["anexos"]]
    if not prontos:
        print("\nERRO: nenhum e-mail em condicoes. Veja os erros acima.")
        return 1
    try:
        outlook = abrir_outlook()
    except RuntimeError as e:
        print(f"\nERRO: {e}")
        return 1

    with tempfile.TemporaryDirectory(prefix="autofiscal_email_") as temporaria:
        for e in prontos:
            m = montar_email(outlook, e["dest"], e["assunto"], e["texto"], e["anexos"], temporaria)
            if args.modo == "enviar":
                m.Send()
                print(f"\nENVIADO: {e['nome']} para {e['dest']['para']}")
            else:
                m.Display(False)
                print(f"\nABERTO para conferir: {e['nome']} (confira e clique em Enviar no Outlook)")

    faltou = [e for e in emails if e not in prontos]
    for e in faltou:
        print(f"ATENÇÃO: {e['nome']} NAO foi {'enviado' if args.modo == 'enviar' else 'aberto'}.")
    return 2 if faltou else 0


if __name__ == "__main__":
    codigo = main()
    if sys.stdin is not None and sys.stdin.isatty():
        try:
            input("Pressione Enter para fechar...")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(codigo)
