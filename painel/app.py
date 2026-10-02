"""Janela principal do painel Autofiscal (tkinter, ja vem com o Python)."""

import datetime as dt
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import VERSAO
from .caminhos import caminho_do_usuario, existe, gravar_estado, ler_estado
from .execucao import Execucao
from .ferramentas import MARCADOR, carregar_todas
from .ini import DocumentoIni

# ------------------------------------------------------------------ visual ----
COR = {
    "fundo": "#f3f5f8", "cartao": "#ffffff", "borda": "#dfe3ea", "texto": "#1f2937",
    "suave": "#6b7280", "lateral": "#1e293b", "lateral_txt": "#cbd5e1",
    "lateral_sel": "#334155", "destaque": "#2563eb", "destaque_hover": "#1d4ed8",
    "verde": "#16a34a", "laranja": "#d97706", "vermelho": "#dc2626",
    "log_fundo": "#0f172a", "log_txt": "#e2e8f0",
}
FONTE = "Segoe UI" if os.name == "nt" else "DejaVu Sans"
MONO = "Consolas" if os.name == "nt" else "DejaVu Sans Mono"
MAX_LINHAS_LOG = 4000

PALAVRAS_ERRO = ("ERRO", "Traceback", "Error", "NAO DEU", "NÃO DEU", "Exception")
PALAVRAS_AVISO = ("ATEN", "AVISO", "Aviso", "pendente", "PENDENTE", "Falta")


def agora():
    return dt.datetime.now().strftime("%H:%M:%S")


def abrir_no_windows(caminho):
    try:
        os.startfile(caminho)                     # Windows
    except AttributeError:
        import subprocess
        subprocess.Popen(["xdg-open", caminho])
    except OSError as e:
        messagebox.showerror("Autofiscal", f"Nao consegui abrir:\n{caminho}\n\n{e}")


class Botao(tk.Label):
    """Botao chato (flat) com cor, que fica igual em qualquer Windows."""

    def __init__(self, pai, texto, comando, primario=False, perigo=False, **kw):
        cor = COR["vermelho"] if perigo else COR["destaque"] if primario else "#e5e7eb"
        txt = "#ffffff" if (primario or perigo) else COR["texto"]
        super().__init__(pai, text=texto, bg=cor, fg=txt, cursor="hand2",
                         font=(FONTE, 10, "bold" if primario else "normal"),
                         padx=14, pady=7, **kw)
        self._cor, self._hover = cor, ("#b91c1c" if perigo else COR["destaque_hover"] if primario else "#d1d5db")
        self._comando = comando
        self._ativo = True
        self.bind("<Button-1>", lambda e: self._ativo and self._comando())
        self.bind("<Enter>", lambda e: self._ativo and self.config(bg=self._hover))
        self.bind("<Leave>", lambda e: self._ativo and self.config(bg=self._cor))

    def ativar(self, sim=True):
        self._ativo = sim
        self.config(bg=self._cor if sim else "#e5e7eb",
                    fg=self.cget("fg") if sim else "#9ca3af",
                    cursor="hand2" if sim else "arrow")
        if sim:
            self.config(fg="#ffffff" if self._cor != "#e5e7eb" else COR["texto"])


# =============================================================================
# Pagina de uma ferramenta
# =============================================================================
class PaginaFerramenta(tk.Frame):
    def __init__(self, pai, app, ferramenta):
        super().__init__(pai, bg=COR["fundo"])
        self.app = app
        self.f = ferramenta
        self.execucao = None
        self.botoes_acao = []
        self.doc = None
        self._montar()

    # ---------------------------------------------------------------- tela ----
    def _montar(self):
        f = self.f
        topo = tk.Frame(self, bg=COR["fundo"])
        topo.pack(fill="x", padx=24, pady=(20, 6))
        tk.Label(topo, text=f.nome, font=(FONTE, 18, "bold"), bg=COR["fundo"],
                 fg=COR["texto"]).pack(anchor="w")
        if f.descricao:
            tk.Label(topo, text=f.descricao, font=(FONTE, 10), bg=COR["fundo"], fg=COR["suave"],
                     wraplength=820, justify="left").pack(anchor="w", pady=(2, 0))

        avisos = []
        if f.erro:
            avisos.append(f.erro)
        falta = f.faltando()
        if falta:
            avisos.append("Falta instalar: " + ", ".join(falta)
                          + ".  Clique em \"Instalar\" ou rode no Prompt:  pip install --user "
                          + " ".join(falta))
        if avisos:
            faixa = tk.Frame(self, bg="#fef3c7", highlightbackground="#f59e0b", highlightthickness=1)
            faixa.pack(fill="x", padx=24, pady=(6, 0))
            tk.Label(faixa, text="  ".join(avisos), bg="#fef3c7", fg="#92400e", font=(FONTE, 9),
                     wraplength=700, justify="left").pack(side="left", padx=10, pady=8)
            if falta:
                Botao(faixa, "Instalar", lambda: self._instalar(falta)).pack(side="right", padx=8, pady=6)

        abas = ttk.Notebook(self)
        abas.pack(fill="both", expand=True, padx=24, pady=(12, 20))
        self.abas = abas

        exe = tk.Frame(abas, bg=COR["fundo"])
        abas.add(exe, text="   Executar   ")
        self._montar_executar(exe)

        if f.config:
            cfg = tk.Frame(abas, bg=COR["fundo"])
            abas.add(cfg, text="   Configurações   ")
            self.aba_config = cfg
            self._montar_config(cfg)

    def _montar_executar(self, pai):
        cartao = tk.Frame(pai, bg=COR["cartao"], highlightbackground=COR["borda"], highlightthickness=1)
        cartao.pack(fill="x", pady=(10, 10))
        if not self.f.acoes:
            tk.Label(cartao, text="Esta ferramenta nao tem acoes cadastradas no ferramenta.json.",
                     bg=COR["cartao"], fg=COR["suave"]).pack(padx=14, pady=14)
        for i, acao in enumerate(self.f.acoes):
            linha = tk.Frame(cartao, bg=COR["cartao"])
            linha.pack(fill="x", padx=14, pady=(12 if i == 0 else 4, 4))
            b = Botao(linha, acao.rotulo, lambda a=acao: self.executar(a),
                      primario=acao.principal or i == 0, width=26)
            b.pack(side="left")
            self.botoes_acao.append(b)
            if acao.descricao:
                tk.Label(linha, text=acao.descricao, bg=COR["cartao"], fg=COR["suave"],
                         font=(FONTE, 9), wraplength=560, justify="left").pack(side="left", padx=14)
        tk.Frame(cartao, bg=COR["cartao"], height=8).pack()

        barra = tk.Frame(pai, bg=COR["fundo"])
        barra.pack(fill="x")
        self.status = tk.Label(barra, text="●  Parado", bg=COR["fundo"], fg=COR["suave"],
                               font=(FONTE, 10, "bold"))
        self.status.pack(side="left")
        self.bt_parar = Botao(barra, "■  Parar", self.parar, perigo=True)
        self.bt_parar.pack(side="right")
        self.bt_parar.ativar(False)
        Botao(barra, "Limpar", self.limpar).pack(side="right", padx=6)
        Botao(barra, "Abrir pasta", lambda: abrir_no_windows(self.f.pasta)).pack(side="right")
        if self.f.ajuda and os.path.exists(self.f.caminho_ajuda):
            Botao(barra, "Ajuda", lambda: abrir_no_windows(self.f.caminho_ajuda)).pack(side="right", padx=6)

        caixa = tk.Frame(pai, bg=COR["log_fundo"])
        caixa.pack(fill="both", expand=True, pady=(8, 0))
        self.log = tk.Text(caixa, bg=COR["log_fundo"], fg=COR["log_txt"], insertbackground="#fff",
                           font=(MONO, 10), wrap="word", relief="flat", padx=10, pady=8,
                           state="disabled", height=14)
        rolar = ttk.Scrollbar(caixa, command=self.log.yview)
        self.log.config(yscrollcommand=rolar.set)
        rolar.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        for tag, cor in (("erro", "#f87171"), ("aviso", "#fbbf24"), ("ok", "#4ade80"),
                         ("eco", "#93c5fd"), ("painel", "#94a3b8")):
            self.log.tag_config(tag, foreground=cor)
        self._escrever("Pronto. Escolha uma acao acima.\n", "painel")

        resp = tk.Frame(pai, bg=COR["fundo"])
        resp.pack(fill="x", pady=(8, 0))
        tk.Label(resp, text="Se o programa pedir uma resposta, digite aqui:", bg=COR["fundo"],
                 fg=COR["suave"], font=(FONTE, 9)).pack(side="left")
        self.entrada = ttk.Entry(resp, font=(FONTE, 10))
        self.entrada.pack(side="left", fill="x", expand=True, padx=8)
        self.entrada.bind("<Return>", lambda e: self.enviar())
        self.bt_enviar = Botao(resp, "Enviar", self.enviar)
        self.bt_enviar.pack(side="left")
        self._ajustar_botoes()

    # ------------------------------------------------------------- execucao ----
    def executar(self, acao):
        if self.execucao and self.execucao.rodando:
            messagebox.showinfo("Autofiscal", f"\"{self.execucao.acao.rotulo}\" ainda esta rodando.\n"
                                "Espere terminar ou clique em Parar.")
            return
        if not os.path.exists(acao.caminho_script):
            messagebox.showerror("Autofiscal", f"Nao achei o script:\n{acao.caminho_script}")
            return
        respostas = self._perguntar(acao)
        if respostas is None:
            return
        if acao.confirmar:
            detalhe = "\n".join(f"• {v}" for v in respostas.values() if v)
            if not messagebox.askyesno(acao.rotulo, acao.confirmar + ("\n\n" + detalhe if detalhe else "")):
                return
        args = acao.montar_args(respostas)
        self._escrever(f"\n▶ {acao.rotulo}  ({agora()})\n", "painel")
        self.execucao = Execucao(acao, args)
        try:
            self.execucao.iniciar()
        except OSError as e:
            self._escrever(f"Nao consegui iniciar: {e}\n", "erro")
            self.execucao = None
            return
        self.app.registrar(self)
        self._ajustar_botoes()
        self.entrada.focus_set()

    def _perguntar(self, acao):
        estado = ler_estado()
        respostas = {}
        for marca, tipo, opcional, titulo, filtro in acao.marcadores():
            chave = f"{self.f.id}|{tipo}|{titulo}"
            anterior = estado.get(chave, "")
            if tipo == "pasta":
                valor = filedialog.askdirectory(parent=self, title=titulo or "Escolha a pasta",
                                                initialdir=anterior if os.path.isdir(anterior or "") else None)
            elif tipo == "arquivo":
                tipos = [("Arquivos", filtro)] if filtro else []
                tipos.append(("Todos os arquivos", "*.*"))
                inicio = os.path.dirname(anterior) if anterior else None
                valor = filedialog.askopenfilename(parent=self, title=titulo or "Escolha o arquivo",
                                                   filetypes=tipos,
                                                   initialdir=inicio if inicio and os.path.isdir(inicio) else None)
            else:
                valor = simpledialog.askstring(acao.rotulo, titulo, parent=self,
                                               initialvalue=anterior or "")
            if valor is None and tipo == "texto":
                return None             # Cancelar na caixa de texto = desistir de rodar
            if not valor:
                # pasta/arquivo cancelado chega como "": nao da para separar de
                # "deixei em branco", entao so a pergunta opcional segue
                if opcional:
                    respostas[marca] = ""
                    continue
                return None
            valor = os.path.normpath(valor) if tipo in ("pasta", "arquivo") else valor.strip()
            respostas[marca] = valor
            estado[chave] = valor
        gravar_estado(estado)
        return respostas

    def enviar(self):
        texto = self.entrada.get()
        if self.execucao and self.execucao.responder(texto):
            self._escrever(texto + "\n", "eco")
            self.entrada.delete(0, "end")

    def parar(self):
        if self.execucao and self.execucao.rodando:
            self.execucao.parado_pelo_usuario = True
            self.execucao.parar()

    def limpar(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _instalar(self, pacotes):
        if self.rodando:
            messagebox.showinfo("Autofiscal", "Pare a ferramenta antes de instalar.")
            return
        from .ferramentas import Acao
        acao = Acao({"rotulo": "Instalar " + ", ".join(pacotes)}, self.f)
        self.abas.select(0)
        self._escrever(f"\n▶ Instalando {' '.join(pacotes)}  ({agora()})\n", "painel")
        ex = Execucao(acao, [], comando=["-m", "pip", "install", "--user"] + list(pacotes))
        try:
            ex.iniciar()
        except OSError as e:
            self._escrever(f"Nao consegui iniciar o pip: {e}\n", "erro")
            return
        self.execucao = ex
        self.app.registrar(self)
        self._ajustar_botoes()

    # chamado pelo app a cada 100 ms
    def receber(self):
        if not self.execucao:
            return False
        terminou = False
        while True:
            try:
                pedaco = self.execucao.saida.get_nowait()
            except Exception:
                break
            if pedaco is None:
                terminou = True
                break
            self._escrever(pedaco)
        if terminou:
            cod = self.execucao.codigo
            if getattr(self.execucao, "parado_pelo_usuario", False):
                self._escrever(f"■ Parado ({agora()})\n", "aviso")
            elif cod == 0:
                self._escrever(f"✔ Terminou ({agora()})\n", "ok")
            else:
                self._escrever(f"✖ Terminou com erro (codigo {cod}) ({agora()})\n", "erro")
            self._ajustar_botoes()
        return terminou

    def _escrever(self, texto, tag=None):
        self.log.config(state="normal")
        for linha in texto.splitlines(keepends=True):
            t = tag
            if t is None:
                if any(p in linha for p in PALAVRAS_ERRO):
                    t = "erro"
                elif any(p in linha for p in PALAVRAS_AVISO):
                    t = "aviso"
            self.log.insert("end", linha, t or ())
        excesso = int(self.log.index("end-1c").split(".")[0]) - MAX_LINHAS_LOG
        if excesso > 0:
            self.log.delete("1.0", f"{excesso + 1}.0")
        self.log.see("end")
        self.log.config(state="disabled")

    def _ajustar_botoes(self):
        rodando = bool(self.execucao and self.execucao.rodando)
        for b in self.botoes_acao:
            b.ativar(not rodando)
        self.bt_parar.ativar(rodando)
        self.bt_enviar.ativar(rodando)
        if rodando:
            self.status.config(text=f"●  Rodando: {self.execucao.acao.rotulo}", fg=COR["verde"])
        else:
            self.status.config(text="●  Parado", fg=COR["suave"])
        self.app.atualizar_lateral()

    @property
    def rodando(self):
        return bool(self.execucao and self.execucao.rodando)

    # ------------------------------------------------------- configuracoes ----
    def _montar_config(self, pai):
        for w in pai.winfo_children():
            w.destroy()
        self.vars = {}
        caminho = self.f.caminho_config
        if not os.path.exists(caminho):
            tk.Label(pai, text=f"O arquivo {self.f.config} ainda nao existe.\n"
                               "Ele e criado na primeira vez que a ferramenta roda. "
                               "Rode uma vez e depois volte aqui.",
                     bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 10), justify="left").pack(
                anchor="w", padx=6, pady=16)
            Botao(pai, "Recarregar", lambda: self._montar_config(pai)).pack(anchor="w", padx=6)
            return
        try:
            self.doc = DocumentoIni(caminho)
        except OSError as e:
            tk.Label(pai, text=f"Nao consegui ler {caminho}: {e}", bg=COR["fundo"],
                     fg=COR["vermelho"]).pack(anchor="w", pady=16)
            return

        barra = tk.Frame(pai, bg=COR["fundo"])
        barra.pack(fill="x", side="bottom", pady=(10, 0))
        self.bt_salvar = Botao(barra, "Salvar alterações", self._salvar_config, primario=True)
        self.bt_salvar.pack(side="left")
        Botao(barra, "Descartar", lambda: self._montar_config(pai)).pack(side="left", padx=6)
        Botao(barra, "Abrir no Bloco de Notas", lambda: abrir_no_windows(caminho)).pack(side="left")
        self.aviso_cfg = tk.Label(barra, text="", bg=COR["fundo"], fg=COR["suave"], font=(FONTE, 9))
        self.aviso_cfg.pack(side="left", padx=12)

        # area com rolagem
        moldura = tk.Frame(pai, bg=COR["fundo"])
        moldura.pack(fill="both", expand=True, pady=(10, 0))
        tela = tk.Canvas(moldura, bg=COR["fundo"], highlightthickness=0)
        rolar = ttk.Scrollbar(moldura, command=tela.yview)
        corpo = tk.Frame(tela, bg=COR["fundo"])
        corpo.bind("<Configure>", lambda e: tela.configure(scrollregion=tela.bbox("all")))
        janela = tela.create_window((0, 0), window=corpo, anchor="nw")
        tela.bind("<Configure>", lambda e: tela.itemconfigure(janela, width=e.width))
        tela.configure(yscrollcommand=rolar.set)
        tela.pack(side="left", fill="both", expand=True)
        rolar.pack(side="right", fill="y")
        # roda do mouse rola o formulario so enquanto o mouse esta em cima dele
        tela.bind("<Enter>", lambda e: tela.bind_all("<MouseWheel>", lambda ev: tela.yview_scroll(
            int(-ev.delta / 120), "units")))
        tela.bind("<Leave>", lambda e: tela.unbind_all("<MouseWheel>"))

        base = self.f.pasta
        for secao in self.doc.secoes():
            cartao = tk.Frame(corpo, bg=COR["cartao"], highlightbackground=COR["borda"],
                              highlightthickness=1)
            cartao.pack(fill="x", pady=(0, 12), padx=(0, 6))
            tk.Label(cartao, text=(secao or "Geral").replace("_", " ").upper(), bg=COR["cartao"],
                     fg=COR["destaque"], font=(FONTE, 9, "bold")).pack(anchor="w", padx=14, pady=(10, 2))
            for campo in [c for c in self.doc.campos if c.secao == secao]:
                self._campo(cartao, campo, base)
            tk.Frame(cartao, bg=COR["cartao"], height=6).pack()
        self._marcar_sujo()

    def _tipo(self, campo):
        tipo = self.f.campos.get(campo.chave, {})
        tipo = tipo.get("tipo") if isinstance(tipo, dict) else tipo
        if tipo:
            return tipo
        v = campo.valor.strip().strip('"')
        if v.lower() in ("sim", "nao", "não", "s", "n"):
            return "simnao"
        if "\\" in v or "/" in v or "%" in v or v.lower().endswith((".xlsx", ".xlsm", ".pdf")):
            k = campo.chave.lower()
            if "pasta" in k or "pdf" in k or not os.path.splitext(v)[1]:
                return "pasta"
            return "arquivo"
        return "texto"

    def _campo(self, pai, campo, base):
        tipo = self._tipo(campo)
        linha = tk.Frame(pai, bg=COR["cartao"])
        linha.pack(fill="x", padx=14, pady=(8, 0))
        rotulo = campo.chave.replace("_", " ")
        rotulo = rotulo[:1].upper() + rotulo[1:]
        tk.Label(linha, text=rotulo, bg=COR["cartao"], fg=COR["texto"],
                 font=(FONTE, 10, "bold")).pack(anchor="w")
        ajuda = self.f.campos.get(campo.chave, {}).get("ajuda") if isinstance(
            self.f.campos.get(campo.chave), dict) else None
        if ajuda or campo.ajuda:
            tk.Label(linha, text=ajuda or campo.ajuda, bg=COR["cartao"], fg=COR["suave"],
                     font=(FONTE, 9), wraplength=720, justify="left").pack(anchor="w")
        var = tk.StringVar(value=campo.valor)
        self.vars[id(campo)] = (campo, var)
        caixa = tk.Frame(linha, bg=COR["cartao"])
        caixa.pack(fill="x", pady=(4, 0))
        if tipo == "simnao":
            maiusc = campo.valor.strip().isupper()
            opcoes = ["SIM", "NAO"] if maiusc else ["sim", "nao"]
            ttk.Combobox(caixa, textvariable=var, values=opcoes, state="readonly",
                         width=10, font=(FONTE, 10)).pack(side="left")
        else:
            ttk.Entry(caixa, textvariable=var, font=(FONTE, 10)).pack(side="left", fill="x", expand=True)
        situacao = tk.Label(linha, text="", bg=COR["cartao"], font=(FONTE, 9))
        if tipo in ("pasta", "arquivo"):
            Botao(caixa, "Procurar…", lambda: self._procurar(var, tipo, base)).pack(side="left", padx=(6, 0))
            situacao.pack(anchor="w")

        def mudou(*_):
            campo.valor = var.get()
            if tipo in ("pasta", "arquivo"):
                real = caminho_do_usuario(var.get(), base)
                if existe(real):
                    ajustado = os.path.normcase(real) != os.path.normcase(var.get().strip().strip('"'))
                    situacao.config(text="✔ encontrado" + ("  (caminho ajustado para este usuário)"
                                                           if ajustado else ""), fg=COR["verde"])
                else:
                    situacao.config(text="⚠ nao encontrado neste computador", fg=COR["laranja"])
            self._marcar_sujo()
        var.trace_add("write", mudou)
        mudou()

    def _procurar(self, var, tipo, base):
        atual = caminho_do_usuario(var.get(), base)
        inicio = atual if os.path.isdir(atual or "") else os.path.dirname(atual or "")
        inicio = inicio if inicio and os.path.isdir(inicio) else None
        if tipo == "pasta":
            novo = filedialog.askdirectory(parent=self, initialdir=inicio)
        else:
            novo = filedialog.askopenfilename(parent=self, initialdir=inicio,
                                              filetypes=[("Planilhas e PDFs", "*.xlsx *.xlsm *.pdf"),
                                                         ("Todos os arquivos", "*.*")])
        if novo:
            var.set(os.path.normpath(novo))

    def _marcar_sujo(self):
        if not self.doc:
            return
        n = len(self.doc.alterados())
        if hasattr(self, "bt_salvar"):
            self.bt_salvar.ativar(n > 0)
            if n:
                self.aviso_cfg.config(text=f"{n} alteração(ões) não salva(s)", fg=COR["laranja"])

    def _salvar_config(self):
        if not self.doc or not self.doc.alterados():
            return
        try:
            self.doc.salvar()
        except OSError as e:
            messagebox.showerror("Autofiscal", f"Nao consegui salvar:\n{e}")
            return
        self.aviso_cfg.config(text=f"✔ Salvo às {agora()}", fg=COR["verde"])
        self.bt_salvar.ativar(False)
        if self.rodando:
            messagebox.showinfo("Autofiscal", "Salvo. A ferramenta esta rodando: pare e inicie de novo "
                                "para ela usar a configuracao nova.")

    def tem_alteracao(self):
        return bool(self.doc and self.doc.alterados())


# =============================================================================
# Janela principal
# =============================================================================
class App(tk.Tk):
    def __init__(self, raiz):
        super().__init__()
        self.raiz = raiz
        self.title("Autofiscal")
        self.geometry("1180x760")
        self.minsize(960, 620)
        self.configure(bg=COR["fundo"])
        self._estilos()
        self.paginas = {}
        self.ativas = set()
        self.atual = None
        self.itens = {}

        self.lateral = tk.Frame(self, bg=COR["lateral"], width=250)
        self.lateral.pack(side="left", fill="y")
        self.lateral.pack_propagate(False)
        self.conteudo = tk.Frame(self, bg=COR["fundo"])
        self.conteudo.pack(side="left", fill="both", expand=True)

        self._carregar()
        self.protocol("WM_DELETE_WINDOW", self.fechar)
        self.after(100, self._ciclo)

    def _estilos(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure("TNotebook", background=COR["fundo"], borderwidth=0)
        s.configure("TNotebook.Tab", font=(FONTE, 10), padding=(8, 6), background="#e5e7eb")
        s.map("TNotebook.Tab", background=[("selected", COR["cartao"])],
              foreground=[("selected", COR["destaque"])])
        s.configure("TEntry", padding=5)
        s.configure("TCombobox", padding=4)

    # --------------------------------------------------------------- lateral ----
    def _carregar(self):
        for w in self.lateral.winfo_children():
            w.destroy()
        for p in self.paginas.values():
            p.destroy()
        self.paginas, self.itens = {}, {}
        self.ferramentas = carregar_todas(self.raiz)

        tk.Label(self.lateral, text="Autofiscal", bg=COR["lateral"], fg="#ffffff",
                 font=(FONTE, 17, "bold")).pack(anchor="w", padx=20, pady=(22, 0))
        tk.Label(self.lateral, text="Ferramentas do trabalho", bg=COR["lateral"],
                 fg=COR["lateral_txt"], font=(FONTE, 9)).pack(anchor="w", padx=20, pady=(0, 16))
        self._item("inicio", "Início")
        tk.Frame(self.lateral, bg="#475569", height=1).pack(fill="x", padx=16, pady=8)
        for f in self.ferramentas:
            self._item(f.id, f.nome)

        rodape = tk.Frame(self.lateral, bg=COR["lateral"])
        rodape.pack(side="bottom", fill="x", padx=16, pady=14)
        rec = tk.Label(rodape, text="⟳  Recarregar ferramentas", bg=COR["lateral"], fg=COR["lateral_txt"],
                       font=(FONTE, 9), cursor="hand2")
        rec.pack(anchor="w")
        rec.bind("<Button-1>", lambda e: self.recarregar())
        tk.Label(rodape, text=f"versão {VERSAO}  ·  usuário {os.environ.get('USERNAME', '')}",
                 bg=COR["lateral"], fg="#64748b", font=(FONTE, 8)).pack(anchor="w", pady=(6, 0))
        self.mostrar("inicio")

    def _item(self, chave, texto):
        item = tk.Frame(self.lateral, bg=COR["lateral"], cursor="hand2")
        item.pack(fill="x", padx=10, pady=1)
        ponto = tk.Label(item, text="●", bg=COR["lateral"], fg=COR["lateral"], font=(FONTE, 9))
        ponto.pack(side="right", padx=10)
        rot = tk.Label(item, text=texto, bg=COR["lateral"], fg=COR["lateral_txt"],
                       font=(FONTE, 10), anchor="w", padx=12, pady=8)
        rot.pack(side="left", fill="x", expand=True)
        for w in (item, rot, ponto):
            w.bind("<Button-1>", lambda e, c=chave: self.mostrar(c))
        self.itens[chave] = (item, rot, ponto)

    def atualizar_lateral(self):
        for chave, (item, rot, ponto) in self.itens.items():
            sel = chave == self.atual
            cor = COR["lateral_sel"] if sel else COR["lateral"]
            for w in (item, rot):
                w.config(bg=cor)
            rot.config(fg="#ffffff" if sel else COR["lateral_txt"],
                       font=(FONTE, 10, "bold" if sel else "normal"))
            pag = self.paginas.get(chave)
            rodando = bool(pag and isinstance(pag, PaginaFerramenta) and pag.rodando)
            ponto.config(bg=cor, fg="#4ade80" if rodando else cor)

    # --------------------------------------------------------------- paginas ----
    def mostrar(self, chave):
        if self.atual and self.atual != chave:
            pag = self.paginas.get(self.atual)
            if isinstance(pag, PaginaFerramenta) and pag.tem_alteracao():
                r = messagebox.askyesnocancel("Autofiscal", "Há configurações não salvas.\nSalvar agora?")
                if r is None:
                    return
                if r:
                    pag._salvar_config()
        for p in self.paginas.values():
            p.pack_forget()
        if chave not in self.paginas:
            if chave == "inicio":
                self.paginas[chave] = self._pagina_inicio()
            else:
                f = next(x for x in self.ferramentas if x.id == chave)
                self.paginas[chave] = PaginaFerramenta(self.conteudo, self, f)
        self.paginas[chave].pack(fill="both", expand=True)
        self.atual = chave
        self.atualizar_lateral()

    def _pagina_inicio(self):
        pag = tk.Frame(self.conteudo, bg=COR["fundo"])
        tk.Label(pag, text="Bem-vindo ao Autofiscal", bg=COR["fundo"], fg=COR["texto"],
                 font=(FONTE, 20, "bold")).pack(anchor="w", padx=28, pady=(26, 2))
        tk.Label(pag, text="Escolha uma ferramenta. Cada uma tem botões para rodar e uma aba de "
                           "configurações.", bg=COR["fundo"], fg=COR["suave"],
                 font=(FONTE, 10)).pack(anchor="w", padx=28)
        grade = tk.Frame(pag, bg=COR["fundo"])
        grade.pack(fill="both", expand=True, padx=22, pady=18)
        if not self.ferramentas:
            tk.Label(grade, text="Nenhuma ferramenta encontrada. Cada ferramenta precisa de uma "
                                 "subpasta com um arquivo ferramenta.json.", bg=COR["fundo"],
                     fg=COR["laranja"]).pack(anchor="w")
        for i, f in enumerate(self.ferramentas):
            c = tk.Frame(grade, bg=COR["cartao"], highlightbackground=COR["borda"],
                         highlightthickness=1, cursor="hand2")
            c.grid(row=i // 3, column=i % 3, sticky="nsew", padx=6, pady=6)
            t = tk.Label(c, text=f.nome, bg=COR["cartao"], fg=COR["texto"], font=(FONTE, 12, "bold"),
                         anchor="w")
            t.pack(fill="x", padx=14, pady=(12, 2))
            d = tk.Label(c, text=f.descricao or " ", bg=COR["cartao"], fg=COR["suave"], font=(FONTE, 9),
                         wraplength=250, justify="left", anchor="w")
            d.pack(fill="x", padx=14)
            a = tk.Label(c, text="Abrir  →", bg=COR["cartao"], fg=COR["destaque"],
                         font=(FONTE, 10, "bold"), anchor="w")
            a.pack(fill="x", padx=14, pady=(8, 12))
            for w in (c, t, d, a):
                w.bind("<Button-1>", lambda e, k=f.id: self.mostrar(k))
        for col in range(3):
            grade.columnconfigure(col, weight=1, uniform="c")
        return pag

    # ------------------------------------------------------------ execucoes ----
    def registrar(self, pagina):
        self.ativas.add(pagina)

    def _ciclo(self):
        for pag in list(self.ativas):
            try:
                if pag.receber():
                    self.ativas.discard(pag)
            except tk.TclError:
                self.ativas.discard(pag)
        self.after(100, self._ciclo)

    def recarregar(self):
        if any(p.rodando for p in self.paginas.values() if isinstance(p, PaginaFerramenta)):
            messagebox.showinfo("Autofiscal", "Pare as ferramentas que estão rodando antes de recarregar.")
            return
        self.atual = None
        self._carregar()

    def fechar(self):
        rodando = [p for p in self.paginas.values() if isinstance(p, PaginaFerramenta) and p.rodando]
        if rodando:
            nomes = "\n".join(f"• {p.f.nome}" for p in rodando)
            if not messagebox.askyesno("Autofiscal", f"Ainda está rodando:\n{nomes}\n\nParar tudo e sair?"):
                return
            for p in rodando:
                p.execucao.parar()
        self.destroy()


def iniciar(raiz):
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)      # letras nítidas
        except Exception:
            pass
    app = App(raiz)
    app.mainloop()
