"""Roda um script de ferramenta em segundo plano e captura o que ele escreve."""

import codecs
import os
import queue
import subprocess
import sys
import threading

from .executor import OPCAO, caminho_executor, instalado

CREATE_NO_WINDOW = 0x08000000


def python_com_console():
    """O painel roda no pythonw (sem janela preta); os scripts rodam no python.exe."""
    exe = sys.executable
    pasta, nome = os.path.split(exe)
    if nome.lower() == "pythonw.exe":
        candidato = os.path.join(pasta, "python.exe")
        if os.path.exists(candidato):
            return candidato
    return exe


def montar_comando(script, args, comando=None):
    """Instalado (.exe) nao ha python.exe: quem roda o script e o executor."""
    if instalado():
        if comando:
            raise OSError("o Autofiscal instalado nao roda comandos do Python (ex.: pip)")
        return [caminho_executor(), OPCAO, script] + list(args)
    if comando:
        return [python_com_console()] + list(comando)
    return [python_com_console(), "-u", script] + list(args)


class Execucao:
    """Um script rodando. O texto chega pela fila `saida` em pedacos."""

    def __init__(self, acao, args, comando=None):
        self.acao = acao
        self.args = args
        self.comando = comando        # comando pronto (ex.: pip), em vez do script
        self.saida = queue.Queue()
        self.proc = None
        self.codigo = None

    @property
    def rodando(self):
        return self.proc is not None and self.proc.poll() is None

    def iniciar(self):
        env = dict(os.environ)
        env.update({"PYTHONIOENCODING": "utf-8", "PYTHONUNBUFFERED": "1",
                    "PYTHONUTF8": "1", "AUTOFISCAL": "1"})
        cmd = montar_comando(self.acao.caminho_script, self.args, self.comando)
        extra = {}
        if os.name == "nt":
            extra["creationflags"] = CREATE_NO_WINDOW
        self.proc = subprocess.Popen(
            cmd, cwd=self.acao.ferramenta.pasta, env=env,
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            **extra)
        threading.Thread(target=self._ler, daemon=True).start()

    def _ler(self):
        decod = codecs.getincrementaldecoder("utf-8")(errors="replace")
        leitor = self.proc.stdout
        while True:
            try:
                pedaco = leitor.read1(4096) if hasattr(leitor, "read1") else leitor.read(1)
            except (OSError, ValueError):
                break
            if not pedaco:
                break
            texto = decod.decode(pedaco)
            if texto:
                self.saida.put(texto.replace("\r\n", "\n"))
        resto = decod.decode(b"", final=True)
        if resto:
            self.saida.put(resto)
        self.codigo = self.proc.wait()
        self.saida.put(None)          # aviso de fim

    def responder(self, texto):
        if not self.rodando:
            return False
        try:
            self.proc.stdin.write((texto + "\n").encode("utf-8"))
            self.proc.stdin.flush()
            return True
        except (OSError, ValueError):
            return False

    def parar(self):
        """Pede para o script parar sem travar a janela: a espera e o kill de
        reserva ficam numa thread."""
        if not self.rodando:
            return
        try:
            self.proc.terminate()
        except OSError:
            pass
        threading.Thread(target=self._garantir_fim, daemon=True).start()

    def _garantir_fim(self):
        try:
            self.proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            try:
                self.proc.kill()
            except OSError:
                pass
