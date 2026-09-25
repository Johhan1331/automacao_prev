from datetime import date, datetime, timedelta
import logging
import os
from pathlib import Path
import subprocess
import sys
import threading
import tkinter as tk
import unicodedata
from tkinter import messagebox, ttk


PASTA_RAIZ = Path(r"X:\Comum-PagamentosDiarios")
PASTA_PROJETO = Path(__file__).resolve().parent
PASTA_RELATORIOS = PASTA_PROJETO / "relatorios_sap"
PASTA_DIRETORIO = PASTA_PROJETO / "manipular_diretorio"

MESES = [
    "01 - JANEIRO", "02 - FEVEREIRO", "03 - MARÇO",
    "04 - ABRIL", "05 - MAIO", "06 - JUNHO",
    "07 - JULHO", "08 - AGOSTO", "09 - SETEMBRO",
    "10 - OUTUBRO", "11 - NOVEMBRO", "12 - DEZEMBRO"
]

COR_FUNDO = "#1b3157"
COR_DESTAQUE = "#5fbb48"
SUFIXO_EMAIL = "@icatuseguros.com.br"


class AutomacaoCancelada(Exception):
    pass


class Aplicacao:

    def __init__(self, janela):
        self.janela = janela
        self.janela.title("Automação de Relatórios SAP")
        self.janela.configure(bg=COR_FUNDO)
        self.janela.resizable(False, False)
        self.janela.protocol("WM_DELETE_WINDOW", self.fechar)

        self.email = tk.StringVar(value=SUFIXO_EMAIL)
        self.senha = tk.StringVar()
        self.data_inicial = tk.StringVar()
        self.data_final = tk.StringVar()
        self.status = tk.StringVar(value="Preencha os dados para iniciar.")
        self.executando = False
        self.cancelando = False
        self.processo = None

        self._criar_tela()

    def _criar_tela(self):
        estilo = ttk.Style(self.janela)
        estilo.configure(
            "Tela.TFrame",
            background=COR_FUNDO
        )
        estilo.configure(
            "Titulo.TLabel",
            background=COR_FUNDO,
            foreground=COR_DESTAQUE
        )
        estilo.configure(
            "Campo.TLabel",
            background=COR_FUNDO,
            foreground=COR_DESTAQUE,
            font=("Segoe UI", 11, "bold")
        )
        estilo.configure(
            "Status.TLabel",
            background=COR_FUNDO,
            foreground=COR_DESTAQUE
        )
        estilo.configure(
            "Campo.TEntry",
            foreground="black",
            fieldbackground="white"
        )
        estilo.configure(
            "Acao.TButton",
            background=COR_DESTAQUE,
            foreground="black"
        )

        principal = ttk.Frame(
            self.janela,
            padding=24,
            style="Tela.TFrame"
        )
        principal.grid(row=0, column=0, sticky="nsew")

        ttk.Label(
            principal,
            text="OPERAÇÕES PREV",
            font=("Segoe UI", 14, "bold"),
            style="Titulo.TLabel"
        ).grid(row=0, column=0, columnspan=2, pady=(0, 2))

        ttk.Label(
            principal,
            text="EXTRAÇÃO DE RELATÓRIOS DE PAGAMENTO - SAP",
            font=("Segoe UI", 11, "bold"),
            style="Titulo.TLabel"
        ).grid(row=1, column=0, columnspan=2, pady=(0, 18))

        ttk.Label(
            principal,
            text="Email",
            style="Campo.TLabel"
        ).grid(row=2, column=0, columnspan=2, pady=(0, 2))
        self.email_entry = ttk.Entry(
            principal,
            textvariable=self.email,
            width=32,
            style="Campo.TEntry"
        )
        self.email_entry.grid(
            row=3,
            column=0,
            columnspan=2,
            pady=(0, 8)
        )
        self.email_entry.bind(
            "<FocusIn>",
            self._posicionar_sufixo_email
        )
        self.email_entry.bind(
            "<KeyRelease>",
            self._manter_sufixo_email
        )

        ttk.Label(
            principal,
            text="Senha",
            style="Campo.TLabel"
        ).grid(row=4, column=0, columnspan=2, pady=(0, 2))
        self.senha_entry = ttk.Entry(
            principal,
            textvariable=self.senha,
            width=32,
            show="*",
            style="Campo.TEntry"
        )
        self.senha_entry.grid(
            row=5,
            column=0,
            columnspan=2,
            pady=(0, 10)
        )

        datas = ttk.Frame(principal, style="Tela.TFrame")
        datas.grid(row=6, column=0, columnspan=2, sticky="ew")
        datas.columnconfigure(0, weight=1)
        datas.columnconfigure(1, weight=1)

        ttk.Label(
            datas,
            text="Data inicial",
            style="Campo.TLabel"
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(
            datas,
            text="Data final",
            style="Campo.TLabel"
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.data_inicial_entry = ttk.Entry(
            datas,
            textvariable=self.data_inicial,
            width=12,
            style="Campo.TEntry"
        )
        self.data_inicial_entry.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=(0, 8),
            pady=(2, 0)
        )
        self.data_final_entry = ttk.Entry(
            datas,
            textvariable=self.data_final,
            width=12,
            style="Campo.TEntry"
        )
        self.data_final_entry.grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(8, 0),
            pady=(2, 0)
        )

        botoes = ttk.Frame(principal, style="Tela.TFrame")
        botoes.grid(row=7, column=0, columnspan=2, pady=(18, 8))

        self.botao = ttk.Button(
            botoes,
            text="Iniciar automação",
            command=self.iniciar,
            style="Acao.TButton"
        )
        self.botao.grid(row=0, column=0, padx=5)

        self.botao_cancelar = ttk.Button(
            botoes,
            text="Cancelar execução",
            command=self.cancelar,
            state="disabled",
            style="Acao.TButton"
        )
        self.botao_cancelar.grid(row=0, column=1, padx=5)

        ttk.Label(
            principal,
            textvariable=self.status,
            wraplength=360,
            style="Status.TLabel"
        ).grid(row=8, column=0, columnspan=2, pady=(0, 10))

        self.log_texto = tk.Text(
            principal,
            width=72,
            height=14,
            state="disabled",
            wrap="word",
            background="white",
            foreground="black"
        )
        self.log_texto.grid(row=9, column=0, columnspan=2)

    def _posicionar_sufixo_email(self, _evento=None):
        self.email_entry.icursor(
            max(0, len(self.email.get()) - len(SUFIXO_EMAIL))
        )

    def _manter_sufixo_email(self, _evento=None):
        valor = self.email.get()
        parte_usuario = valor.split(SUFIXO_EMAIL, 1)[0]
        self.email.set(parte_usuario + SUFIXO_EMAIL)
        self.email_entry.icursor(len(parte_usuario))

    def _validar(self):
        parte_usuario = self.email.get().removesuffix(SUFIXO_EMAIL).strip()
        if not parte_usuario:
            self.email_entry.focus_set()
            raise ValueError("Informe o e-mail.")
        if not self.senha.get():
            self.senha_entry.focus_set()
            raise ValueError("Informe a senha.")

        valores = []
        campos_data = (
              ("data inicial", self.data_inicial_entry, self.data_inicial.get()),
              ("data final", self.data_final_entry, self.data_final.get()),
        )
        for nome, campo, valor in campos_data:
            try:
                data = datetime.strptime(valor.strip(), "%d.%m.%Y")
            except ValueError as erro:
                campo.focus_set()
                raise ValueError(
                    f"A {nome} deve usar o formato DD.MM.AAAA."
                ) from erro

            if data.date() > date.today():
                campo.focus_set()
                raise ValueError(
                    f"A {nome} não pode ser posterior a hoje."
                )

            valores.append(data)

        if valores[0] > valores[1]:
            self.data_final_entry.focus_set()
            raise ValueError("A data inicial não pode ser posterior à data final.")

    def _registrar(self, mensagem):
        texto = "".join(
            caractere
            for caractere in unicodedata.normalize("NFKD", mensagem)
            if not unicodedata.combining(caractere)
        )
        texto = texto.encode("ascii", "ignore").decode("ascii").rstrip()
        if not texto:
            return

        self.janela.after(0, self._mostrar_log, texto)
        if hasattr(self, "logger"):
            self.logger.info(texto)

    def _mostrar_log(self, mensagem):
        self.log_texto.configure(state="normal")
        self.log_texto.insert("end", mensagem + "\n")
        self.log_texto.see("end")
        self.log_texto.configure(state="disabled")

    def _configurar_log(self, pasta_manha):
        pasta_manha.mkdir(parents=True, exist_ok=True)
        caminho_log = pasta_manha / (
            f"LOG_AUTOMACAO_{datetime.now():%Y%m%d_%H%M%S}.log"
        )

        self.logger = logging.getLogger("automacao")
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False
        self.logger.handlers.clear()
        manipulador = logging.FileHandler(caminho_log, encoding="utf-8")
        manipulador.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%d/%m/%Y %H:%M:%S"
        ))
        self.logger.addHandler(manipulador)
        self._registrar(f"Arquivo de log: {caminho_log}")

    def iniciar(self):
        if self.executando:
            return

        try:
            self._validar()
        except ValueError as erro:
            messagebox.showerror("Dados inválidos", str(erro))
            return

        self.executando = True
        self.cancelando = False
        self.botao.configure(state="disabled")
        self.botao_cancelar.configure(state="normal")
        self.status.set("Automação em execução...")
        threading.Thread(target=self._executar, daemon=True).start()

    def cancelar(self):
        if not self.executando or self.cancelando:
            return

        self.cancelando = True
        self.status.set("Cancelando execução...")
        self.botao_cancelar.configure(state="disabled")

        processo = self.processo
        if processo is not None and processo.poll() is None:
            subprocess.run(
                ["taskkill", "/PID", str(processo.pid), "/T", "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False
            )

    def _executar_script(self, script, ambiente):
        self._registrar(f"Iniciando: {script.name}")
        processo = subprocess.Popen(
            [sys.executable, str(script)],
            cwd=str(script.parent),
            env=ambiente,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1
        )
        self.processo = processo
        saida = []

        for linha in processo.stdout:
            saida.append(linha)
            self._registrar(linha)

        codigo = processo.wait()
        self.processo = None

        if self.cancelando:
            raise AutomacaoCancelada()

        if (
            codigo != 0
            and script.name == "extracao_sap.py"
            and any("Senha recusada pelo SAP" in linha for linha in saida)
        ):
            raise RuntimeError(
                "Não foi possível entrar no SAP. "
                "Verifique o e-mail e a senha."
            )

        if codigo != 0:
            raise RuntimeError(
                f"A etapa {script.name} terminou com código {codigo}."
            )

        self._registrar(f"Concluído: {script.name}")

    def _executar(self):
        hoje = date.today()
        data_pasta = hoje + timedelta(days=1)
        pasta_raiz_teste = PASTA_RAIZ / "AUTOMACAO_TESTE"
        pasta_manha = (
            pasta_raiz_teste
            / str(data_pasta.year)
            / MESES[data_pasta.month - 1]
            / data_pasta.strftime("%d.%m.%Y")
            / "MANHÃ"
        )
        self._configurar_log(pasta_manha)

        ambiente = os.environ.copy()
        ambiente.update({
            "AUTOMACAO_EMAIL": self.email.get().strip(),
            "AUTOMACAO_SENHA": self.senha.get(),
            "AUTOMACAO_DATA_INICIAL": self.data_inicial.get().strip(),
            "AUTOMACAO_DATA_FINAL": self.data_final.get().strip(),
            "AUTOMACAO_PASTA_RAIZ": str(pasta_raiz_teste),
            "AUTOMACAO_DATA_PASTA": data_pasta.isoformat(),
            "AUTOMACAO_PASTA_ORIGEM": str(PASTA_RAIZ),
        })

        etapas = [
            PASTA_DIRETORIO / "abrir_pasta.py",
            PASTA_DIRETORIO / "copiar_planilha.py",
            PASTA_RELATORIOS / "extracao_sap.py",
            PASTA_RELATORIOS / "config_excel.py",
        ]

        try:
            self._registrar("Início da automação")
            for etapa in etapas:
                self._executar_script(etapa, ambiente)
            self._registrar("AUTOMAÇÃO FINALIZADA COM SUCESSO")
            self.janela.after(0, self.status.set, "Automação finalizada com sucesso.")
        except AutomacaoCancelada:
            self._registrar("AUTOMAÇÃO CANCELADA PELO USUÁRIO")
            self.janela.after(0, self.status.set, "Automação cancelada.")
        except Exception as erro:
            self._registrar(f"ERRO: {erro}")
            self.janela.after(0, self.status.set, "A automação foi interrompida com erro.")
            if "Não foi possível entrar no SAP" in str(erro):
                self.janela.after(0, self.senha_entry.focus_set)
                self.janela.after(
                    0,
                    messagebox.showwarning,
                    "Dados do SAP",
                    "Não foi possível entrar no SAP. "
                    "Corrija o e-mail ou a senha e tente novamente."
                )
            else:
                self.janela.after(
                    0,
                    messagebox.showerror,
                    "Erro na automação",
                    str(erro)
                )
        finally:
            self.executando = False
            self.processo = None
            self.janela.after(0, self.botao.configure, {"state": "normal"})
            self.janela.after(0, self.botao_cancelar.configure, {"state": "disabled"})

    def fechar(self):
        if self.executando:
            messagebox.showwarning(
                "Automação em execução",
                "Aguarde a conclusão antes de fechar a aplicação."
            )
            return
        self.janela.destroy()


if __name__ == "__main__":
    janela = tk.Tk()
    Aplicacao(janela)
    janela.mainloop()
