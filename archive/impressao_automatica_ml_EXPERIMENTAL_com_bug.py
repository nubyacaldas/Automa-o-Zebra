import os
import time
import zipfile
import logging
import win32print
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path
from typing import List
import csv
import re
from PyPDF2 import PdfReader

# ================= CONFIGURAÇÕES =================
PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
PREFIXO_ZIP = "Etiqueta MercadoEnvios"
IMPRESSORA_ZEBRA = "Etiquetadora Mercado"
NOME_ARQUIVO_TXT = "Etiqueta de envio.txt"
NOME_PDF = "Controle.pdf"

PASTA_LOG = "C:\\Automacao_Zebra\\logs"
Path(PASTA_LOG).mkdir(parents=True, exist_ok=True)

ARQUIVO_LOG_ETIQUETAS = os.path.join(PASTA_LOG, "etiquetas.csv")

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(PASTA_LOG, "log.txt"), encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(_name_)

IMPRESSORA_SELECIONADA = IMPRESSORA_ZEBRA
PADRAO_ETIQUETA = ""

# ================= UTIL =================
def arquivo_pronto(caminho, tentativas=5, intervalo=1):
    if not os.path.exists(caminho):
        return False

    tamanho_anterior = -1

    for _ in range(tentativas):
        try:
            tamanho_atual = os.path.getsize(caminho)
            if tamanho_atual == tamanho_anterior:
                return True
            tamanho_anterior = tamanho_atual
            time.sleep(intervalo)
        except OSError:
            return False

    return False


def obter_impressoras() -> List[str]:
    try:
        impressoras = win32print.EnumPrinters(2)
        return [i[2] for i in impressoras]
    except Exception as e:
        logger.error(f"Erro ao listar impressoras: {e}")
        return []

# ================= EXTRAÇÃO =================
def extrair_dados_pdf(caminho_pdf):
    try:
        reader = PdfReader(caminho_pdf)
        texto = ""

        for page in reader.pages:
            texto += page.extract_text()

        pedido_ml = re.search(r"Venda:\s*(\d+)", texto)

        return pedido_ml.group(1) if pedido_ml else ""

    except Exception as e:
        logger.warning(f"Erro ao ler PDF: {e}")
        return ""


def extrair_dados_zpl(caminho_txt):
    try:
        with open(caminho_txt, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()

        # NF
        nf = re.search(r"NF:\s*(\d+)", conteudo)

        # Destinatário
        destinatario = re.search(r"\^FD([A-Z\s]+\([A-Z0-9]+\))", conteudo)

        # Data despacho
        despacho = re.search(r"Despachar:\s*(.*?)\^FS", conteudo)

        return {
            "nf": nf.group(1) if nf else "",
            "destinatario": destinatario.group(1) if destinatario else "",
            "despacho": despacho.group(1) if despacho else ""
        }

    except Exception as e:
        logger.warning(f"Erro ao extrair ZPL: {e}")
        return {"nf": "", "destinatario": "", "despacho": ""}

# ================= LOG =================
def registrar_log(dados):
    existe = os.path.exists(ARQUIVO_LOG_ETIQUETAS)

    with open(ARQUIVO_LOG_ETIQUETAS, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        if not existe:
            writer.writerow([
                "data_hora",
                "pedido_ml",
                "nf",
                "destinatario",
                "data_despacho"
            ])

        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            dados["pedido_ml"],
            dados["nf"],
            dados["destinatario"],
            dados["despacho"]
        ])

# ================= UI =================
class JanelaConfiguracao:
    def _init_(self, root):
        self.root = root
        self.root.title("Configuração")

        self.listbox = tk.Listbox(root)
        self.listbox.pack(fill=tk.BOTH, expand=True)

        for i in obter_impressoras():
            self.listbox.insert(tk.END, i)

        tk.Button(root, text="Confirmar", command=self.confirmar).pack()

        self.impressora_selecionada = None

    def confirmar(self):
        selecao = self.listbox.curselection()
        if not selecao:
            messagebox.showerror("Erro", "Selecione uma impressora")
            return

        self.impressora_selecionada = self.listbox.get(selecao[0])
        self.root.destroy()


def abrir_janela_configuracao():
    root = tk.Tk()
    janela = JanelaConfiguracao(root)
    root.mainloop()
    return janela.impressora_selecionada or IMPRESSORA_ZEBRA

# ================= IMPRESSÃO =================
def enviar_para_impressora(caminho_txt):
    try:
        hPrinter = win32print.OpenPrinter(IMPRESSORA_SELECIONADA)

        win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta", None, "RAW"))

        with open(caminho_txt, "rb") as f:
            dados = f.read()

        texto = dados.decode("utf-8", errors="ignore")
        texto = texto.replace("^PQ2", "^PQ1")
        dados = texto.encode("utf-8")

        win32print.WritePrinter(hPrinter, dados)
        win32print.EndDocPrinter(hPrinter)

        return True

    except Exception as e:
        logger.error(f"Erro impressão: {e}")
        return False

    finally:
        try:
            win32print.ClosePrinter(hPrinter)
        except:
            pass

# ================= LOOP =================
arquivos_processados = set()


def processar_fila():
    for arquivo in os.listdir(PASTA_DOWNLOADS):

        if not (arquivo.startswith(PREFIXO_ZIP) and arquivo.endswith(".zip")):
            continue

        caminho_zip = os.path.join(PASTA_DOWNLOADS, arquivo)

        if caminho_zip in arquivos_processados:
            continue

        if not arquivo_pronto(caminho_zip):
            continue

        try:
            with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
                zip_ref.extractall(PASTA_DOWNLOADS)

            txt_path = os.path.join(PASTA_DOWNLOADS, NOME_ARQUIVO_TXT)
            pdf_path = os.path.join(PASTA_DOWNLOADS, NOME_PDF)

            if not os.path.exists(txt_path):
                continue

            pedido_ml = extrair_dados_pdf(pdf_path)
            dados_zpl = extrair_dados_zpl(txt_path)

            dados = {
                "pedido_ml": pedido_ml,
                "nf": dados_zpl["nf"],
                "destinatario": dados_zpl["destinatario"],
                "despacho": dados_zpl["despacho"]
            }

            sucesso = enviar_para_impressora(txt_path)

            if sucesso:
                registrar_log(dados)

                os.remove(txt_path)
                os.remove(caminho_zip)
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)

                arquivos_processados.add(caminho_zip)

        except Exception as e:
            logger.error(f"Erro geral: {e}")


if _name_ == "_main_":
    logger.info("Iniciando automação...")

    IMPRESSORA_SELECIONADA = abrir_janela_configuracao()

    while True:
        processar_fila()
        time.sleep(2)