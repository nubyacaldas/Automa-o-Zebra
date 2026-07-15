import os
import time
import zipfile
import logging
import win32print
from datetime import datetime
import re
import csv

# ================= CONFIG =================
PASTA_DOWNLOADS = r"C:\Users\mercado\Downloads"
PREFIXO_ZIP = "Etiqueta MercadoEnvios"

PASTA_LOG = r"C:\Automacao_Zebra\logs"
os.makedirs(PASTA_LOG, exist_ok=True)

ARQUIVO_LOG = os.path.join(PASTA_LOG, "etiquetas.csv")

IMPRESSORA = "Etiquetadora Mercado"

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(PASTA_LOG, "log.txt"), encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# ================= UTIL =================
def arquivo_pronto(caminho):
    tamanho_antigo = -1

    for _ in range(5):
        if not os.path.exists(caminho):
            return False

        tamanho = os.path.getsize(caminho)

        if tamanho == tamanho_antigo:
            return True

        tamanho_antigo = tamanho
        time.sleep(1)

    return False


# ================= EXTRAÇÃO =================
def extrair_dados_zpl(caminho_txt):
    try:
        with open(caminho_txt, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()

        # -------- PEDIDO (PACK ID COMPLETO) --------
        pack_prefix = re.search(r"Pack ID:\s*(\d+)", conteudo)
        numeros = re.findall(r"\^FD(\d{8,})\^FS", conteudo)

        pedido_ml = ""
        if pack_prefix and numeros:
            for num in numeros:
                if len(num) >= 8:
                    pedido_ml = pack_prefix.group(1) + num
                    break

        # -------- NF --------
        nf_match = re.search(r"NF:\s*(\d+)", conteudo)
        nf = nf_match.group(1) if nf_match else ""

        # -------- DESPACHO --------
        despacho_match = re.search(r"Despachar:\s*(.*?)\^FS", conteudo)
        despacho = despacho_match.group(1) if despacho_match else ""

        # -------- DESTINATÁRIO (CORRETO) --------
        destinatario = ""

        bloco_match = re.search(r"\^FX  RECEIVER ZONE(.*?)\^FX", conteudo, re.DOTALL)

        if bloco_match:
            bloco = bloco_match.group(1)

            match = re.search(r"\^FD(.+?)\^FS", bloco)
            if match:
                destinatario = match.group(1).strip()

        return {
            "pedido_ml": pedido_ml,
            "nf": nf,
            "destinatario": destinatario,
            "despacho": despacho
        }

    except Exception as e:
        logging.error(f"Erro ao extrair dados: {e}")
        return {
            "pedido_ml": "",
            "nf": "",
            "destinatario": "",
            "despacho": ""
        }


# ================= LOG CSV =================
def registrar_log(dados):
    existe = os.path.exists(ARQUIVO_LOG)

    with open(ARQUIVO_LOG, "a", newline="", encoding="utf-8") as f:
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


# ================= IMPRESSÃO =================
def imprimir(caminho_txt):
    try:
        hPrinter = win32print.OpenPrinter(IMPRESSORA)

        win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta", None, "RAW"))

        with open(caminho_txt, "rb") as f:
            dados = f.read()

        # evita duplicidade
        texto = dados.decode("utf-8", errors="ignore")
        texto = texto.replace("^PQ2", "^PQ1")
        dados = texto.encode("utf-8")

        win32print.WritePrinter(hPrinter, dados)
        win32print.EndDocPrinter(hPrinter)

        return True

    except Exception as e:
        logging.error(f"Erro impressão: {e}")
        return False

    finally:
        try:
            win32print.ClosePrinter(hPrinter)
        except:
            pass


# ================= LOOP =================
def processar():
    for arquivo in os.listdir(PASTA_DOWNLOADS):

        if not arquivo.endswith(".zip"):
            continue

        if PREFIXO_ZIP not in arquivo:
            continue

        caminho_zip = os.path.join(PASTA_DOWNLOADS, arquivo)

        if not arquivo_pronto(caminho_zip):
            continue

        logging.info(f"Processando: {arquivo}")

        try:
            with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
                zip_ref.extractall(PASTA_DOWNLOADS)

            txt_path = os.path.join(PASTA_DOWNLOADS, "Etiqueta de envio.txt")

            if not os.path.exists(txt_path):
                logging.error("TXT não encontrado")
                continue

            dados = extrair_dados_zpl(txt_path)

            sucesso = imprimir(txt_path)

            if sucesso:
                registrar_log(dados)

                os.remove(txt_path)
                os.remove(caminho_zip)

                logging.info("Etiqueta processada e registrada")

        except Exception as e:
            logging.error(f"Erro geral: {e}")


# ================= MAIN =================
if __name__ == "__main__":
    logging.info("Iniciando automação...")

    while True:
        processar()
        time.sleep(2)