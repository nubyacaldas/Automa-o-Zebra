import os
import time
import zipfile
import win32print
from datetime import datetime

# ================= CONFIGURAÇÕES =================
PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
PREFIXO_ZIP = "Etiqueta MercadoEnvios"
IMPRESSORA_ZEBRA = "Etiquetadora Mercado"

PASTA_LOG = "C:\\Automacao_Zebra\\logs"
os.makedirs(PASTA_LOG, exist_ok=True)
# =================================================

def log(mensagem):
    with open(os.path.join(PASTA_LOG, "log.txt"), "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now()}] {mensagem}\n")

def enviar_para_impressora(caminho_txt):
    hPrinter = None
    try:
        hPrinter = win32print.OpenPrinter(IMPRESSORA_ZEBRA)

        job = win32print.StartDocPrinter(
            hPrinter,
            1,
            ("Etiqueta MercadoEnvios", None, "RAW")
        )

        win32print.StartPagePrinter(hPrinter)

        with open(caminho_txt, "rb") as f:
            dados = f.read()  # ZPL puro, sem conversão

        bytes_escritos = win32print.WritePrinter(hPrinter, dados)

        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)

        log(f"Etiqueta enviada. Bytes escritos: {bytes_escritos}")
        return True

    except Exception as e:
        log(f"ERRO AO IMPRIMIR: {e}")
        return False

    finally:
        if hPrinter:
            win32print.ClosePrinter(hPrinter)

log("Automação iniciada.")

arquivos_processados = set()

while True:
    try:
        for arquivo in os.listdir(PASTA_DOWNLOADS):
            if arquivo.startswith(PREFIXO_ZIP) and arquivo.endswith(".zip"):
                caminho_zip = os.path.join(PASTA_DOWNLOADS, arquivo)

                if caminho_zip in arquivos_processados:
                    continue

                with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
                    zip_ref.extractall(PASTA_DOWNLOADS)

                txt_path = os.path.join(PASTA_DOWNLOADS, "Etiqueta de envio.txt")

                if os.path.exists(txt_path):
                    sucesso = enviar_para_impressora(txt_path)

                    # tempo real para o spooler processar
                    time.sleep(5)

                    if sucesso:
                        os.remove(txt_path)
                        os.remove(caminho_zip)
                        arquivos_processados.add(caminho_zip)
                        log("Arquivos removidos após impressão com sucesso.")
                    else:
                        log("Arquivo NÃO removido devido a falha na impressão.")

        time.sleep(2)

    except Exception as e:
        log(f"ERRO GERAL: {e}")
