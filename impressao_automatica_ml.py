import os
import time
import zipfile
import win32print
from datetime import datetime

# ================= CONFIGURAÇÕES =================
PASTA_DOWNLOADS = r"C:\Users\mercado\Downloads"
PALAVRA_CHAVE_ZIP = "MercadoEnvios"
IMPRESSORA_ZEBRA = "Etiquetadora Mercado"

PASTA_LOG = r"C:\Automacao_Zebra\logs"
os.makedirs(PASTA_LOG, exist_ok=True)

INTERVALO_LOOP = 2  # segundos
# =================================================


def log(mensagem):
    with open(os.path.join(PASTA_LOG, "log.txt"), "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now()}] {mensagem}\n")


def arquivo_pronto(caminho, tentativas=5, intervalo=1):
    """
    Verifica se o arquivo existe e se o tamanho permanece estável
    """
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


def enviar_para_impressora(caminho_txt):
    hPrinter = None
    try:
        hPrinter = win32print.OpenPrinter(IMPRESSORA_ZEBRA)

        win32print.StartDocPrinter(
            hPrinter,
            1,
            ("Etiqueta MercadoEnvios", None, "RAW")
        )

        with open(caminho_txt, "rb") as f:
            dados = f.read()

        # ========= CORREÇÃO DE DUPLICIDADE (^PQ) =========
        try:
            texto = dados.decode("utf-8", errors="ignore")
            texto = texto.replace("^PQ2", "^PQ1")
            texto = texto.replace("^PQ,2", "^PQ,1")
            texto = texto.replace("^PQ 2", "^PQ 1")
            dados = texto.encode("utf-8")
        except Exception as e:
            log(f"Aviso: falha ao normalizar ^PQ: {e}")
        # =================================================

        bytes_escritos = win32print.WritePrinter(hPrinter, dados)

        win32print.EndDocPrinter(hPrinter)

        log(f"Etiqueta impressa com sucesso ({bytes_escritos} bytes).")
        return True

    except Exception as e:
        log(f"ERRO AO IMPRIMIR: {e}")
        return False

    finally:
        if hPrinter:
            win32print.ClosePrinter(hPrinter)


log("Automação iniciada.")

while True:
    try:
        for arquivo in os.listdir(PASTA_DOWNLOADS):

            if not arquivo.lower().endswith(".zip"):
                continue

            if PALAVRA_CHAVE_ZIP.lower() not in arquivo.lower():
                continue

            caminho_zip = os.path.join(PASTA_DOWNLOADS, arquivo)

            # garante que o ZIP terminou de baixar
            if not arquivo_pronto(caminho_zip):
                continue

            try:
                with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
                    zip_ref.extractall(PASTA_DOWNLOADS)
            except zipfile.BadZipFile:
                log(f"ZIP inválido ou incompleto: {arquivo}")
                continue
            except FileNotFoundError:
                log(f"ZIP não encontrado no momento da extração: {arquivo}")
                continue

            # procura o TXT mais recente extraído
            txt_path = None
            arquivos_txt = [
                os.path.join(PASTA_DOWNLOADS, f)
                for f in os.listdir(PASTA_DOWNLOADS)
                if f.lower().endswith(".txt")
            ]

            if arquivos_txt:
                txt_path = max(arquivos_txt, key=os.path.getmtime)

            if txt_path and os.path.exists(txt_path):
                sucesso = enviar_para_impressora(txt_path)

                if sucesso:
                    try:
                        os.remove(txt_path)
                        os.remove(caminho_zip)
                        log("TXT e ZIP removidos após impressão.")
                    except Exception as e:
                        log(f"Falha ao remover arquivos: {e}")
                else:
                    log("Falha na impressão. Arquivos mantidos.")
            else:
                log("Nenhum arquivo TXT encontrado após extração.")

        time.sleep(INTERVALO_LOOP)

    except Exception as e:
        log(f"ERRO GERAL: {e}"+"teste")