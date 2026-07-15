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
ARQUIVO_PROCESSADOS = os.path.join(PASTA_LOG, "processados.txt")

IMPRESSORA = "Etiquetadora Mercado"

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(PASTA_LOG, "log.txt"), encoding="utf-8"),
        logging.StreamHandler()
    ]
)


# ================= CONTROLE DE PROCESSADOS =================
def carregar_processados():
    """Retorna um set com os nomes de ZIPs já processados com sucesso."""
    if not os.path.exists(ARQUIVO_PROCESSADOS):
        return set()
    with open(ARQUIVO_PROCESSADOS, "r", encoding="utf-8") as f:
        return set(linha.strip() for linha in f if linha.strip())


def registrar_processado(nome_arquivo):
    """Salva o nome do ZIP no controle para não reprocessar."""
    with open(ARQUIVO_PROCESSADOS, "a", encoding="utf-8") as f:
        f.write(nome_arquivo + "\n")


# ================= UTIL =================
def arquivo_pronto(caminho):
    """Aguarda o arquivo parar de crescer (download concluído)."""
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


def deletar_arquivo(caminho):
    """Tenta deletar um arquivo e loga se falhar."""
    try:
        if os.path.exists(caminho):
            os.remove(caminho)
            return True
    except Exception as e:
        logging.error(f"Não foi possível deletar {caminho}: {e}")
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

        # -------- DESTINATÁRIO --------
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
        return {"pedido_ml": "", "nf": "", "destinatario": "", "despacho": ""}


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
    hPrinter = None
    try:
        with open(caminho_txt, "rb") as f:
            dados = f.read()

        # CORREÇÃO PRINCIPAL: regex para capturar ^PQ em qualquer formato (^PQ2, ^PQ02, ^PQ2 , etc.)
        texto = dados.decode("utf-8", errors="ignore")
        texto = re.sub(r"\^PQ\d+", "^PQ1", texto)
        dados = texto.encode("utf-8")

        hPrinter = win32print.OpenPrinter(IMPRESSORA)
        win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta", None, "RAW"))
        win32print.WritePrinter(hPrinter, dados)
        win32print.EndDocPrinter(hPrinter)

        # Aguarda o spooler aceitar e fechar o job antes de seguir
        time.sleep(3)

        logging.info("Job enviado à impressora com sucesso.")
        return True

    except Exception as e:
        logging.error(f"Erro impressão: {e}")
        return False

    finally:
        if hPrinter:
            try:
                win32print.ClosePrinter(hPrinter)
            except:
                pass


# ================= LOOP =================
def processar(ja_processados):
    for arquivo in os.listdir(PASTA_DOWNLOADS):

        if not arquivo.endswith(".zip"):
            continue

        if PREFIXO_ZIP not in arquivo:
            continue

        # CORREÇÃO: pula arquivos que já foram processados com sucesso
        if arquivo in ja_processados:
            continue

        caminho_zip = os.path.join(PASTA_DOWNLOADS, arquivo)

        if not arquivo_pronto(caminho_zip):
            logging.warning(f"Arquivo ainda não está pronto: {arquivo}")
            continue

        logging.info(f"Processando: {arquivo}")

        txt_path = None

        try:
            # CORREÇÃO: descobre o nome do TXT direto do ZIP, sem assumir nome fixo
            with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
                nomes = zip_ref.namelist()
                txts = [n for n in nomes if n.lower().endswith(".txt")]

                if not txts:
                    logging.error(f"Nenhum TXT encontrado dentro de {arquivo}")
                    continue

                zip_ref.extractall(PASTA_DOWNLOADS)
                txt_path = os.path.join(PASTA_DOWNLOADS, txts[0])

            if not os.path.exists(txt_path):
                logging.error(f"TXT extraído não encontrado: {txt_path}")
                continue

            dados = extrair_dados_zpl(txt_path)
            sucesso = imprimir(txt_path)

            if sucesso:
                registrar_log(dados)
                registrar_processado(arquivo)
                ja_processados.add(arquivo)

                # CORREÇÃO: tenta deletar e loga se falhar, sem travar o fluxo
                deletar_arquivo(txt_path)
                deletar_arquivo(caminho_zip)

                logging.info(f"Etiqueta processada: pedido {dados['pedido_ml']} | NF {dados['nf']} | {dados['destinatario']}")
            else:
                logging.error(f"Impressão falhou para {arquivo}. Arquivos mantidos para nova tentativa.")

        except zipfile.BadZipFile:
            logging.error(f"ZIP corrompido ou inválido: {arquivo}")

        except Exception as e:
            logging.error(f"Erro geral ao processar {arquivo}: {e}")

            # Garante limpeza do TXT mesmo em erro, para não vazar entre ZIPs
            if txt_path:
                deletar_arquivo(txt_path)


# ================= MAIN =================
if __name__ == "__main__":
    logging.info("Iniciando automação...")
    ja_processados = carregar_processados()

    while True:
        processar(ja_processados)
        time.sleep(2)