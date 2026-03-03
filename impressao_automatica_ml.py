import os
import time
import zipfile
import logging
import win32print
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path
from typing import Optional, List

# ================= CONFIGURAÇÕES =================
PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
PREFIXO_ZIP = "Etiqueta MercadoEnvios"
IMPRESSORA_ZEBRA = "Etiquetadora Mercado"
NOME_ARQUIVO_TXT = "Etiqueta de envio.txt"  # Configurável

PASTA_LOG = "C:\\Automacao_Zebra\\logs"
Path(PASTA_LOG).mkdir(parents=True, exist_ok=True)

# Padrão de nome para as etiquetas (pode conter tokens como {date}, {time}, {datetime})
PADRAO_ETIQUETA = ""  # Será preenchido pela interface

# Limite de arquivos em cache para evitar vazamento de memória
MAX_ARQUIVOS_CACHE = 1000

# Configurar logging profissional
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(PASTA_LOG, "log.txt"), encoding="utf-8"),
        logging.StreamHandler()  # Também exibe no console
    ]
)
logger = logging.getLogger(__name__)
# =================================================

# Variável global para armazenar a impressora selecionada
IMPRESSORA_SELECIONADA = IMPRESSORA_ZEBRA


def obter_impressoras() -> List[str]:
    """
    Lista todas as impressoras disponíveis no computador.
    
    Returns:
        Lista com nomes das impressoras
    """
    try:
        impressoras = win32print.EnumPrinters(2)
        return [impressora[2] for impressora in impressoras]
    except Exception as e:
        logger.error(f"Erro ao listar impressoras: {e}")
        return []


class JanelaConfiguracao:
    """Janela gráfica para seleção de impressora."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Configuração - Automação Zebra")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        # Estilo
        style = ttk.Style()
        style.theme_use('clam')
        
        # Frame principal
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(
            main_frame, 
            text="Configuração da Automação",
            font=("Arial", 14, "bold")
        )
        titulo.pack(pady=(0, 20))
        
        # Label para impressora
        label_impressora = ttk.Label(
            main_frame,
            text="Selecione a Impressora Zebra:",
            font=("Arial", 10)
        )
        label_impressora.pack(anchor="w", pady=(0, 10))
        
        # Listbox com scrollbar
        frame_lista = ttk.Frame(main_frame)
        frame_lista.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        scrollbar = ttk.Scrollbar(frame_lista)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox_impressoras = tk.Listbox(
            frame_lista,
            yscrollcommand=scrollbar.set,
            font=("Arial", 10),
            height=10
        )
        self.listbox_impressoras.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox_impressoras.yview)
        
        # Carregar impressoras
        self.carregar_impressoras()
        
        # Campo para o padrão de nome da etiqueta
        label_padrao = ttk.Label(
            main_frame,
            text="Padrão do nome da etiqueta (ex: Etiqueta_{date}):",
            font=("Arial", 10)
        )
        label_padrao.pack(anchor="w", pady=(10, 5))

        self.entry_padrao = ttk.Entry(main_frame, font=("Arial", 10))
        self.entry_padrao.pack(fill=tk.X)

        # Frame de botões
        frame_botoes = ttk.Frame(main_frame)
        frame_botoes.pack(fill=tk.X, pady=(10, 0))
        
        botao_confirmar = ttk.Button(
            frame_botoes,
            text="✓ Confirmar",
            command=self.confirmar_selecao
        )
        botao_confirmar.pack(side=tk.LEFT, padx=(0, 10))
        
        botao_recarregar = ttk.Button(
            frame_botoes,
            text="🔄 Recarregar",
            command=self.recarregar_impressoras
        )
        botao_recarregar.pack(side=tk.LEFT)
        
        # Info
        info_label = ttk.Label(
            main_frame,
            text="A automação iniciará após confirmar.",
            font=("Arial", 9),
            foreground="gray"
        )
        info_label.pack(anchor="w", pady=(10, 0))
        
        self.impressora_selecionada = None
        self.padrao_etiqueta = ""
    
    def carregar_impressoras(self):
        """Carrega lista de impressoras na listbox."""
        self.listbox_impressoras.delete(0, tk.END)
        impressoras = obter_impressoras()
        
        if not impressoras:
            messagebox.showwarning(
                "Aviso",
                "Nenhuma impressora encontrada no computador!"
            )
            self.listbox_impressoras.insert(0, "Nenhuma impressora disponível")
            return
        
        for impressora in impressoras:
            self.listbox_impressoras.insert(tk.END, impressora)
        
        # Selecionar a padrão se existir
        try:
            idx = impressoras.index(IMPRESSORA_ZEBRA)
            self.listbox_impressoras.selection_set(idx)
            self.listbox_impressoras.see(idx)
        except ValueError:
            if impressoras:
                self.listbox_impressoras.selection_set(0)
    
    def recarregar_impressoras(self):
        """Recarrega a lista de impressoras."""
        self.carregar_impressoras()
        messagebox.showinfo("Sucesso", "Lista de impressoras recarregada!")
    
    def confirmar_selecao(self):
        """Confirma a seleção e fecha a janela."""
        selecao = self.listbox_impressoras.curselection()
        
        if not selecao:
            messagebox.showerror("Erro", "Selecione uma impressora!")
            return
        
        self.impressora_selecionada = self.listbox_impressoras.get(selecao[0])
        # ler padrão de etiqueta
        self.padrao_etiqueta = self.entry_padrao.get().strip()
        logger.info(f"Impressora selecionada: {self.impressora_selecionada}")
        logger.info(f"Padrão de etiqueta definido: {self.padrao_etiqueta}")
        self.root.destroy()


def aplicar_padrao(padrao: str) -> str:
    """Aplicar tokens simples a um padrão de string.

    O usuário pode definir algo como "Etiqueta_{date}_{time}" e
    os tokens serão substituídos pelos valores atuais.
    """
    now = datetime.now()
    try:
        return padrao.format(
            date=now.strftime("%Y%m%d"),
            time=now.strftime("%H%M%S"),
            datetime=now.strftime("%Y%m%d%H%M%S")
        )
    except Exception:
        # Se houver erro na formatação, devolve o texto cru
        return padrao


def abrir_janela_configuracao() -> tuple[str, str]:
    """
    Abre janela de configuração e retorna impressora e padrão de etiqueta.

    Returns:
        Uma tupla (impressora, padrao_etiqueta)
    """
    root = tk.Tk()
    janela = JanelaConfiguracao(root)
    root.mainloop()
    
    impressora = janela.impressora_selecionada or IMPRESSORA_ZEBRA
    padrao = getattr(janela, "padrao_etiqueta", "")
    return impressora, padrao



def enviar_para_impressora(caminho_txt: str, max_tentativas: int = 3) -> bool:
    """
    Envia arquivo ZPL para impressora Zebra com retry.
    
    Args:
        caminho_txt: Caminho do arquivo ZPL
        max_tentativas: Número máximo de tentativas
        
    Returns:
        True se sucesso, False caso contrário
    """
    if not os.path.exists(caminho_txt):
        logger.error(f"Arquivo não encontrado: {caminho_txt}")
        return False
    
    hPrinter = None
    try:
        for tentativa in range(1, max_tentativas + 1):
            try:
                hPrinter = win32print.OpenPrinter(IMPRESSORA_SELECIONADA)
                
                job = win32print.StartDocPrinter(
                    hPrinter,
                    1,
                    ("Etiqueta MercadoEnvios", None, "RAW")
                )

                win32print.StartPagePrinter(hPrinter)

                with open(caminho_txt, "rb") as f:
                    dados = f.read()

                bytes_escritos = win32print.WritePrinter(hPrinter, dados)

                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)

                logger.info(f"Etiqueta enviada com sucesso. Bytes: {bytes_escritos}")
                return True
                
            except Exception as e:
                logger.warning(f"Tentativa {tentativa}/{max_tentativas} falhou: {e}")
                if tentativa < max_tentativas:
                    time.sleep(2)  # Aguardar antes de retentar
                else:
                    raise

    except Exception as e:
        logger.error(f"Erro ao enviar para impressora: {e}")
        return False

    finally:
        if hPrinter:
            try:
                win32print.ClosePrinter(hPrinter)
            except Exception as e:
                logger.warning(f"Erro ao fechar impressora: {e}")


def extrair_e_validar_zip(caminho_zip: str) -> Optional[str]:
    """
    Extrai ZIP e retorna caminho do arquivo TXT extraído.
    
    Args:
        caminho_zip: Caminho do arquivo ZIP
        
    Returns:
        Caminho do arquivo TXT ou None se inválido
    """
    try:
        with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
            zip_ref.extractall(PASTA_DOWNLOADS)
        
        txt_path = os.path.join(PASTA_DOWNLOADS, NOME_ARQUIVO_TXT)
        
        if os.path.exists(txt_path):
            return txt_path
        else:
            logger.error(f"Arquivo TXT não encontrado no ZIP: {caminho_zip}")
            return None
            
    except zipfile.BadZipFile:
        logger.error(f"ZIP corrompido: {caminho_zip}")
        return None
    except Exception as e:
        logger.error(f"Erro ao extrair ZIP {caminho_zip}: {e}")
        return None


logger.info("Automação iniciada.")


# ================= LOOP PRINCIPAL =================
arquivos_processados = set()

def limpar_cache_se_necessario():
    """Limpa cache se ultrapassar o limite para evitar vazamento de memória."""
    if len(arquivos_processados) > MAX_ARQUIVOS_CACHE:
        arquivos_processados.clear()
        logger.info("Cache de arquivos limpo (limite atingido)")

def processar_fila():
    """Processa fila de downloads uma vez."""
    try:
        if not os.path.exists(PASTA_DOWNLOADS):
            logger.warning(f"Pasta não existe: {PASTA_DOWNLOADS}")
            return

        for arquivo in os.listdir(PASTA_DOWNLOADS):
            if arquivo.startswith(PREFIXO_ZIP) and arquivo.endswith(".zip"):
                caminho_zip = os.path.join(PASTA_DOWNLOADS, arquivo)

                if caminho_zip in arquivos_processados:
                    continue

                logger.info(f"Processando: {arquivo}")

                # Extrair e validar ZIP
                txt_path = extrair_e_validar_zip(caminho_zip)
                
                if not txt_path:
                    logger.error(f"Não foi possível extrair {arquivo}")
                    arquivos_processados.add(caminho_zip)
                    continue

                # se houver um padrão, renomear o arquivo de texto para corresponder
                if PADRAO_ETIQUETA:
                    novo_nome = aplicar_padrao(PADRAO_ETIQUETA)
                    base, ext = os.path.splitext(txt_path)
                    novo_caminho = os.path.join(PASTA_DOWNLOADS, f"{novo_nome}{ext}")
                    try:
                        os.replace(txt_path, novo_caminho)
                        logger.info(f"Arquivo renomeado para padrão: {novo_caminho}")
                        txt_path = novo_caminho
                    except Exception as e:
                        logger.warning(f"Não foi possível renomear para padrão ({PADRAO_ETIQUETA}): {e}")

                # Enviar para impressora
                time.sleep(1)  # Tempo para o arquivo ser liberado
                sucesso = enviar_para_impressora(txt_path)

                # Tempo para a impressora processar
                time.sleep(5)

                # Limpeza de arquivos
                try:
                    if txt_path and os.path.exists(txt_path):
                        os.remove(txt_path)
                    
                    if os.path.exists(caminho_zip):
                        os.remove(caminho_zip)
                    
                    arquivos_processados.add(caminho_zip)
                    logger.info(f"Limpeza concluída: {arquivo}")
                    
                except Exception as e:
                    logger.error(f"Erro na limpeza de arquivos: {e}")
                
                limpar_cache_se_necessario()

    except Exception as e:
        logger.error(f"Erro geral no processamento: {e}")


if __name__ == "__main__":
    try:
        # Abrir janela de configuração
        logger.info("Abrindo janela de configuração...")
        IMPRESSORA_SELECIONADA, PADRAO_ETIQUETA = abrir_janela_configuracao()
        logger.info(f"Usando impressora: {IMPRESSORA_SELECIONADA}")
        if PADRAO_ETIQUETA:
            logger.info(f"Padrão de nome de etiqueta: {PADRAO_ETIQUETA}")
        
        while True:
            processar_fila()
            time.sleep(2)  # Aguardar antes de verificar novamente
            
    except KeyboardInterrupt:
        logger.info("Automação interrompida pelo usuário.")
    except Exception as e:
        logger.critical(f"Erro crítico: {e}", exc_info=True)

