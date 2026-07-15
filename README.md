# Zebra Agent — Automação de Impressão de Etiquetas

Automação em Python que eliminou um processo manual de 10–15 minutos por pedido, reduzindo para menos de 1 minuto com um único clique.

---

## O problema

O fluxo de impressão de etiquetas no setor de estoque exigia várias etapas manuais:

1. Baixar o arquivo ZIP do Mercado Livre
2. Extrair o arquivo manualmente
3. Localizar o TXT com a etiqueta
4. Copiar e colar em programa intermediário
5. Acionar a impressão

Cada pedido levava entre **10 e 15 minutos**. O processo era repetitivo, suscetível a erro humano e ocupava tempo produtivo do operador a cada ciclo.

---

## A solução

Script Python que monitora a pasta de Downloads continuamente. Quando detecta um novo ZIP do Mercado Livre, extrai o arquivo, envia o ZPL diretamente para a impressora Zebra via spooler do Windows — **sem nenhuma intervenção manual**.

```
Downloads (ZIP) → extração automática → TXT (ZPL) → spooler Windows → Impressora Zebra
```

---

## Resultado

| Métrica | Antes | Depois |
|---|---|---|
| Tempo por pedido | 10–15 min | < 1 min |
| Etapas manuais | 5 | 0 |
| Intervenção humana | A cada pedido | Nenhuma |

Redução de **~93% no tempo de processamento** por pedido.

---

## Como funciona

1. Script roda em segundo plano monitorando a pasta `Downloads` em loop contínuo
2. Detecta arquivos ZIP com o prefixo `"Etiqueta MercadoEnvios"`
3. Extrai o ZIP e localiza o arquivo `"Etiqueta de envio.txt"` (ZPL puro)
4. Abre conexão com a impressora Zebra via `win32print` e envia o ZPL em modo RAW
5. Aguarda o spooler processar, remove os arquivos temporários e volta a monitorar
6. Registra cada operação em log (`C:\Automacao_Zebra\logs\log.txt`)

---

## Detalhes técnicos

**Comunicação com a impressora**
- Driver Zebra instalado no Windows
- Envio via `win32print.WritePrinter()` em modo `RAW`
- ZPL (Zebra Programming Language) enviado diretamente — sem conversão PDF ou imagem

**Tratamento de erros**
- Controle de reimpressão: arquivos já processados são rastreados em memória (`set`)
- Arquivos removidos apenas após impressão confirmada — preservados em caso de falha
- Log completo com timestamp, bytes enviados e erros

---

## Tecnologias

| Biblioteca | Uso |
|---|---|
| `os` | acesso a diretórios, remoção de arquivos |
| `time` | delays e monitoramento em loop |
| `zipfile` | extração dos ZIPs do Mercado Livre |
| `win32print` | comunicação direta com impressora Windows |
| `datetime` | timestamps nos logs |

- **Python 3** — linguagem principal
- **Windows** — execução local via spooler
- **Impressora Zebra** — hardware de destino (ZPL)

---

## Inicialização automática

O script roda como processo em segundo plano, iniciado por um script VBS (`iniciar_automacao.vbs`) configurado para rodar silenciosamente (sem janela de terminal visível):

```vbs
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "pythonw C:\Automacao_Zebra\impressao_automatica_02_04.py", 0, False
```

## Organização do repositório

- `impressao_automatica_02_04.py` — versão em produção (nome mantido idêntico ao usado na máquina real, para rastreabilidade)
- `iniciar_automacao.vbs` — script de inicialização silenciosa
- `archive/` — versões anteriores, mantidas como referência histórica do processo de evolução do script (não estão em uso)

Dados operacionais (logs, CSV de etiquetas processadas, controle de arquivos já processados) não são versionados aqui — crescem continuamente e não são código-fonte. Veja `.gitignore`.

## Contexto

Desenvolvido internamente na [Tradipar](https://tradipar.com.br) — distribuidora B2B/B2C de ferramentas, EPIs e equipamentos industriais com presença em mais de 22 estados brasileiros.

Criado como iniciativa própria para resolver um gargalo operacional real no setor de estoque.

---

## Autores

**Núbya** — analista de CRM & automação de processos
[LinkedIn](https://www.linkedin.com/in/nubya-santos-caldas/) · [GitHub](https://github.com/nubyacaldas)

**Eduardo Vieira** — infraestrutura e redes
[LinkedIn](https://www.linkedin.com/in/eduardo-vieira-78b853285/) · [GitHub](https://github.com/Edu936)
