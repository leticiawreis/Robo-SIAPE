import csv
import logging
import os
import shutil
import sys
import tempfile
import time
import zipfile

from datetime import datetime

import requests
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def _obter_pasta_base():
    """Retorna a pasta raiz do projeto (Robo-SIAPE/), onde 'saida/' deve nascer.

    Rodando como script (.py), essa é a pasta onde interface.py/robo_siape.py
    estão de fato.

    Rodando como .exe (PyInstaller --onefile), o executável sempre fica
    dentro de dist/ (ex.: Robo-SIAPE/dist/RoboSIAPE.exe) — então subimos
    mais um nível a partir do .exe para chegar na raiz do projeto
    (Robo-SIAPE/), em vez de deixar a saida/ nascer dentro de dist/.
    """
    if getattr(sys, "frozen", False):
        pasta_dist = os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(pasta_dist)
    return os.path.dirname(os.path.abspath(__file__))


PASTA_SAIDA = os.path.join(_obter_pasta_base(), "saida")
PASTA_EXTRAIDA = os.path.join(PASTA_SAIDA, "_processamento")
MAX_TENTATIVAS_DOWNLOAD = 3

# Índices mantidos por compatibilidade com telas antigas (não são mais usados
# para navegação em navegador — o download agora é direto via requests).
INDICES_ANOS = {
    "2026": 1, "2025": 2, "2024": 3, "2023": 4, "2022": 5,
    "2021": 6, "2020": 7, "2019": 8, "2018": 9, "2017": 10,
    "2016": 11, "2015": 12, "2014": 13, "2013": 14,
}

MESES_OPCOES = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
    "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
}

MESES_2026 = {
    "janeiro", "fevereiro", "março", "abril", "maio", "junho"
}

PALAVRAS_MONETARIAS = [
    "remuneracao", "remuneração", "remunerações", "remuneracoes", "abate-teto",
    "gratificação", "gratificacao", "férias", "ferias", "irrf", "pss", "rpgs",
    "dedução", "deducao", "deduções", "deducoes", "pensão", "pensao", "fundo",
    "taxa", "verbas", "total"

]

FORMATO_MOEDA = "R$ #,##0.00"

CABECALHOS_HTTP = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36"
    )
}


class QueueLogHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        try:
            mensagem = self.format(record)
            self.log_queue.put(("log", mensagem))
        except Exception:
            self.handleError(record)


def meses_disponiveis(ano):
    if ano == "2026":
        return [mes for mes in MESES_OPCOES if mes in MESES_2026]

    return list(MESES_OPCOES.keys())


def criar_pasta_execucao(ano, mes):
    """Cria uma pasta exclusiva para esta execução, sempre nova.

    Mesmo repetindo o mesmo ano/mês, a pasta nunca é reaproveitada nem
    sobrescrita — o nome carrega um carimbo de data/hora e, no
    improvável caso de colisão no mesmo segundo, ganha um sufixo extra.
    """
    marca_tempo = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_base = f"{ano}-{MESES_OPCOES[mes]:02d}_{marca_tempo}"

    caminho = os.path.join(PASTA_SAIDA, nome_base)
    sufixo = 1
    while os.path.exists(caminho):
        sufixo += 1
        caminho = os.path.join(PASTA_SAIDA, f"{nome_base}_{sufixo}")

    os.makedirs(caminho, exist_ok=True)
    return caminho


def configurar_logger(ano, mes, log_queue, pasta_execucao):
    caminho_log = os.path.join(
        pasta_execucao,
        f"execucao_{ano}_{MESES_OPCOES[mes]:02d}.log"
    )

    logger = logging.getLogger(f"robo_siape_{id(log_queue)}")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()

    arquivo_handler = logging.FileHandler(
        caminho_log,
        encoding="utf-8"
    )

    fila_handler = QueueLogHandler(log_queue)

    formato = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%d/%m/%Y %H:%M:%S"
    )

    arquivo_handler.setFormatter(formato)
    fila_handler.setFormatter(formato)

    logger.addHandler(arquivo_handler)
    logger.addHandler(fila_handler)

    return logger, caminho_log


def criar_pasta_trabalho():
    return tempfile.mkdtemp(
        prefix="robo_siape_"
    )


def limpar_pasta_trabalho(pasta):
    if pasta and os.path.exists(pasta):
        shutil.rmtree(pasta, ignore_errors=True)


# =====================================================================
# DOWNLOAD — direto via requests (headless, sem navegador e sem CAPTCHA)
# =====================================================================
def baixar_pacote(ano, mes, tipo, pasta_destino, logger):
    """Baixa o pacote .zip do Portal da Transparência diretamente por HTTP.

    Tenta primeiro o link "oficial" da página de download e, se ele falhar,
    tenta o endereço estático da CGU (destino do redirecionamento), que
    costuma escapar de bloqueios de WAF.
    """
    competencia = f"{ano}{MESES_OPCOES[mes]:02d}"
    nome_do_pacote = f"{competencia}_{tipo}"

    enderecos = [
        f"https://portaldatransparencia.gov.br/download-de-dados/servidores/{nome_do_pacote}",
        f"https://dadosabertos-download.cgu.gov.br/PortalDaTransparencia/saida/servidores/{nome_do_pacote}.zip",
    ]

    os.makedirs(pasta_destino, exist_ok=True)
    caminho_final = os.path.join(pasta_destino, f"{nome_do_pacote}.zip")
    caminho_parcial = caminho_final + ".parcial"

    ultimo_erro = None

    for endereco in enderecos:
        logger.info("Baixando de: %s", endereco)

        try:
            with requests.get(
                endereco,
                headers=CABECALHOS_HTTP,
                stream=True,
                timeout=(15, 120),
            ) as resposta:
                if resposta.status_code == 404:
                    raise FileNotFoundError(
                        f"Pacote {nome_do_pacote} não publicado (404)."
                    )
                resposta.raise_for_status()

                tamanho_informado = int(resposta.headers.get("content-length") or 0)
                bytes_gravados = 0
                proximo_marco = 10

                with open(caminho_parcial, "wb") as arquivo:
                    for bloco in resposta.iter_content(chunk_size=512 * 1024):
                        arquivo.write(bloco)
                        bytes_gravados += len(bloco)

                        if tamanho_informado:
                            percentual = (bytes_gravados / tamanho_informado) * 100
                            if percentual >= proximo_marco:
                                logger.info("Progresso do download: %d%%", int(percentual))
                                proximo_marco += 10

                if not zipfile.is_zipfile(caminho_parcial):
                    os.remove(caminho_parcial)
                    raise ValueError("O conteúdo recebido não é um ZIP válido.")

                shutil.move(caminho_parcial, caminho_final)
                logger.info(
                    "Download concluído: %s (%.1f MB)",
                    os.path.basename(caminho_final),
                    bytes_gravados / 1024 / 1024,
                )
                return caminho_final

        except FileNotFoundError:
            # Pacote realmente não existe para essa competência/tipo — não
            # adianta tentar o outro endereço.
            raise
        except (requests.RequestException, ValueError) as erro:
            logger.warning("Falha ao baixar de %s: %s", endereco, erro)
            ultimo_erro = erro
            if os.path.exists(caminho_parcial):
                os.remove(caminho_parcial)

    raise RuntimeError(
        f"Não foi possível baixar o pacote em nenhum endereço. Último erro: {ultimo_erro}"
    )


def baixar_com_retentativas(ano, mes, tipo, logger):
    pasta_download = criar_pasta_trabalho()

    for tentativa in range(1, MAX_TENTATIVAS_DOWNLOAD + 1):
        try:
            logger.info(
                "Tentativa de download %d/%d.",
                tentativa,
                MAX_TENTATIVAS_DOWNLOAD
            )

            arquivo = baixar_pacote(ano, mes, tipo, pasta_download, logger)

            logger.info(
                "Download validado na tentativa %d/%d.",
                tentativa,
                MAX_TENTATIVAS_DOWNLOAD
            )

            return arquivo, pasta_download

        except FileNotFoundError:
            limpar_pasta_trabalho(pasta_download)
            raise

        except Exception as erro:
            logger.warning(
                "Falha na tentativa %d/%d: %s",
                tentativa,
                MAX_TENTATIVAS_DOWNLOAD,
                erro
            )

            if tentativa == MAX_TENTATIVAS_DOWNLOAD:
                limpar_pasta_trabalho(pasta_download)
                raise RuntimeError(
                    "O download falhou após 3 tentativas."
                ) from erro

            logger.info("Tentando o download novamente.")
            time.sleep(2)

    raise RuntimeError("Não foi possível realizar o download.")


# =====================================================================
# TRATAMENTO DO CSV / GERAÇÃO DO EXCEL
# =====================================================================
def encontrar_arquivo_remuneracao(pasta_extraida):
    candidatos = []

    for raiz, _, arquivos in os.walk(pasta_extraida):
        for nome in arquivos:
            nome_normalizado = nome.lower()

            if "remuneracao" in nome_normalizado or "remuneração" in nome_normalizado:
                if nome_normalizado.endswith(".csv"):
                    candidatos.append(os.path.join(raiz, nome))

    if not candidatos:
        raise FileNotFoundError(
            "Não encontrei o CSV de remuneração dentro do ZIP."
        )

    return candidatos[0]


def preparar_csv(caminho):
    encodings = ("utf-8-sig", "utf-8", "cp1252", "latin1")

    ultimo_erro = None

    for encoding in encodings:
        arquivo = None
        try:
            arquivo = open(
                caminho,
                "r",
                encoding=encoding,
                newline=""
            )

            arquivo.read(4096)
            arquivo.seek(0)

            leitor = csv.reader(
                arquivo,
                delimiter=";"
            )

            return arquivo, leitor

        except UnicodeDecodeError as erro:
            ultimo_erro = erro
            if arquivo is not None:
                arquivo.close()

    raise UnicodeDecodeError(
        "csv",
        b"",
        0,
        1,
        f"Não foi possível identificar a codificação do CSV: {ultimo_erro}"
    )


def coluna_vazia(valor):
    return valor is None or str(valor).strip() == ""



def identificar_colunas_validas(caminho_csv, logger):
    arquivo, leitor = preparar_csv(caminho_csv)

    try:
        cabecalho = next(leitor, None)

        if not cabecalho:
            raise ValueError(
                "O CSV de remuneração está vazio."
            )

        cabecalho = list(cabecalho)
        preenchidas = [False] * len(cabecalho)
        total_linhas = 0

        for linha in leitor:
            if not any(
                not coluna_vazia(valor)
                for valor in linha
            ):
                continue

            total_linhas += 1

            for indice in range(len(cabecalho)):
                if indice < len(linha) and not coluna_vazia(linha[indice]):
                    preenchidas[indice] = True

        colunas_validas = [
            indice
            for indice, preenchida in enumerate(preenchidas)
            if preenchida
        ]

        cabecalho_limpo = [
            cabecalho[indice]
            for indice in colunas_validas
        ]

        logger.info(
            "Tratamento identificado: %d colunas válidas e %d registros não vazios.",
            len(cabecalho_limpo),
            total_linhas
        )

        return cabecalho_limpo, colunas_validas

    finally:
        arquivo.close()


def converter_valor_monetario(valor):
    if valor is None:
        return None

    texto = str(valor).strip()

    if not texto:
        return None

    texto = (
        texto
        .replace("R$", "")
        .replace(" ", "")
    )

    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    else:
        partes = texto.split(".")
        if len(partes) > 2:
            texto = "".join(partes)

    try:
        return float(texto)
    except ValueError:
        return None


def ler_linhas_tratadas(caminho_csv, colunas_validas):
    arquivo, leitor = preparar_csv(caminho_csv)

    try:
        next(leitor, None)

        for linha in leitor:
            if not any(
                not coluna_vazia(valor)
                for valor in linha
            ):
                continue

            linha_saida = []

            for indice in colunas_validas:
                valor = linha[indice] if indice < len(linha) else ""

                if coluna_vazia(valor):
                    linha_saida.append("")
                else:
                    linha_saida.append(valor)

            yield linha_saida

    finally:
        arquivo.close()

def identificar_colunas_monetarias(cabecalho):
    return {
        indice
        for indice, nome in enumerate(cabecalho)
        if nome is not None
        and any(
            palavra in str(nome).lower()
            for palavra in PALAVRAS_MONETARIAS
        )
    }


def calcular_largura_colunas(cabecalho):
    larguras = []

    for nome in cabecalho:
        nome_coluna = (
            str(nome).lower().strip()
            if nome is not None
            else ""
        )

        if "nome" in nome_coluna:
            largura = 60
        elif "cpf" in nome_coluna:
            largura = 18
        elif "cargo" in nome_coluna:
            largura = 35
        elif "órgão" in nome_coluna or "orgao" in nome_coluna:
            largura = 30
        elif any(
            palavra in nome_coluna
            for palavra in PALAVRAS_MONETARIAS
        ):
            largura = 18
        elif "data" in nome_coluna:
            largura = 15
        else:
            largura = min(
                max(len(nome_coluna) + 4, 15),
                30
            )

        larguras.append(largura)

    return larguras



def criar_planilha_formatada(
    cabecalho,
    linhas,
    larguras,
    caminho_saida,
    logger
):
    inicio = time.time()

    workbook = Workbook(write_only=True)
    planilha = workbook.create_sheet(title="Remuneração")

    fonte_cabecalho = Font(
        bold=True,
        color="FFFFFF"
    )

    preenchimento_cabecalho = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    alinhamento_cabecalho = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    borda_cabecalho = Border(
        bottom=Side(
            style="thin",
            color="FFFFFF"
        )
    )

    colunas_monetarias = identificar_colunas_monetarias(
        cabecalho
    )

    for indice, largura in enumerate(larguras, start=1):
        letra = get_column_letter(indice)
        planilha.column_dimensions[letra].width = largura

    planilha.freeze_panes = "A2"
    planilha.row_dimensions[1].height = 35

    linha_cabecalho = []

    for valor in cabecalho:
        celula = WriteOnlyCell(
            planilha,
            value=valor
        )
        celula.font = fonte_cabecalho
        celula.fill = preenchimento_cabecalho
        celula.alignment = alinhamento_cabecalho
        celula.border = borda_cabecalho
        linha_cabecalho.append(celula)

    planilha.append(linha_cabecalho)

    total = 0

    for linha in linhas:
        linha_saida = []

        for indice, valor in enumerate(linha):
            if indice in colunas_monetarias:
                numero = converter_valor_monetario(valor)

                if numero is not None:
                    celula = WriteOnlyCell(
                        planilha,
                        value=numero
                    )
                    celula.number_format = FORMATO_MOEDA
                    linha_saida.append(celula)
                else:
                    linha_saida.append(valor)
            else:
                linha_saida.append(valor)

        planilha.append(linha_saida)
        total += 1

        if total % 50000 == 0:
            logger.info(
                "%d registros processados.",
                total
            )

    ultima_letra = get_column_letter(
        len(larguras)
    )

    planilha.auto_filter.ref = (
        f"A1:{ultima_letra}{total + 1}"
    )

    workbook.save(caminho_saida)

    segundos = time.time() - inicio

    logger.info(
        "Planilha criada com %d registros em %.1f segundos.",
        total,
        segundos
    )

def extrair_zip(caminho_zip, pasta_extraida, logger):
    logger.info("Extraindo o arquivo ZIP.")

    with zipfile.ZipFile(caminho_zip, "r") as zip_ref:
        zip_ref.extractall(pasta_extraida)

    logger.info("ZIP extraído com sucesso.")


def remover_arquivos_desnecessarios(pasta_extraida, logger):
    nomes_desnecessarios = {
        "afastamentos",
        "cadastro",
        "observacoes",
        "observações",
    }

    removidos = 0

    for raiz, _, arquivos in os.walk(pasta_extraida):
        for arquivo in arquivos:
            nome_sem_extensao = os.path.splitext(arquivo)[0].lower()

            if any(
                nome in nome_sem_extensao
                for nome in nomes_desnecessarios
            ):
                caminho = os.path.join(raiz, arquivo)

                try:
                    os.remove(caminho)
                    removidos += 1
                    logger.info(
                        "Arquivo desnecessário removido: %s",
                        arquivo
                    )
                except OSError as erro:
                    logger.warning(
                        "Não foi possível remover %s: %s",
                        arquivo,
                        erro
                    )

    logger.info(
        "Limpeza concluída. %d arquivos removidos.",
        removidos
    )


def validar_periodo(ano, mes, tipo=None):
    if ano not in INDICES_ANOS:
        raise ValueError("Ano inválido.")

    if mes not in MESES_OPCOES:
        raise ValueError("Mês inválido.")

    hoje = datetime.now()

    if (int(ano), MESES_OPCOES[mes]) > (hoje.year, hoje.month):
        raise ValueError(
            "Não é permitido executar para um mês futuro."
        )

    if mes not in meses_disponiveis(ano):
        raise ValueError(
            "O mês selecionado não está disponível para esse ano."
        )



def limpar_arquivos_extraidos(pasta):
    if not pasta or not os.path.isdir(pasta):
        return

    for nome in os.listdir(pasta):
        caminho = os.path.join(pasta, nome)
        try:
            if os.path.isfile(caminho) or os.path.islink(caminho):
                os.remove(caminho)
            elif os.path.isdir(caminho):
                shutil.rmtree(caminho, ignore_errors=True)
        except OSError:
            pass

def executar_pipeline_completo(
    ano,
    mes,
    tipo,
    log_queue,
):
    validar_periodo(ano, mes, tipo)

    os.makedirs(PASTA_SAIDA, exist_ok=True)

    pasta_execucao = criar_pasta_execucao(ano, mes)

    logger, caminho_log = configurar_logger(
        ano,
        mes,
        log_queue,
        pasta_execucao
    )

    pasta_download = None

    try:
        logger.info(
            "Início da execução para %s/%02d — pacote: %s.",
            ano,
            MESES_OPCOES[mes],
            tipo,
        )
        logger.info(
            "Pasta desta execução: %s",
            pasta_execucao,
        )

        caminho_zip, pasta_download = baixar_com_retentativas(
            ano,
            mes,
            tipo,
            logger,
        )

        os.makedirs(PASTA_EXTRAIDA, exist_ok=True)

        pasta_extraida = PASTA_EXTRAIDA

        extrair_zip(
            caminho_zip,
            pasta_extraida,
            logger
        )

        remover_arquivos_desnecessarios(
            pasta_extraida,
            logger
        )

        arquivo_remuneracao = encontrar_arquivo_remuneracao(
            pasta_extraida
        )

        logger.info(
            "CSV de remuneração encontrado: %s",
            arquivo_remuneracao
        )

        caminho_bruto = os.path.join(
            pasta_execucao,
            f"base_bruta_{ano}_{MESES_OPCOES[mes]:02d}.csv"
        )

        shutil.copy2(arquivo_remuneracao, caminho_bruto)
        logger.info(
            "Planilha bruta preservada em: %s",
            caminho_bruto
        )

        cabecalho, colunas_validas = identificar_colunas_validas(
            arquivo_remuneracao,
            logger
        )

        larguras = calcular_largura_colunas(
            cabecalho
        )

        caminho_saida = os.path.join(
            pasta_execucao,
            f"base_tratada_{ano}_{MESES_OPCOES[mes]:02d}.xlsx"
        )

        criar_planilha_formatada(
            cabecalho,
            ler_linhas_tratadas(
                arquivo_remuneracao,
                colunas_validas
            ),
            larguras,
            caminho_saida,
            logger
        )

        logger.info(
            "Execução concluída com sucesso."
        )

        return caminho_saida, caminho_log

    except Exception:
        logger.exception(
            "Erro durante a execução."
        )
        raise

    finally:
        if pasta_download:
            limpar_pasta_trabalho(pasta_download)

        limpar_arquivos_extraidos(PASTA_EXTRAIDA)

        for handler in list(logger.handlers):
            handler.flush()
            handler.close()
            logger.removeHandler(handler)