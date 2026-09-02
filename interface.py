import hashlib
import json
import logging
import os
import queue
import re
import sys
import threading
from collections import OrderedDict
from datetime import date, datetime

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QRectF, QUrl
from PyQt6.QtGui import QDesktopServices, QFont, QColor, QPainter, QPen, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from robo_siape import (
    INDICES_ANOS,
    meses_disponiveis,
    executar_pipeline_completo,
    PASTA_SAIDA,
    ExecucaoCancelada,
)

TIPO_PACOTE = "Servidores_SIAPE"

# Pasta onde ficam as imagens usadas no diálogo de ajuda passo a passo
# (prints da própria interface, já com setas/realces indicando onde clicar).
#
# Rodando como script (.py), fica ao lado deste arquivo. Rodando como .exe
# (PyInstaller --onefile, com a pasta embutida via --add-data no build.py),
# os arquivos são extraídos em tempo de execução para uma pasta temporária
# apontada por sys._MEIPASS — é lá que precisamos procurar, não ao lado do
# .exe.
def _obter_pasta_assets_ajuda():
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "assets", "ajuda")


PASTA_ASSETS_AJUDA = _obter_pasta_assets_ajuda()


ARQUIVO_USUARIOS = "usuarios.json"
ARQUIVO_HISTORICO = "historico_execucoes.json"
ARQUIVO_CONFIG = "config.json"


def preparar_pastas():
    """Garante que a pasta de saída (PASTA_SAIDA) exista no disco.

    Cria a pasta (e eventuais pastas intermediárias) caso ainda não
    existam. Não faz nada se a pasta já existir (exist_ok=True).
    """
    os.makedirs(PASTA_SAIDA, exist_ok=True)


BG = "#F5F1E7"
SURFACE = "#FFFDF8"
SURFACE_2 = "#F8F3E8"
SURFACE_3 = "#F1EBDD"
BORDER = "#E3DCCF"
YELLOW = "#D5A63A"
YELLOW_LIGHT = "#E7BD58"
TEXT = "#25221D"
MUTED = "#756E63"
SUCCESS = "#4F9B73"
ERROR = "#C85B5B"
WARNING = "#C58D2D"


def carregar_usuarios():
    """Carrega o dicionário de usuários cadastrados a partir do arquivo JSON.

    Se o arquivo ARQUIVO_USUARIOS ainda não existir em disco, retorna um
    dicionário vazio (nenhum usuário cadastrado). Caso exista, lê e
    decodifica o conteúdo JSON e o retorna como dict.

    Retorna:
        dict: mapeia nome de usuário -> dados do usuário (ex: hash da senha).
    """
    if not os.path.exists(ARQUIVO_USUARIOS):
        return {}
    with open(ARQUIVO_USUARIOS, "r", encoding="utf-8") as arquivo:
        return json.load(arquivo)


def salvar_usuarios(usuarios):
    """Persiste o dicionário de usuários no arquivo JSON local.

    Sobrescreve o conteúdo de ARQUIVO_USUARIOS com o dicionário
    fornecido, formatado com indentação e mantendo acentuação (ensure_ascii=False).

    Parâmetros:
        usuarios (dict): dicionário completo de usuários a ser salvo.
    """
    with open(ARQUIVO_USUARIOS, "w", encoding="utf-8") as arquivo:
        json.dump(usuarios, arquivo, indent=2, ensure_ascii=False)


def hash_senha(senha):
    """Gera o hash SHA-256 de uma senha em texto puro.

    Usado tanto no cadastro (para salvar a senha de forma não reversível)
    quanto no login (para comparar a senha digitada com a armazenada).

    Parâmetros:
        senha (str): senha em texto puro.

    Retorna:
        str: representação hexadecimal do hash SHA-256 da senha.
    """
    return hashlib.sha256(senha.encode("utf-8")).hexdigest()


def usuario_valido(usuario):
    """Valida se um nome de usuário segue o formato permitido.

    Regras: de 3 a 30 caracteres, contendo apenas letras (maiúsculas ou
    minúsculas), números, underscore (_), ponto (.) ou hífen (-).

    Parâmetros:
        usuario (str): nome de usuário a validar.

    Retorna:
        bool: True se o nome de usuário for válido, False caso contrário.
    """
    return re.fullmatch(r"[A-Za-z0-9_.-]{3,30}", usuario) is not None


# =====================================================================
# HISTÓRICO DE EXECUÇÕES — alimenta o gráfico e os indicadores do painel
# =====================================================================
def carregar_historico():
    """Carrega o histórico de execuções do robô a partir do arquivo JSON.

    Se o arquivo não existir, ou se seu conteúdo estiver corrompido/
    ilegível (JSONDecodeError ou erro de leitura), retorna uma lista
    vazia em vez de lançar exceção — assim o painel nunca quebra por
    causa de um histórico corrompido.

    Retorna:
        list: lista de dicionários, cada um representando uma execução
        registrada (ano, mês, status, arquivo gerado, data).
    """
    if not os.path.exists(ARQUIVO_HISTORICO):
        return []
    try:
        with open(ARQUIVO_HISTORICO, "r", encoding="utf-8") as arquivo:
            return json.load(arquivo)
    except (json.JSONDecodeError, OSError):
        return []


def salvar_historico(historico):
    """Persiste a lista de histórico de execuções no arquivo JSON.

    Mantém apenas os 200 registros mais recentes (historico[-200:]) para
    evitar que o arquivo cresça indefinidamente.

    Parâmetros:
        historico (list): lista completa de execuções a salvar.
    """
    with open(ARQUIVO_HISTORICO, "w", encoding="utf-8") as arquivo:
        json.dump(historico[-200:], arquivo, indent=2, ensure_ascii=False)


def registrar_execucao(ano, mes, status, arquivo_gerado=None):
    """Adiciona um novo registro de execução ao histórico e salva em disco.

    Carrega o histórico atual, acrescenta um novo item com os dados da
    execução (ano, mês, status, caminho do arquivo gerado e timestamp
    atual) e grava o histórico atualizado de volta no arquivo.

    Parâmetros:
        ano (str): ano da execução (ex: "2026").
        mes (str): mês da execução, por extenso (ex: "janeiro").
        status (str): resultado da execução ("sucesso" ou "erro").
        arquivo_gerado (str, opcional): caminho do arquivo Excel gerado,
            quando a execução teve sucesso.

    Retorna:
        list: o histórico já atualizado (com o novo item incluído).
    """
    historico = carregar_historico()
    historico.append(
        {
            "ano": ano,
            "mes": mes,
            "status": status,
            "arquivo": arquivo_gerado,
            "data": datetime.now().isoformat(timespec="seconds"),
        }
    )
    salvar_historico(historico)
    return historico


def estatisticas_execucoes():
    """Calcula estatísticas agregadas sobre o histórico de execuções.

    Percorre o histórico salvo e conta o total de execuções, quantas
    tiveram status "sucesso" e a taxa de sucesso em percentual
    (arredondada para o inteiro mais próximo).

    Retorna:
        tuple: (total, sucesso, taxa) onde
            total (int): quantidade total de execuções registradas;
            sucesso (int): quantidade de execuções com status "sucesso";
            taxa (int): percentual de sucesso (0 a 100), ou 0 se não
                houver nenhuma execução registrada.
    """
    historico = carregar_historico()
    total = len(historico)
    sucesso = sum(1 for item in historico if item.get("status") == "sucesso")
    taxa = round((sucesso / total) * 100) if total else 0
    return total, sucesso, taxa


def contagem_execucoes_por_dia(qtd_dias=6):
    """Agrupa e conta as execuções pela data real em que foram feitas.

    Diferente de agrupar pela competência escolhida (o ano/mês do
    arquivo baixado), esta função agrupa pelo timestamp real de cada
    execução — o campo "data" gravado no histórico no momento em que o
    robô rodou —, no formato "dd/mm/aa". Ou seja, o gráfico reflete o
    dia, mês e ano em que cada execução de fato aconteceu, e não qual
    competência foi mais baixada.

    Retorna apenas os últimos `qtd_dias` dias com execução registrada;
    se houver menos dias do que `qtd_dias`, completa a lista com zeros e
    rótulos vazios no início, para manter o tamanho fixo (usado pelo
    BarChartWidget).

    Parâmetros:
        qtd_dias (int): quantidade de dias (barras) a retornar.
            Padrão: 6.

    Retorna:
        tuple: (chaves, valores) onde
            chaves (list[str]): rótulos dos dias (ex: ["01/09/26", ...]);
            valores (list[int]): quantidade de execuções em cada dia,
                na mesma ordem das chaves.
    """
    historico = carregar_historico()
    contagem = OrderedDict()
    for item in historico:
        data_str = str(item.get("data", "")).strip()
        try:
            data = datetime.fromisoformat(data_str)
        except ValueError:
            continue
        chave = data.strftime("%d/%m/%y")
        contagem[chave] = contagem.get(chave, 0) + 1
    chaves = list(contagem.keys())[-qtd_dias:]
    valores = [contagem[c] for c in chaves]
    while len(valores) < qtd_dias:
        valores.insert(0, 0)
        chaves.insert(0, "")
    return chaves, valores


def ultima_execucao():
    """Retorna o registro da execução mais recente do histórico.

    Retorna:
        dict | None: o último item do histórico (mais recente), ou None
        se ainda não houver nenhuma execução registrada.
    """
    historico = carregar_historico()
    if not historico:
        return None
    return historico[-1]


# =====================================================================
# CONFIGURAÇÕES DO USUÁRIO — preferências reais aplicadas no painel
# =====================================================================
def carregar_config():
    """Carrega as preferências do usuário (config.json), com valores padrão.

    Começa com um dicionário padrão (notificar ao concluir = True, abrir
    pasta automaticamente = False) e, se o arquivo ARQUIVO_CONFIG existir
    e for válido, sobrescreve esses valores com o conteúdo salvo. Se o
    arquivo estiver corrompido ou ilegível, os valores padrão são
    mantidos silenciosamente.

    Retorna:
        dict: configurações efetivas do usuário.
    """
    padrao = {"notificar_conclusao": True, "abrir_pasta_automatico": False}
    if os.path.exists(ARQUIVO_CONFIG):
        try:
            with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as arquivo:
                padrao.update(json.load(arquivo))
        except (json.JSONDecodeError, OSError):
            pass
    return padrao


def salvar_config(config):
    """Persiste o dicionário de configurações do usuário em disco.

    Parâmetros:
        config (dict): configurações completas a serem salvas em
            ARQUIVO_CONFIG.
    """
    with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as arquivo:
        json.dump(config, arquivo, indent=2, ensure_ascii=False)


# =====================================================================
# ARQUIVOS DE SAÍDA — lista real do que o robô já gerou em disco
# =====================================================================
def formatar_tamanho(num_bytes):
    """Converte um tamanho em bytes para uma string legível (B, KB, MB, GB).

    Parâmetros:
        num_bytes (int | float): tamanho em bytes.

    Retorna:
        str: tamanho formatado com a unidade mais adequada, por exemplo
        "512 B", "3.4 MB" ou "1.2 GB".
    """
    tamanho = float(num_bytes)
    for unidade in ("B", "KB", "MB", "GB"):
        if tamanho < 1024 or unidade == "GB":
            return f"{tamanho:.0f} {unidade}" if unidade == "B" else f"{tamanho:.1f} {unidade}"
        tamanho /= 1024
    return f"{tamanho:.1f} GB"


def listar_arquivos_saida(limite=6):
    """Lista as planilhas (.xlsx) mais recentes geradas pelo robô.

    Percorre recursivamente a PASTA_SAIDA (ignorando a subpasta
    "_processamento", usada apenas como área de trabalho temporária) e
    coleta informações de cada arquivo .xlsx encontrado: nome, caminho
    completo, tamanho formatado e data de modificação. Os resultados são
    ordenados do mais recente para o mais antigo.

    Parâmetros:
        limite (int): quantidade máxima de arquivos a retornar. Padrão: 6.

    Retorna:
        list[dict]: lista de arquivos, cada um com as chaves "nome",
        "caminho", "tamanho", "data" e "mtime". Retorna lista vazia se a
        PASTA_SAIDA ainda não existir.
    """
    if not os.path.isdir(PASTA_SAIDA):
        return []
    itens = []
    for raiz, pastas, arquivos in os.walk(PASTA_SAIDA):
        if "_processamento" in pastas:
            pastas.remove("_processamento")
        for nome in arquivos:
            if not nome.lower().endswith(".xlsx"):
                continue
            caminho = os.path.join(raiz, nome)
            info = os.stat(caminho)
            itens.append(
                {
                    "nome": nome,
                    "caminho": caminho,
                    "tamanho": formatar_tamanho(info.st_size),
                    "data": datetime.fromtimestamp(info.st_mtime).strftime("%d/%m/%Y"),
                    "mtime": info.st_mtime,
                }
            )
    itens.sort(key=lambda item: item["mtime"], reverse=True)
    return itens[:limite]


# =====================================================================
# RESUMO DOS LOGS — consolida todas as execuções num único .txt
# =====================================================================
def gerar_resumo_logs():
    """Gera um resumo em texto (.txt) de todas as execuções já realizadas.

    Percorre recursivamente a PASTA_SAIDA (ignorando "_processamento") em
    busca de todos os arquivos .log já gerados, lê o conteúdo de cada um
    e conta quantas linhas de INFO, WARNING e ERROR ele possui. Em
    seguida, monta um resumo único (mais recente por último) com o total
    de execuções, a taxa de sucesso geral e, para cada execução, sua
    pasta, data/hora, status e contagem de linhas por nível de log.

    O resumo é sempre regravado do zero em "resumo_execucoes.txt", dentro
    da PASTA_SAIDA, refletindo o estado atual de todas as execuções.

    Retorna:
        str | None: caminho do arquivo de resumo gerado, ou None se
        ainda não existir nenhum log registrado.
    """
    if not os.path.isdir(PASTA_SAIDA):
        return None

    logs = []
    for raiz, pastas, arquivos in os.walk(PASTA_SAIDA):
        if "_processamento" in pastas:
            pastas.remove("_processamento")
        for nome in arquivos:
            if nome.lower().endswith(".log"):
                caminho = os.path.join(raiz, nome)
                logs.append((os.path.getmtime(caminho), caminho))

    if not logs:
        return None

    logs.sort(key=lambda item: item[0])

    total, sucesso, taxa = estatisticas_execucoes()

    linhas = [
        "RESUMO DAS EXECUÇÕES — ROBÔ SIAPE",
        f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        "=" * 60,
        f"Total de execuções registradas: {total}",
        f"Execuções com sucesso: {sucesso}",
        f"Taxa de sucesso: {taxa}%",
        f"Total de logs encontrados: {len(logs)}",
        "=" * 60,
        "",
    ]

    for mtime, caminho in logs:
        pasta_execucao = os.path.basename(os.path.dirname(caminho))
        try:
            with open(caminho, "r", encoding="utf-8", errors="ignore") as arquivo:
                conteudo = arquivo.readlines()
        except OSError:
            conteudo = []

        infos = sum(1 for linha in conteudo if re.search(r"\bINFO\b", linha))
        avisos = sum(1 for linha in conteudo if re.search(r"\bWARNING\b", linha))
        erros = sum(1 for linha in conteudo if re.search(r"\bERROR\b", linha))
        data_log = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M:%S")
        # Uma execução interrompida pode não ter nenhum ERROR no log.
        # Nesse caso, o status correto não é SUCESSO.
        interrompida = any(
            "Execução interrompida pelo usuário" in linha
            or "Execução cancelada pelo usuário" in linha
            or "Download interrompido pelo usuário" in linha
            for linha in conteudo
        )

        if interrompida:
            status = "INTERROMPIDO"
        elif erros:
            status = "ERRO"
        else:
            status = "SUCESSO"

        linhas.append(f"Execução: {pasta_execucao}")
        linhas.append(f"  Data/hora do log: {data_log}")
        linhas.append(f"  Status: {status}")
        linhas.append(
            f"  Linhas de log: {len(conteudo)}  "
            f"(INFO: {infos}  WARNING: {avisos}  ERROR: {erros})"
        )
        linhas.append(f"  Arquivo: {os.path.basename(caminho)}")
        linhas.append("-" * 60)

    caminho_resumo = os.path.join(PASTA_SAIDA, "resumo_execucoes.txt")
    with open(caminho_resumo, "w", encoding="utf-8") as arquivo:
        arquivo.write("\n".join(linhas))

    return caminho_resumo


def estilizar_combo(combo):
    """Aplica o estilo visual padrão (CSS/QSS) a um QComboBox da interface.

    Define altura mínima e uma folha de estilos (cores de fundo, borda,
    destaque ao passar o mouse/focar, e aparência da lista suspensa) de
    acordo com a paleta de cores do app (SURFACE_2, BORDER, YELLOW etc.).

    Parâmetros:
        combo (QComboBox): o combo box a ser estilizado (alterado in-place).
    """
    combo.setMinimumHeight(42)
    combo.setStyleSheet(
        f"""
        QComboBox {{
            background: {SURFACE_2};
            color: {TEXT};
            border: 1px solid {BORDER};
            border-radius: 10px;
            padding: 0 14px;
            font-size: 14px;
        }}
        QComboBox:hover, QComboBox:focus {{
            border: 1px solid {YELLOW};
        }}
        QComboBox::drop-down {{
            border: none;
            width: 34px;
        }}
        QComboBox QAbstractItemView {{
            background: {SURFACE_2};
            color: {TEXT};
            border: 1px solid {BORDER};
            selection-background-color: {YELLOW};
            selection-color: #111111;
            padding: 6px;
        }}
        """
    )


def criar_titulo(texto, subtitulo=None):
    """Cria um bloco de título (e subtítulo opcional) reutilizado nas telas.

    Monta um QWidget contendo um QLabel de título (objectName "pageTitle")
    e, se fornecido, um QLabel de subtítulo (objectName "pageSubtitle")
    logo abaixo, com quebra de linha automática.

    Parâmetros:
        texto (str): texto principal do título.
        subtitulo (str, opcional): texto secundário exibido abaixo do
            título. Se None, nenhum subtítulo é adicionado.

    Retorna:
        QWidget: bloco pronto para ser inserido em um layout.
    """
    bloco = QWidget()
    layout = QVBoxLayout(bloco)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(5)

    titulo = QLabel(texto)
    titulo.setObjectName("pageTitle")
    layout.addWidget(titulo)

    if subtitulo:
        sub = QLabel(subtitulo)
        sub.setObjectName("pageSubtitle")
        sub.setWordWrap(True)
        layout.addWidget(sub)

    return bloco


class LogoMark(QWidget):
    """Widget customizado que desenha o "logotipo" do app (um quadrado
    amarelo arredondado com as iniciais "RS", de Robô SIAPE).

    O desenho é feito manualmente via QPainter no evento paintEvent,
    em vez de usar uma imagem/arquivo externo.
    """

    def __init__(self, parent=None):
        """Inicializa o widget com um tamanho fixo de 54x54 pixels.

        Parâmetros:
            parent (QWidget, opcional): widget pai, se houver.
        """
        super().__init__(parent)
        self.setFixedSize(44, 44)

    def paintEvent(self, event):
        """Desenha o logotipo: um retângulo arredondado amarelo com as
        letras "RS" centralizadas em preto.

        Chamado automaticamente pelo Qt sempre que o widget precisa ser
        redesenhado (não deve ser chamado manualmente).

        Parâmetros:
            event (QPaintEvent): evento de pintura fornecido pelo Qt.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setBrush(QColor(YELLOW))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(2, 2, 40, 40, 12, 12)
        painter.setPen(QPen(QColor("#0D0D0D"), 2))
        painter.setFont(QFont("Arial", 13, QFont.Weight.Bold))
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "RS")


class Campo(QLineEdit):
    """Campo de texto (QLineEdit) estilizado, usado em formulários de
    login, cadastro e demais entradas de texto do app.
    """

    def __init__(self, placeholder="", senha=False):
        """Cria o campo de texto já com placeholder, altura e estilo padrão.

        Parâmetros:
            placeholder (str): texto de dica exibido quando o campo está
                vazio.
            senha (bool): se True, o campo oculta o texto digitado (modo
                senha); se False, exibe o texto normalmente.
        """
        super().__init__()
        self.setPlaceholderText(placeholder)
        self.setMinimumHeight(42)
        self.setEchoMode(
            QLineEdit.EchoMode.Password if senha else QLineEdit.EchoMode.Normal
        )
        self.setStyleSheet(
            f"""
            QLineEdit {{
                background: {SURFACE_2};
                color: {TEXT};
                border: 1px solid {BORDER};
                border-radius: 10px;
                padding: 0 15px;
                font-size: 14px;
                selection-background-color: {YELLOW};
                selection-color: #111111;
            }}
            QLineEdit:hover, QLineEdit:focus {{
                border: 1px solid {YELLOW};
                background: {SURFACE_3};
            }}
            """
        )


class Botao(QPushButton):
    """Botão (QPushButton) padronizado do app, com variações de estilo
    controladas por propriedades customizadas ("principal" e "compacto"),
    usadas pela folha de estilos QSS global para diferenciar aparência
    (botão de destaque vs. botão secundário/compacto).
    """

    def __init__(self, texto, principal=False, compacto=False, perigo=False, grande=False):
        """Cria o botão com cursor de mão, altura e propriedades de estilo.

        Parâmetros:
            texto (str): texto exibido no botão.
            principal (bool): se True, aplica o estilo de botão de
                destaque (cor amarela, usado em ações principais).
            compacto (bool): se True, usa altura reduzida (usado em
                botões secundários/ações rápidas).
            perigo (bool): se True, aplica o estilo de alerta (vermelho),
                usado em ações destrutivas/de interrupção (ex: "Parar").
            grande (bool): se True (combinado com compacto=True), aumenta
                altura e fonte do botão — usado quando poucos botões
                dividem uma barra e podem ocupar mais espaço (ex: barra
                de atalhos do painel).
        """
        super().__init__(texto)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(56 if grande else (38 if compacto else 44))
        self.setProperty("principal", principal)
        self.setProperty("compacto", compacto)
        self.setProperty("perigo", perigo)
        self.setProperty("grande", grande)


class Card(QFrame):
    """Contêiner visual (QFrame) usado como "cartão" nas telas do app,
    com objectName "card" para receber o estilo de fundo/borda/raio
    definido na folha de estilos global.
    """

    def __init__(self, parent=None):
        """Inicializa o frame e define seu objectName para estilização.

        Parâmetros:
            parent (QWidget, opcional): widget pai, se houver.
        """
        super().__init__(parent)
        self.setObjectName("card")


class GaugeWidget(QWidget):
    """Indicador circular (gauge/velocímetro) que exibe um valor numérico
    real (contagem ou percentual) como um arco colorido sobre um arco de
    fundo, com um rótulo abaixo.

    Usado no painel para mostrar, por exemplo, o total de execuções e a
    taxa de sucesso.
    """

    def __init__(self, valor=0, maximo=100, sufixo="", rotulo="", parent=None):
        """Inicializa o gauge com os valores iniciais a exibir.

        Parâmetros:
            valor (int|float): valor atual a ser representado no arco.
            maximo (int|float): valor máximo possível (define 100% do
                arco); nunca é permitido ser menor que 1.
            sufixo (str): texto exibido após o valor (ex: "%").
            rotulo (str): texto descritivo exibido abaixo do valor
                (ex: "Execuções").
            parent (QWidget, opcional): widget pai, se houver.
        """
        super().__init__(parent)
        self.valor = valor
        self.maximo = max(maximo, 1)
        self.sufixo = sufixo
        self.rotulo = rotulo
        self.setMinimumSize(100, 112)

    def definir_valor(self, valor, maximo=None, rotulo=None):
        """Atualiza o valor (e opcionalmente o máximo e o rótulo) exibidos
        pelo gauge, e solicita o redesenho do widget.

        Parâmetros:
            valor (int|float): novo valor a exibir.
            maximo (int|float, opcional): novo valor máximo, se for
                necessário alterá-lo.
            rotulo (str, opcional): novo texto de rótulo, se for
                necessário alterá-lo.
        """
        self.valor = valor
        if maximo is not None:
            self.maximo = max(maximo, 1)
        if rotulo is not None:
            self.rotulo = rotulo
        self.update()

    def paintEvent(self, event):
        """Desenha o arco de fundo, o arco de valor (proporcional a
        valor/máximo) e os textos de valor e rótulo, centralizados.

        Chamado automaticamente pelo Qt quando o widget precisa ser
        redesenhado.

        Parâmetros:
            event (QPaintEvent): evento de pintura fornecido pelo Qt.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        lado = min(self.width(), self.height() - 26) - 16
        x = (self.width() - lado) / 2
        y = 8
        area = QRectF(x, y, lado, lado)

        caneta_fundo = QPen(QColor(BORDER))
        caneta_fundo.setWidth(8)
        caneta_fundo.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(caneta_fundo)
        painter.drawArc(area, 225 * 16, -270 * 16)

        proporcao = min(self.valor / self.maximo, 1.0) if self.maximo else 0
        caneta_valor = QPen(QColor(YELLOW))
        caneta_valor.setWidth(8)
        caneta_valor.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(caneta_valor)
        painter.drawArc(area, 225 * 16, int(-270 * proporcao * 16))

        painter.setPen(QColor(TEXT))
        painter.setFont(QFont("Segoe UI", 16, QFont.Weight.Black))
        texto_valor = f"{self.valor}{self.sufixo}"
        painter.drawText(area, Qt.AlignmentFlag.AlignCenter, texto_valor)

        rotulo_rect = QRectF(0, y + lado - 6, self.width(), 26)
        painter.setPen(QColor(MUTED))
        painter.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
        painter.drawText(rotulo_rect, Qt.AlignmentFlag.AlignHCenter, self.rotulo.upper())


class BarChartWidget(QWidget):
    """Mini gráfico de barras desenhado manualmente (via QPainter),
    exibindo a quantidade de execuções por dia real de execução (dados
    vindos de contagem_execucoes_por_dia).
    """

    def __init__(self, valores=None, rotulos=None, parent=None):
        """Inicializa o gráfico com os valores e rótulos iniciais.

        Parâmetros:
            valores (list[int], opcional): altura de cada barra
                (quantidade de execuções por período). Padrão: lista vazia.
            rotulos (list[str], opcional): rótulo exibido abaixo de cada
                barra (ex: "Jan/26"). Padrão: lista vazia.
            parent (QWidget, opcional): widget pai, se houver.
        """
        super().__init__(parent)
        self.valores = valores or []
        self.rotulos = rotulos or []
        self.setMinimumHeight(100)

    def definir_dados(self, valores, rotulos):
        """Substitui os dados do gráfico e solicita seu redesenho.

        Parâmetros:
            valores (list[int]): novos valores (altura das barras).
            rotulos (list[str]): novos rótulos, um por barra.
        """
        self.valores = valores
        self.rotulos = rotulos
        self.update()

    def paintEvent(self, event):
        """Desenha cada barra com altura proporcional ao seu valor
        (relativo ao maior valor da lista), destacando a última barra
        (mês mais recente) com uma cor diferente, e desenha o rótulo de
        cada barra logo abaixo dela.

        Não desenha nada se a lista de valores estiver vazia.

        Parâmetros:
            event (QPaintEvent): evento de pintura fornecido pelo Qt.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        if not self.valores:
            return

        maximo = max(self.valores) or 1
        n = len(self.valores)
        margem = 6
        base_y = self.height() - 20
        altura_max = self.height() - 34
        largura_total = self.width() - margem * 2
        espaco = 8
        largura_barra = max((largura_total - espaco * (n - 1)) / n, 6)

        painter.setPen(Qt.PenStyle.NoPen)
        for indice, valor in enumerate(self.valores):
            altura = 4 if valor == 0 else max(6, (valor / maximo) * altura_max)
            pos_x = margem + indice * (largura_barra + espaco)
            pos_y = base_y - altura
            cor = QColor(YELLOW if indice == n - 1 else YELLOW_LIGHT if valor else BORDER)
            painter.setBrush(cor)
            painter.drawRoundedRect(QRectF(pos_x, pos_y, largura_barra, altura), 3, 3)

            if indice < len(self.rotulos) and self.rotulos[indice]:
                painter.setPen(QColor(MUTED))
                painter.setFont(QFont("Segoe UI", 7, QFont.Weight.Bold))
                painter.drawText(
                    QRectF(pos_x - 4, self.height() - 15, largura_barra + 8, 14),
                    Qt.AlignmentFlag.AlignHCenter,
                    self.rotulos[indice],
                )
                painter.setPen(Qt.PenStyle.NoPen)


class ToggleSwitch(QPushButton):
    """Interruptor liga/desliga (estilo "switch" mobile) desenhado
    manualmente sobre um QPushButton "checkable", usado para as
    preferências reais do usuário (ex: notificar ao concluir).
    """

    def __init__(self, marcado=False, parent=None):
        """Inicializa o switch com o estado inicial (ligado/desligado).

        Parâmetros:
            marcado (bool): estado inicial do switch (True = ligado).
            parent (QWidget, opcional): widget pai, se houver.
        """
        super().__init__(parent)
        self.setCheckable(True)
        self.setChecked(marcado)
        self.setFixedSize(40, 22)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setObjectName("toggleSwitch")

    def paintEvent(self, event):
        """Desenha a "trilha" do switch (colorida se ligado, cinza se
        desligado) e o "botão" circular deslizante, posicionado à
        direita quando ligado e à esquerda quando desligado.

        Parâmetros:
            event (QPaintEvent): evento de pintura fornecido pelo Qt.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        raio = self.height() / 2

        cor_fundo = QColor(YELLOW) if self.isChecked() else QColor(BORDER)
        painter.setBrush(cor_fundo)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(0, 0, self.width(), self.height(), raio, raio)

        diametro = self.height() - 6
        pos_x = self.width() - diametro - 3 if self.isChecked() else 3
        painter.setBrush(QColor("#FFFFFF"))
        painter.drawEllipse(int(pos_x), 3, diametro, diametro)


class Worker(QThread):
    """Thread secundária (QThread) que executa o pipeline completo do
    robô (download, extração, tratamento e geração da planilha) em
    segundo plano, para não travar a interface gráfica.

    Comunica-se com a interface principal através de sinais Qt:
    log_signal (mensagens de log em tempo real), done_signal (execução
    concluída com sucesso) e error_signal (execução falhou).
    """

    log_signal = pyqtSignal(str)
    done_signal = pyqtSignal(str, str)
    error_signal = pyqtSignal(str)
    cancelado_signal = pyqtSignal()

    def __init__(self, ano, mes, tipo):
        """Guarda os parâmetros da execução e cria a fila interna usada
        para repassar mensagens de log da thread de trabalho para a
        thread da interface.

        Parâmetros:
            ano (str): ano selecionado para a execução.
            mes (str): mês selecionado para a execução.
            tipo (str): tipo do pacote de dados a baixar (TIPO_PACOTE).
        """
        super().__init__()
        self.ano = ano
        self.mes = mes
        self.tipo = tipo
        self.log_queue = queue.Queue()
        self.evento_parar = threading.Event()

    def solicitar_parada(self):
        """Sinaliza ao pipeline (rodando em executar_pipeline_completo)
        que ele deve parar assim que atingir o próximo ponto seguro de
        verificação.

        A parada é cooperativa, não instantânea: dependendo de onde a
        execução estiver no momento do clique (ex: no meio de um bloco
        de download), pode levar alguns segundos até o cancelamento
        de fato ser efetivado e o sinal cancelado_signal ser emitido.
        """
        self.evento_parar.set()

    def run(self):
        """Ponto de entrada da thread: dispara uma thread auxiliar para
        repassar os logs da fila para a interface (via _processar_fila)
        e executa o pipeline completo (executar_pipeline_completo).

        Se a execução for bem-sucedida, coloca um evento "done" na fila
        com os caminhos gerados; se o usuário tiver solicitado parada
        (ExecucaoCancelada), coloca um evento "cancelado"; se falhar por
        qualquer outro motivo, registra a exceção no log e coloca um
        evento "error" na fila com a mensagem de erro. Ao final, aguarda
        a thread de repasse de log terminar (com timeout).
        """
        thread_fila = threading.Thread(target=self._processar_fila, daemon=True)
        thread_fila.start()

        try:
            caminho, caminho_log = executar_pipeline_completo(
                self.ano,
                self.mes,
                self.tipo,
                self.log_queue,
                self.evento_parar,
            )
            self.log_queue.put(("done", (caminho, caminho_log)))
        except ExecucaoCancelada:
            self.log_queue.put(("cancelado", None))
        except Exception as erro:
            logging.getLogger("robo_siape").exception("Falha na execução.")
            self.log_queue.put(("error", str(erro)))

        thread_fila.join(timeout=2)

    def _processar_fila(self):
        """Loop executado em thread auxiliar que consome continuamente a
        fila de eventos (log_queue) enquanto a thread principal do
        Worker estiver rodando (ou enquanto ainda houver itens na fila),
        emitindo o sinal Qt correspondente a cada evento:
        - "log": emite log_signal com a mensagem;
        - "done": emite done_signal com os caminhos gerados e encerra o loop;
        - "error": emite error_signal com a mensagem de erro e encerra o loop;
        - "cancelado": emite cancelado_signal e encerra o loop.
        """
        while self.isRunning() or not self.log_queue.empty():
            try:
                tipo, valor = self.log_queue.get(timeout=0.15)
            except queue.Empty:
                continue

            if tipo == "log":
                self.log_signal.emit(valor)
            elif tipo == "done":
                caminho, caminho_log = valor
                self.done_signal.emit(caminho, caminho_log)
                break
            elif tipo == "error":
                self.error_signal.emit(valor)
                break
            elif tipo == "cancelado":
                self.cancelado_signal.emit()
                break


class DialogoAjuda(QDialog):
    """Diálogo de ajuda em formato de passo a passo explicativo.

    Mostra, uma tela por vez, um print real da interface já anotado
    com setas/realces indicando exatamente onde clicar, acompanhado de
    um texto curto explicando aquele passo. O usuário navega entre os
    passos com os botões "‹ Anterior" / "Próximo ›", e uma trilha de
    bolinhas no topo indica em qual passo ele está.
    """

    # Cada passo é (arquivo_da_imagem, título, texto explicativo).
    PASSOS = [
        (
            "passo1_painel.png",
            "1. Onde iniciar o robô",
            "No Painel de controle, clique no botão <b>\"CONFIGURAR ROBÔ →\"</b>, "
            "no card em destaque na parte de baixo do painel.<br><br>"
            "Ele leva direto para a tela de configuração da execução.",
        ),
        (
            "passo2_periodo.png",
            "2. Escolha o período",
            "Na tela <b>\"Nova execução\"</b>, selecione:<br><br>"
            "&nbsp;&nbsp;<b>1</b> — o <b>ano</b> da competência;<br>"
            "&nbsp;&nbsp;<b>2</b> — o <b>mês</b> da competência.<br><br>"
            "O pacote baixado é sempre o \"Servidores_SIAPE\". Períodos futuros "
            "ficam bloqueados automaticamente, seguindo a disponibilidade do "
            "Portal da Transparência.",
        ),
        (
            "passo3_iniciar.png",
            "3. Iniciar o robô",
            "Com o ano e o mês escolhidos, clique em <b>\"INICIAR ROBÔ →\"</b>.<br><br>"
            "A partir daí o robô baixa o pacote (modo headless, sem abrir "
            "navegador), valida e extrai o ZIP, trata o CSV de remuneração e "
            "gera a planilha Excel final, com cada etapa registrada em log, "
            "em tempo real, na tela seguinte.",
        ),
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ajuda — Robô SIAPE")
        self.setMinimumWidth(620)
        self.passo_atual = 0

        raiz = QVBoxLayout(self)
        raiz.setContentsMargins(28, 24, 28, 22)
        raiz.setSpacing(14)

        cabecalho = QLabel("🤖  Passo a passo: como iniciar uma execução")
        cabecalho.setObjectName("dialogTitle")
        cabecalho.setStyleSheet("font-size: 18px; font-weight: 900;")
        cabecalho.setWordWrap(True)
        raiz.addWidget(cabecalho)

        # Trilha de bolinhas indicando o passo atual
        self.trilha = QHBoxLayout()
        self.trilha.setSpacing(6)
        self.trilha.addStretch()
        self._bolinhas = []
        for i in range(len(self.PASSOS)):
            bolinha = QLabel("●")
            bolinha.setStyleSheet("font-size: 11px; color: #D9D1C3;")
            self._bolinhas.append(bolinha)
            self.trilha.addWidget(bolinha)
        self.trilha.addStretch()
        raiz.addLayout(self.trilha)

        self.titulo_passo = QLabel()
        self.titulo_passo.setObjectName("cardTitle")
        raiz.addWidget(self.titulo_passo)

        self.imagem_label = QLabel()
        self.imagem_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.imagem_label.setStyleSheet(
            "border: 1px solid #E3DCCF; border-radius: 10px; background: #FFFFFF;"
        )
        raiz.addWidget(self.imagem_label)

        self.texto_passo = QLabel()
        self.texto_passo.setObjectName("dialogText")
        self.texto_passo.setWordWrap(True)
        self.texto_passo.setTextFormat(Qt.TextFormat.RichText)
        raiz.addWidget(self.texto_passo)

        rodape = QHBoxLayout()
        self.btn_anterior = QPushButton("‹  Anterior")
        self.btn_anterior.setObjectName("dialogSecundario")
        self.btn_anterior.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_anterior.clicked.connect(self.passo_anterior)
        rodape.addWidget(self.btn_anterior)

        rodape.addStretch()

        self.btn_proximo = QPushButton()
        self.btn_proximo.setObjectName("dialogPrimary")
        self.btn_proximo.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_proximo.clicked.connect(self.proximo_ou_fechar)
        rodape.addWidget(self.btn_proximo)

        raiz.addLayout(rodape)

        self._atualizar_passo()

    def _caminho_imagem(self, nome_arquivo):
        return os.path.join(PASTA_ASSETS_AJUDA, nome_arquivo)

    def _atualizar_passo(self):
        """Redesenha o conteúdo do diálogo para refletir self.passo_atual:
        imagem anotada, título, texto explicativo, trilha de bolinhas e
        estado/rótulo dos botões de navegação.
        """
        arquivo, titulo, texto = self.PASSOS[self.passo_atual]

        self.titulo_passo.setText(titulo)
        self.texto_passo.setText(texto)

        caminho = self._caminho_imagem(arquivo)
        pixmap = QPixmap(caminho)
        if not pixmap.isNull():
            pixmap = pixmap.scaledToWidth(560, Qt.TransformationMode.SmoothTransformation)
            self.imagem_label.setPixmap(pixmap)
            self.imagem_label.setText("")
        else:
            self.imagem_label.setPixmap(QPixmap())
            self.imagem_label.setText(
                f"(imagem não encontrada: {os.path.join('assets', 'ajuda', arquivo)})"
            )

        for i, bolinha in enumerate(self._bolinhas):
            if i == self.passo_atual:
                bolinha.setStyleSheet("font-size: 11px; color: #D5A63A;")
            else:
                bolinha.setStyleSheet("font-size: 11px; color: #D9D1C3;")

        self.btn_anterior.setEnabled(self.passo_atual > 0)

        e_ultimo = self.passo_atual == len(self.PASSOS) - 1
        self.btn_proximo.setText("Entendi" if e_ultimo else "Próximo  ›")

    def passo_anterior(self):
        if self.passo_atual > 0:
            self.passo_atual -= 1
            self._atualizar_passo()

    def proximo_ou_fechar(self):
        if self.passo_atual < len(self.PASSOS) - 1:
            self.passo_atual += 1
            self._atualizar_passo()
        else:
            self.accept()


class App(QMainWindow):
    """Janela principal da aplicação (QMainWindow).

    Gerencia toda a navegação entre telas (login, cadastro, painel
    principal, configuração de nova execução e acompanhamento de
    execução) usando um QStackedWidget, além da barra lateral de
    navegação e o estado de login/execução em andamento.
    """

    def __init__(self):
        """Inicializa o estado da aplicação (usuário logado, worker,
        configurações), monta a interface (sidebar + área de conteúdo)
        e exibe a tela de login como ponto de partida.
        """
        super().__init__()

        self.usuario_logado = None
        self.worker = None
        self.config = carregar_config()
        self.tela = QStackedWidget()

        # Estado da execução do robô, mantido aqui (e não só nos widgets
        # da tela de execução) para que a execução sobreviva à navegação
        # entre telas: o usuário pode ir para o painel enquanto o robô
        # roda e voltar depois sem perder o log nem travar a interface.
        self.log_execucao_atual = []  # linhas de log acumuladas da execução atual/mais recente
        self.execucao_status_final = None  # None enquanto roda; "sucesso"/"erro"/"cancelado" ao terminar
        self.execucao_visivel = False  # True apenas quando a tela de execução está na frente

        self.setWindowTitle("Robô SIAPE")
        preparar_pastas()
        self.setMinimumSize(1250, 750) #Não deixa o usuário diminuir a interface
        self.setMaximumSize(1250, 750) #Não deixa o usuário aumentar a interface
        self.resize(1250, 750)

        # Remove o botão de maximizar (e o duplo clique na barra de
        # título que também maximiza): a janela maximizada estava
        # bugando o layout, então a janela continua redimensionável
        # normalmente pelas bordas, só não pode mais ser maximizada.
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowMaximizeButtonHint)

        self._montar_interface()
        self.tela_login()

    def _montar_interface(self):
        """Monta o esqueleto visual da janela: um layout horizontal com a
        barra lateral (sidebar) à esquerda e a área de conteúdo
        (QStackedWidget "tela") à direita, definindo a janela central.
        """
        central = QWidget()
        central_layout = QHBoxLayout(central)
        central_layout.setContentsMargins(0, 0, 0, 0)
        central_layout.setSpacing(0)

        self.sidebar = self._criar_sidebar()
        central_layout.addWidget(self.sidebar)

        self.conteudo = QFrame()
        self.conteudo.setObjectName("contentArea")
        content_layout = QVBoxLayout(self.conteudo)
        content_layout.setContentsMargins(32, 26, 32, 24)

        # A tela fica dentro de um QScrollArea: assim, se a janela for
        # deixada em modo "janela" (não maximizada) menor do que o
        # conteúdo precisa, aparece uma barra de rolagem em vez dos
        # cartões e botões se espremerem/desalinharem.
        self.scroll_area = QScrollArea()
        self.scroll_area.setObjectName("mainScrollArea")
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setWidget(self.tela)
        content_layout.addWidget(self.scroll_area)

        central_layout.addWidget(self.conteudo, 1)
        self.setCentralWidget(central)

    def _criar_sidebar(self):
        """Constrói a barra lateral de navegação: logotipo e nome do app,
        separador, seção "SISTEMA" com o botão de navegação (Painel),
        um cartão mostrando o usuário conectado, e o botão de sair.

        Retorna:
            QFrame: o frame completo da barra lateral, pronto para ser
            adicionado ao layout principal.
        """
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(210)

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(24, 30, 24, 28)
        layout.setSpacing(8)

        topo = QHBoxLayout()
        topo.setSpacing(12)
        topo.addWidget(LogoMark())

        marca = QVBoxLayout()
        marca.setSpacing(0)
        nome = QLabel("ROBÔ SIAPE")
        nome.setObjectName("brandName")
        marca.addWidget(nome)
        topo.addLayout(marca)
        layout.addLayout(topo)

        linha = QFrame()
        linha.setFrameShape(QFrame.Shape.HLine)
        linha.setObjectName("separator")
        layout.addWidget(linha)
        layout.addSpacing(12)

        self.nav_status = QLabel("●  SISTEMA")
        self.nav_status.setObjectName("sideSection")
        layout.addWidget(self.nav_status)

        self.nav_home = self._nav_button("⌂", "Painel")
        self.nav_home.clicked.connect(self.tela_principal)
        layout.addWidget(self.nav_home)

        layout.addStretch()

        self.user_card = QFrame()
        self.user_card.setObjectName("userCard")
        user_layout = QVBoxLayout(self.user_card)
        user_layout.setContentsMargins(14, 13, 14, 13)

        user_label = QLabel("USUÁRIO CONECTADO")
        user_label.setObjectName("sideCaption")
        user_layout.addWidget(user_label)

        self.side_user = QLabel("—")
        self.side_user.setObjectName("sideUser")
        user_layout.addWidget(self.side_user)

        layout.addWidget(self.user_card)

        sair = self._nav_button("↪", "Sair")
        sair.clicked.connect(self.tela_login)
        layout.addWidget(sair)

        return sidebar

    def _nav_button(self, icon, texto):
        """Cria um botão de navegação padrão da sidebar, combinando um
        ícone (emoji/caractere) com um texto.

        Parâmetros:
            icon (str): caractere/emoji usado como ícone do botão.
            texto (str): rótulo textual do botão.

        Retorna:
            QPushButton: botão pronto, ainda sem ação conectada (o
            chamador deve conectar o sinal `clicked`).
        """
        botao = QPushButton()
        botao.setText(f"{icon}   {texto}")
        botao.setCursor(Qt.CursorShape.PointingHandCursor)
        botao.setMinimumHeight(40)
        botao.setObjectName("navButton")
        return botao

    def _definir_usuario(self):
        """Atualiza o rótulo do usuário exibido na sidebar com o nome do
        usuário atualmente logado (ou "—" se nenhum estiver logado).
        """
        self.side_user.setText(self.usuario_logado or "—")

    def limpar_tela(self):
        """Remove e destrói todos os widgets atualmente empilhados no
        QStackedWidget "tela", preparando o terreno para desenhar uma
        nova tela do zero (evita telas antigas acumuladas em memória).

        Também marca a tela de execução como não visível: se o Worker
        emitir um sinal (log/conclusão/erro/cancelamento) enquanto o
        usuário estiver em outra tela, os manipuladores sabem que não
        devem tentar atualizar widgets que acabaram de ser destruídos
        aqui (o que antes derrubava a interface).
        """
        self.execucao_visivel = False
        while self.tela.count():
            widget = self.tela.widget(0)
            self.tela.removeWidget(widget)
            widget.deleteLater()

    def _pagina(self):
        """Cria uma nova página vazia (QWidget com QVBoxLayout) e já a
        adiciona ao QStackedWidget "tela", retornando ambos para que o
        método chamador possa continuar montando o conteúdo da página.

        Retorna:
            tuple: (pagina, layout) — o widget da página e seu layout
            vertical, já configurado com espaçamento e margens padrão.
        """
        pagina = QWidget()
        layout = QVBoxLayout(pagina)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(20)
        self.tela.addWidget(pagina)
        return pagina, layout

    def _label(self, texto):
        """Cria um QLabel estilizado como rótulo de campo de formulário
        (objectName "fieldLabel"), usado acima de campos de entrada
        (ex: "USUÁRIO", "SENHA", "ANO").

        Parâmetros:
            texto (str): texto do rótulo.

        Retorna:
            QLabel: rótulo pronto para ser adicionado a um layout.
        """
        label = QLabel(texto)
        label.setObjectName("fieldLabel")
        return label

    def _mensagem(self, titulo, texto, tipo="info"):
        """Exibe uma caixa de diálogo modal (QMessageBox) com título,
        mensagem e ícone apropriado ao tipo informado.

        Parâmetros:
            titulo (str): título da janela de diálogo.
            texto (str): corpo da mensagem exibida ao usuário.
            tipo (str): um de "error", "warning", "success" ou "info"
                (padrão), que define o ícone exibido.
        """
        box = QMessageBox(self)
        box.setWindowTitle(titulo)
        box.setText(texto)
        box.setIcon(
            {
                "error": QMessageBox.Icon.Critical,
                "warning": QMessageBox.Icon.Warning,
                "success": QMessageBox.Icon.Information,
                "info": QMessageBox.Icon.Information,
            }.get(tipo, QMessageBox.Icon.Information)
        )
        box.exec()

    def tela_login(self):
        """Monta e exibe a tela de login: esconde a sidebar (usuário
        ainda não autenticado), cria um cartão centralizado com logotipo,
        título, campos de usuário e senha (com Enter disparando o login),
        botão "ENTRAR" e link para a tela de cadastro.
        """
        self.limpar_tela()
        self.sidebar.hide()

        pagina, layout = self._pagina()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        card = Card()
        card.setMaximumWidth(500)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(48, 44, 48, 44)
        card_layout.setSpacing(15)

        logo = LogoMark()
        card_layout.addWidget(logo, alignment=Qt.AlignmentFlag.AlignHCenter)

        titulo = QLabel("Bem-vindo de volta")
        titulo.setObjectName("authTitle")
        titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(titulo)

        subtitulo = QLabel("Acesse o painel de automação do SIAPE")
        subtitulo.setObjectName("authSubtitle")
        subtitulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(subtitulo)
        card_layout.addSpacing(14)

        card_layout.addWidget(self._label("USUÁRIO"))
        self.entry_usuario = Campo("Digite seu usuário")
        self.entry_usuario.returnPressed.connect(self.fazer_login)
        card_layout.addWidget(self.entry_usuario)

        card_layout.addWidget(self._label("SENHA"))
        self.entry_senha = Campo("Digite sua senha", senha=True)
        self.entry_senha.returnPressed.connect(self.fazer_login)
        card_layout.addWidget(self.entry_senha)

        entrar = Botao("ENTRAR  →", principal=True)
        entrar.clicked.connect(self.fazer_login)
        card_layout.addWidget(entrar)

        cadastro = QPushButton("Ainda não possui acesso?  Criar usuário")
        cadastro.setObjectName("linkButton")
        cadastro.clicked.connect(self.tela_cadastro)
        card_layout.addWidget(cadastro, alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(card)

    def fazer_login(self):
        """Valida as credenciais digitadas nos campos de usuário/senha da
        tela de login e, se corretas, define o usuário como logado e
        navega para a tela principal.

        Fluxo de validação:
        1. Verifica se usuário e senha foram preenchidos.
        2. Verifica se o usuário existe na base local (usuarios.json).
        3. Compara o hash da senha digitada com o hash armazenado.
        Em qualquer falha, exibe uma mensagem de erro/aviso e interrompe
        o login.
        """
        usuario = self.entry_usuario.text().strip()
        senha = self.entry_senha.text()

        if not usuario or not senha:
            self._mensagem("Atenção", "Preencha usuário e senha.", "warning")
            return

        usuarios = carregar_usuarios()

        if usuario not in usuarios:
            self._mensagem("Acesso negado", "Usuário não cadastrado.", "error")
            return

        if usuarios[usuario]["senha"] != hash_senha(senha):
            self._mensagem("Acesso negado", "Senha incorreta.", "error")
            return

        self.usuario_logado = usuario
        self._definir_usuario()
        self.tela_principal()

    def tela_cadastro(self):
        """Monta e exibe a tela de criação de novo usuário: cartão
        centralizado com campos de usuário, senha e confirmação de
        senha, botão "CRIAR USUÁRIO" e link para voltar ao login.
        """
        self.limpar_tela()
        self.sidebar.hide()

        pagina, layout = self._pagina()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        card = Card()
        card.setMaximumWidth(500)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(48, 40, 48, 40)
        card_layout.setSpacing(14)

        logo = LogoMark()
        card_layout.addWidget(logo, alignment=Qt.AlignmentFlag.AlignHCenter)

        titulo = QLabel("Criar acesso")
        titulo.setObjectName("authTitle")
        titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(titulo)

        subtitulo = QLabel("Cadastre um usuário local para acessar o sistema")
        subtitulo.setObjectName("authSubtitle")
        subtitulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(subtitulo)
        card_layout.addSpacing(12)

        card_layout.addWidget(self._label("USUÁRIO"))
        self.novo_usuario = Campo("3 a 30 caracteres")
        card_layout.addWidget(self.novo_usuario)

        card_layout.addWidget(self._label("SENHA"))
        self.nova_senha = Campo("Crie uma senha", senha=True)
        card_layout.addWidget(self.nova_senha)

        card_layout.addWidget(self._label("CONFIRMAR SENHA"))
        self.confirma_senha = Campo("Repita a senha", senha=True)
        card_layout.addWidget(self.confirma_senha)

        cadastrar = Botao("CRIAR USUÁRIO  →", principal=True)
        cadastrar.clicked.connect(self.cadastrar_usuario)
        card_layout.addWidget(cadastrar)

        voltar = QPushButton("← Voltar ao login")
        voltar.setObjectName("linkButton")
        voltar.clicked.connect(self.tela_login)
        card_layout.addWidget(voltar, alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(card)

    def cadastrar_usuario(self):
        """Valida os dados preenchidos na tela de cadastro e, se tudo
        estiver correto, cria o novo usuário na base local.

        Fluxo de validação:
        1. Verifica se todos os campos foram preenchidos.
        2. Verifica se o nome de usuário atende ao formato (usuario_valido).
        3. Verifica se a senha e a confirmação coincidem.
        4. Verifica se o nome de usuário já não está em uso.
        Se aprovado em todas as etapas, salva o usuário (com a senha em
        hash) e retorna à tela de login com mensagem de sucesso.
        """
        usuario = self.novo_usuario.text().strip()
        senha = self.nova_senha.text()
        confirma = self.confirma_senha.text()

        if not usuario or not senha or not confirma:
            self._mensagem("Atenção", "Preencha todos os campos.", "warning")
            return

        if not usuario_valido(usuario):
            self._mensagem(
                "Usuário inválido",
                "Use de 3 a 30 caracteres: letras, números, _ , . ou -.",
                "error",
            )
            return

        if senha != confirma:
            self._mensagem("Erro", "As senhas não coincidem.", "error")
            return

        usuarios = carregar_usuarios()

        if usuario in usuarios:
            self._mensagem("Erro", "Esse usuário já está cadastrado.", "error")
            return

        usuarios[usuario] = {"senha": hash_senha(senha)}
        salvar_usuarios(usuarios)

        self._mensagem("Tudo certo", "Cadastro realizado com sucesso!", "success")
        self.tela_login()

    def tela_principal(self):
        """Monta e exibe o painel principal (dashboard) do app, com:
        cabeçalho e indicador de status (executando/pronto); uma linha
        com o cartão de status do robô, atividade recente (gráfico de
        barras) e ações rápidas; uma linha com atalhos e preferências
        (toggles); uma linha com a lista de arquivos gerados e os
        indicadores circulares (gauges); e um cartão de destaque que
        convida a iniciar uma nova execução (ou acompanhar a atual, se
        já houver uma em andamento).
        """
        self.limpar_tela()
        self.sidebar.show()
        self._definir_usuario()
        self.config = carregar_config()

        em_execucao = bool(self.worker and self.worker.isRunning())

        pagina, layout = self._pagina()
        layout.setSpacing(12)

        # ---------------------------------------------------------------
        # Cabeçalho
        # ---------------------------------------------------------------
        topo = QHBoxLayout()
        topo.addWidget(
            criar_titulo(
                "Painel de controle",
                "Baixe, trate e organize a remuneração do SIAPE em poucos cliques — "
                "direto do Portal da Transparência.",
            )
        )
        topo.addStretch()

        badge = QLabel("●  EXECUTANDO" if em_execucao else "●  PRONTO")
        badge.setObjectName("statusBadgeRunning" if em_execucao else "statusBadge")
        topo.addWidget(badge, alignment=Qt.AlignmentFlag.AlignTop)
        layout.addLayout(topo)

        # ---------------------------------------------------------------
        # Linha 1 — status do robô / atividade recente / ações rápidas
        # ---------------------------------------------------------------
        linha1 = QHBoxLayout()
        linha1.setSpacing(12)
        linha1.addWidget(self._card_status_robo(em_execucao), 3)
        linha1.addWidget(self._card_atividade(), 3)
        linha1.addWidget(self._card_acoes_rapidas(), 3)
        layout.addLayout(linha1)

        # ---------------------------------------------------------------
        # Linha 2 — barra de atalhos + preferências (toggles reais)
        # ---------------------------------------------------------------
        linha2 = QHBoxLayout()
        linha2.setSpacing(12)
        linha2.addWidget(self._barra_atalhos(), 3)
        linha2.addWidget(self._painel_preferencias(), 1)
        layout.addLayout(linha2)

        # ---------------------------------------------------------------
        # Linha 3 — arquivos gerados + indicadores (gauges reais)
        # ---------------------------------------------------------------
        linha3 = QHBoxLayout()
        linha3.setSpacing(12)
        linha3.addWidget(self._card_arquivos(), 2)
        linha3.addWidget(self._card_gauges(), 1)
        layout.addLayout(linha3)

        # ---------------------------------------------------------------
        # Chamada para nova execução
        # ---------------------------------------------------------------
        destaque = Card()
        dlayout = QHBoxLayout(destaque)
        dlayout.setContentsMargins(26, 22, 26, 22)

        left = QVBoxLayout()
        tag = QLabel("AUTOMAÇÃO SIAPE")
        tag.setObjectName("eyebrow")
        left.addWidget(tag)

        titulo = QLabel(
            "Robô em execução no momento" if em_execucao else "Pronto para iniciar uma nova execução?"
        )
        titulo.setObjectName("cardTitle")
        left.addWidget(titulo)

        desc = QLabel(
            "Em poucos cliques, o robô baixa o pacote certo direto do Portal da "
            "Transparência (modo headless, sem abrir navegador), remove o que não "
            "interessa, trata os valores monetários e entrega uma planilha Excel "
            "pronta para usar — com cada etapa registrada em log."
        )
        desc.setObjectName("cardText")
        desc.setWordWrap(True)
        left.addWidget(desc)
        dlayout.addLayout(left, 1)

        # O botão de Ajuda que ficava aqui foi removido: já existe o
        # mesmo atalho na barra "Resumo dos logs / Perfil / Ajuda"
        # logo acima, e duplicar a ação nos dois lugares não fazia
        # sentido. Sobra só o botão principal, centralizado
        # verticalmente neste bloco.
        iniciar = Botao(
            "ACOMPANHAR EXECUÇÃO  →" if em_execucao else "CONFIGURAR ROBÔ  →",
            principal=True,
        )
        iniciar.setMinimumWidth(230)
        iniciar.clicked.connect(
            (lambda: self.tela_execucao(self.worker.ano, self.worker.mes, self.worker.tipo, novo=False))
            if em_execucao
            else self.tela_config_robo
        )
        dlayout.addWidget(iniciar, 0, Qt.AlignmentFlag.AlignVCenter)

        layout.addWidget(destaque)

    # ---------------------------------------------------------------
    # Cartões do painel
    # ---------------------------------------------------------------
    def _card_status_robo(self, em_execucao):
        """Monta o cartão escuro de status do robô, exibindo se está
        "Executando" ou "Pronto", os dados da última execução registrada
        (mês/ano e data, formatados) e o nome do usuário conectado.

        Parâmetros:
            em_execucao (bool): se True, exibe o estado "Executando";
                caso contrário, exibe "Pronto".

        Retorna:
            Card: cartão pronto para ser adicionado ao layout do painel.
        """
        card = Card()
        card.setObjectName("darkCard")
        card.setMinimumHeight(160)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(20, 18, 20, 16)
        lay.setSpacing(8)

        topo = QHBoxLayout()
        tag = QLabel("STATUS DO ROBÔ")
        tag.setObjectName("darkEyebrow")
        topo.addWidget(tag)
        topo.addStretch()
        lay.addLayout(topo)

        valor = QLabel("Executando" if em_execucao else "Pronto")
        valor.setObjectName("darkValue")
        lay.addWidget(valor)

        ultima = ultima_execucao()
        if ultima:
            try:
                data_fmt = datetime.fromisoformat(ultima["data"]).strftime("%d/%m/%Y")
            except ValueError:
                data_fmt = ultima["data"][:10]
            legenda = (
                f"Última execução: {ultima['mes'].capitalize()}/{ultima['ano']} "
                f"({data_fmt})"
            )
        else:
            legenda = "Nenhuma execução registrada ainda."
        info = QLabel(legenda)
        info.setObjectName("darkCaption")
        info.setWordWrap(True)
        lay.addWidget(info)

        lay.addStretch()

        rodape = QLabel(f"Usuário conectado: {self.usuario_logado or '—'}")
        rodape.setObjectName("darkFoot")
        lay.addWidget(rodape)
        return card

    def _card_atividade(self):
        """Monta o cartão de atividade recente: quantidade total de
        execuções registradas e o mini gráfico de barras (BarChartWidget)
        com a contagem de execuções nos últimos 6 meses.

        Retorna:
            Card: cartão pronto para ser adicionado ao layout do painel.
        """
        card = Card()
        card.setMinimumHeight(160)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(20, 18, 20, 16)
        lay.setSpacing(6)

        tag = QLabel("ATIVIDADE")
        tag.setObjectName("eyebrow")
        lay.addWidget(tag)

        total, sucesso, taxa = estatisticas_execucoes()
        titulo = QLabel(f"{total} execuções registradas")
        titulo.setObjectName("cardTitle")
        lay.addWidget(titulo)

        rotulos, valores = contagem_execucoes_por_dia()
        self.grafico_atividade = BarChartWidget(valores, rotulos)
        lay.addWidget(self.grafico_atividade, 1)

        legenda = QLabel("Execuções por dia (últimos 6 dias com execução)")
        legenda.setObjectName("cardText")
        lay.addWidget(legenda)
        return card

    def _card_acoes_rapidas(self):
        """Monta o cartão de ações rápidas do painel, com botões para:
        abrir a pasta de saída, atualizar o painel (recarregando a tela
        principal) e abrir o log da última execução.

        Retorna:
            Card: cartão pronto para ser adicionado ao layout do painel.
        """
        card = Card()
        card.setMinimumHeight(160)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(20, 18, 20, 16)
        lay.setSpacing(8)

        tag = QLabel("AÇÕES RÁPIDAS")
        tag.setObjectName("eyebrow")
        lay.addWidget(tag)

        abrir = Botao("Abrir pasta de saída", compacto=True)
        abrir.clicked.connect(self.abrir_pasta_saida)
        lay.addWidget(abrir)

        atualizar = Botao("Atualizar painel", compacto=True)
        atualizar.clicked.connect(self.tela_principal)
        lay.addWidget(atualizar)

        ver_log = Botao("Ver último log", compacto=True)
        ver_log.clicked.connect(self.abrir_ultimo_log)
        lay.addWidget(ver_log)

        lay.addStretch()
        return card

    def abrir_ultimo_log(self):
        """Localiza o arquivo de log (.log) mais recente dentro da pasta
        de saída (PASTA_SAIDA, ignorando "_processamento") e o abre no
        aplicativo padrão do sistema operacional.

        Se a pasta de saída ainda não existir ou nenhum log for
        encontrado, exibe uma mensagem informativa em vez de abrir algo.
        """
        if not os.path.isdir(PASTA_SAIDA):
            self._mensagem("Sem execuções", "Ainda não há nenhuma execução registrada.", "info")
            return

        logs = []
        for raiz, pastas, arquivos in os.walk(PASTA_SAIDA):
            if "_processamento" in pastas:
                pastas.remove("_processamento")
            for nome in arquivos:
                if nome.lower().endswith(".log"):
                    caminho = os.path.join(raiz, nome)
                    logs.append((os.path.getmtime(caminho), caminho))

        if not logs:
            self._mensagem("Sem execuções", "Ainda não há nenhum log registrado.", "info")
            return

        logs.sort(key=lambda item: item[0], reverse=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(logs[0][1]))

    def resumo_logs(self):
        """Gera (ou atualiza) o arquivo de resumo com todas as execuções
        já registradas (via gerar_resumo_logs) e o abre no aplicativo
        padrão do sistema (ex: bloco de notas).

        Se ainda não houver nenhum log registrado, exibe uma mensagem
        informativa em vez de tentar abrir um arquivo inexistente.
        """
        caminho = gerar_resumo_logs()
        if not caminho:
            self._mensagem(
                "Sem logs",
                "Ainda não há nenhum log registrado. Faça uma execução primeiro.",
                "info",
            )
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(caminho))

    def abrir_ajuda(self):
        """Abre o diálogo de ajuda em formato de passo a passo (DialogoAjuda),
        com prints reais da interface anotados com setas/realces mostrando
        exatamente onde clicar para iniciar uma execução do robô.
        """
        dialogo = DialogoAjuda(self)
        dialogo.exec()

    def _barra_atalhos(self):
        """Monta uma barra horizontal de atalhos rápidos (resumo dos
        logs, perfil do usuário e ajuda), exibida em formato de cartão.

        Fica só com ações que não existem em nenhum outro lugar do
        painel — "Nova execução", "Pasta de saída" e "Atualizar" já
        aparecem no card de Ações Rápidas, então saíram daqui para não
        duplicar. Usa o mesmo estilo "compacto" (sem "grande") dos
        botões do card de Ações Rápidas, para manter a mesma altura e
        peso visual entre os dois blocos.

        Retorna:
            Card: cartão pronto para ser adicionado ao layout do painel.
        """
        card = Card()
        lay = QHBoxLayout(card)
        lay.setContentsMargins(16, 12, 16, 12)
        lay.setSpacing(12)

        atalhos = [
            ("☰", "Resumo dos logs", self.resumo_logs),
            ("◉", "Perfil", lambda: self._mensagem(
                "Usuário", f"Conectado como: {self.usuario_logado or '—'}", "info"
            )),
            ("？", "Ajuda", self.abrir_ajuda),
        ]
        for icone, texto, acao in atalhos:
            botao = Botao(f"{icone}   {texto}", compacto=True)
            botao.clicked.connect(acao)
            lay.addWidget(botao, 1)

        return card

    def _painel_preferencias(self):
        """Monta o cartão de preferências do usuário, com dois
        interruptores (ToggleSwitch) reais: "Notificar ao concluir" e
        "Abrir pasta automaticamente", já refletindo o valor atual salvo
        em config.json e persistindo qualquer alteração imediatamente.

        Retorna:
            Card: cartão pronto para ser adicionado ao layout do painel.
        """
        card = Card()
        lay = QVBoxLayout(card)
        lay.setContentsMargins(18, 12, 18, 12)
        lay.setSpacing(10)

        linha1 = QHBoxLayout()
        rotulo1 = QLabel("Notificar ao concluir")
        rotulo1.setObjectName("cardText")
        linha1.addWidget(rotulo1, 1)
        toggle1 = ToggleSwitch(self.config.get("notificar_conclusao", True))
        toggle1.toggled.connect(lambda marcado: self._alterar_preferencia("notificar_conclusao", marcado))
        linha1.addWidget(toggle1)
        lay.addLayout(linha1)

        linha2 = QHBoxLayout()
        rotulo2 = QLabel("Abrir pasta automaticamente")
        rotulo2.setObjectName("cardText")
        linha2.addWidget(rotulo2, 1)
        toggle2 = ToggleSwitch(self.config.get("abrir_pasta_automatico", False))
        toggle2.toggled.connect(lambda marcado: self._alterar_preferencia("abrir_pasta_automatico", marcado))
        linha2.addWidget(toggle2)
        lay.addLayout(linha2)

        return card

    def _alterar_preferencia(self, chave, valor):
        """Atualiza uma chave específica das preferências em memória e
        persiste imediatamente o dicionário completo em config.json.

        Parâmetros:
            chave (str): nome da preferência a alterar (ex:
                "notificar_conclusao").
            valor: novo valor da preferência (geralmente bool).
        """
        self.config[chave] = valor
        salvar_config(self.config)

    def _card_arquivos(self):
        """Monta o cartão com a lista dos arquivos finais (.xlsx) mais
        recentemente gerados pelo robô (via listar_arquivos_saida),
        exibindo uma linha clicável para cada um. Se não houver nenhum
        arquivo ainda, exibe uma mensagem explicativa.

        Retorna:
            Card: cartão pronto para ser adicionado ao layout do painel.
        """
        card = Card()
        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 18, 22, 18)
        lay.setSpacing(10)

        cabecalho = QHBoxLayout()
        tag = QLabel("📁  ARQUIVOS FINAIS DA EXECUÇÃO")
        tag.setObjectName("eyebrow")
        cabecalho.addWidget(tag)
        cabecalho.addStretch()
        lay.addLayout(cabecalho)

        arquivos = listar_arquivos_saida()
        if not arquivos:
            vazio = QLabel("Nenhum arquivo gerado ainda. Inicie uma execução para ver os resultados aqui.")
            vazio.setObjectName("cardText")
            vazio.setWordWrap(True)
            lay.addWidget(vazio)
        else:
            for item in arquivos:
                lay.addWidget(self._linha_arquivo(item))

        lay.addStretch()
        return card

    def _linha_arquivo(self, item):
        """Cria uma linha clicável representando um único arquivo gerado,
        mostrando nome, tamanho e data. Um clique na linha abre o
        arquivo no aplicativo padrão do sistema (ex: Excel).

        Parâmetros:
            item (dict): dicionário com as chaves "nome", "caminho",
                "tamanho" e "data" (formato retornado por
                listar_arquivos_saida).

        Retorna:
            QFrame: linha pronta para ser adicionada ao cartão de
            arquivos.
        """
        linha = QFrame()
        linha.setObjectName("fileRow")
        linha.setCursor(Qt.CursorShape.PointingHandCursor)
        lay = QHBoxLayout(linha)
        lay.setContentsMargins(12, 9, 12, 9)

        nome = QLabel(f"📄  {item['nome']}")
        nome.setObjectName("fileName")
        lay.addWidget(nome, 1)

        detalhes = QLabel(f"{item['tamanho']}   •   {item['data']}")
        detalhes.setObjectName("fileMeta")
        lay.addWidget(detalhes)

        def abrir(_evento, caminho=item["caminho"]):
            """Manipulador do clique na linha: abre o arquivo indicado
            (caminho fixado por padrão ao criar a função) no aplicativo
            padrão do sistema operacional.
            """
            QDesktopServices.openUrl(QUrl.fromLocalFile(caminho))

        linha.mousePressEvent = abrir
        return linha

    def _card_gauges(self):
        """Monta o cartão com os dois indicadores circulares (gauges):
        total de execuções registradas e taxa de sucesso em percentual,
        calculados a partir de estatisticas_execucoes().

        Retorna:
            Card: cartão pronto para ser adicionado ao layout do painel.
        """
        card = Card()
        lay = QHBoxLayout(card)
        lay.setContentsMargins(10, 14, 10, 10)
        lay.setSpacing(6)

        total, sucesso, taxa = estatisticas_execucoes()

        gauge_total = GaugeWidget(total, max(total, 1), sufixo="", rotulo="Execuções")
        gauge_taxa = GaugeWidget(taxa, 100, sufixo="%", rotulo="Taxa de sucesso")

        lay.addWidget(gauge_total)
        lay.addWidget(gauge_taxa)
        return card

    def abrir_pasta_saida(self):
        """Garante que a pasta de saída exista e a abre no explorador de
        arquivos padrão do sistema operacional.
        """
        preparar_pastas()
        QDesktopServices.openUrl(QUrl.fromLocalFile(PASTA_SAIDA))

    def tela_config_robo(self):
        """Monta e exibe a tela de configuração de uma nova execução:
        formulário para escolher ano e mês (o combo de mês é atualizado
        dinamicamente conforme o ano escolhido), um aviso sobre o
        funcionamento do download, botões de voltar/iniciar, e um
        painel lateral resumindo visualmente as 5 etapas do fluxo do
        robô (download, validação, extração, tratamento, geração).

        Se já houver uma execução em andamento, não abre o formulário
        (isso poderia disparar um segundo Worker rodando em paralelo);
        em vez disso, avisa o usuário e o leva de volta ao
        acompanhamento da execução atual.
        """
        if self.worker and self.worker.isRunning():
            self._mensagem(
                "Execução em andamento",
                "Já existe uma execução em andamento. Aguarde ela terminar "
                "ou clique em \"Parar\" antes de iniciar uma nova.",
                "warning",
            )
            self.tela_execucao(self.worker.ano, self.worker.mes, self.worker.tipo, novo=False)
            return

        self.limpar_tela()
        self.sidebar.show()
        self._definir_usuario()

        pagina, layout = self._pagina()
        layout.addWidget(
            criar_titulo(
                "Nova execução",
                "Defina o período e o pacote de dados que serão baixados e processados.",
            )
        )

        centro = QHBoxLayout()
        centro.setSpacing(22)

        card = Card()
        card.setMaximumWidth(650)
        form = QVBoxLayout(card)
        form.setContentsMargins(30, 30, 30, 30)
        form.setSpacing(12)

        titulo = QLabel("Parâmetros da execução")
        titulo.setObjectName("cardTitle")
        form.addWidget(titulo)

        texto = QLabel(f"Selecione o período. O pacote baixado é sempre {TIPO_PACOTE}.")
        texto.setObjectName("cardText")
        texto.setWordWrap(True)
        form.addWidget(texto)
        form.addSpacing(8)

        form.addWidget(self._label("ANO"))
        self.combo_ano = QComboBox()
        self.combo_ano.addItems(sorted(INDICES_ANOS.keys(), reverse=True))
        self.combo_ano.currentTextChanged.connect(self.atualizar_meses)
        estilizar_combo(self.combo_ano)
        form.addWidget(self.combo_ano)

        form.addWidget(self._label("MÊS"))
        self.combo_mes = QComboBox()
        estilizar_combo(self.combo_mes)
        form.addWidget(self.combo_mes)

        form.addSpacing(12)

        aviso = QLabel(
            "O download é feito diretamente do Portal da Transparência (sem abrir "
            "navegador). O sistema bloqueia automaticamente períodos futuros e "
            "mantém as regras de disponibilidade do portal."
        )
        aviso.setObjectName("hint")
        aviso.setWordWrap(True)
        form.addWidget(aviso)

        botoes = QHBoxLayout()
        voltar = Botao("← Voltar", compacto=True)
        voltar.clicked.connect(self.tela_principal)
        iniciar = Botao("INICIAR ROBÔ  →", principal=True, compacto=True)
        iniciar.clicked.connect(self.iniciar_robo)
        botoes.addWidget(voltar)
        botoes.addStretch()
        botoes.addWidget(iniciar)
        form.addLayout(botoes)

        centro.addWidget(card)

        resumo = Card()
        resumo.setMaximumWidth(340)
        rlay = QVBoxLayout(resumo)
        rlay.setContentsMargins(25, 25, 25, 25)

        tag = QLabel("FLUXO")
        tag.setObjectName("eyebrow")
        rlay.addWidget(tag)

        for numero, texto in [
            ("01", "Baixar o pacote escolhido (headless)"),
            ("02", "Validar o arquivo ZIP recebido"),
            ("03", "Extrair e limpar arquivos desnecessários"),
            ("04", "Tratar CSV de remuneração"),
            ("05", "Gerar Excel + log"),
        ]:
            row = QHBoxLayout()
            n = QLabel(numero)
            n.setObjectName("flowNumber")
            n.setFixedWidth(35)
            row.addWidget(n)
            t = QLabel(texto)
            t.setObjectName("flowText")
            t.setWordWrap(True)
            row.addWidget(t, 1)
            rlay.addLayout(row)

        rlay.addStretch()
        centro.addWidget(resumo)
        layout.addLayout(centro)
        layout.addStretch()

        self.atualizar_meses()

    def atualizar_meses(self):
        """Atualiza a lista de meses disponíveis no combo de mês,
        conforme o ano atualmente selecionado no combo de ano (chamando
        meses_disponiveis). Disparado automaticamente sempre que o ano
        selecionado muda.
        """
        ano = self.combo_ano.currentText()
        self.combo_mes.clear()
        self.combo_mes.addItems(meses_disponiveis(ano))

    def validar_periodo(self, ano, mes, tipo):
        """Valida se a combinação de ano/mês/tipo escolhida na interface
        pode ser executada, checando:
        1. Se o ano está entre os anos conhecidos (INDICES_ANOS).
        2. Se o mês está disponível para aquele ano
           (meses_disponiveis).
        3. Se o período não está no futuro em relação à data atual.

        Parâmetros:
            ano (str): ano selecionado.
            mes (str): mês selecionado, por extenso.
            tipo (str): tipo do pacote (não usado na validação atual,
                mantido por compatibilidade de assinatura).

        Retorna:
            tuple: (valido, mensagem) onde valido (bool) indica se o
            período pode ser executado, e mensagem (str) contém o motivo
            da rejeição (string vazia se válido).
        """
        if ano not in INDICES_ANOS:
            return False, "Ano inválido."

        if mes not in meses_disponiveis(ano):
            return False, "Mês inválido ou ainda não disponível para esse ano."

        numero_mes = {
            "janeiro": 1,
            "fevereiro": 2,
            "março": 3,
            "abril": 4,
            "maio": 5,
            "junho": 6,
            "julho": 7,
            "agosto": 8,
            "setembro": 9,
            "outubro": 10,
            "novembro": 11,
            "dezembro": 12,
        }[mes]

        if (int(ano), numero_mes) > (date.today().year, date.today().month):
            return False, "Não é permitido selecionar um mês futuro."

        return True, ""

    def iniciar_robo(self):
        """Lê o ano e o mês escolhidos no formulário de configuração,
        valida se estão preenchidos e se o período é válido
        (validar_periodo) e, se tudo estiver correto, navega para a tela
        de execução (tela_execucao), que efetivamente dispara o robô.
        Caso contrário, exibe uma mensagem de erro/aviso explicando o
        motivo.
        """
        ano = self.combo_ano.currentText()
        mes = self.combo_mes.currentText()
        tipo = TIPO_PACOTE

        if not ano or not mes:
            self._mensagem("Atenção", "Selecione o ano e o mês.", "warning")
            return

        valido, mensagem = self.validar_periodo(ano, mes, tipo)

        if not valido:
            self._mensagem("Entrada inválida", mensagem, "error")
            return

        self.tela_execucao(ano, mes, tipo, novo=True)

    def tela_execucao(self, ano, mes, tipo, novo=True):
        """Monta e exibe a tela de acompanhamento de execução.

        Monta: cabeçalho com o status; cartão com o período/tipo
        selecionado e uma dica sobre o funcionamento do download; um
        console de log (QTextEdit somente leitura); e um rodapé com o
        botão "Voltar ao painel" (à esquerda) e o botão "Parar" (no
        canto inferior direito, ao lado do rótulo de resultado).

        Parâmetros:
            ano (str): ano da execução.
            mes (str): mês da execução, por extenso.
            tipo (str): tipo do pacote a ser baixado.
            novo (bool): se True (padrão), inicia uma execução nova do
                zero — cria e dispara um novo Worker, conectando seus
                sinais (log_signal, done_signal, error_signal,
                cancelado_signal) aos métodos que atualizam a
                interface. Se False, apenas reabre a tela de
                acompanhamento de uma execução que já existe
                (self.worker), sem criar um novo Worker — usado ao
                clicar em "Acompanhar execução" no painel, para voltar
                a uma execução que já estava rodando (ou que já
                terminou enquanto o usuário estava em outra tela). Em
                ambos os casos o console de log é (re)povoado a partir
                de self.log_execucao_atual, que é preservado mesmo
                quando o usuário navega para outra tela — é isso que
                permite ir ao painel durante o processamento sem que a
                interface trave ou feche ao voltar.
        """
        self.limpar_tela()
        self.sidebar.show()
        self._definir_usuario()

        pagina, layout = self._pagina()

        topo = QHBoxLayout()
        topo.addWidget(
            criar_titulo(
                "Execução em andamento",
                f"Processando {tipo} — {mes.capitalize()}/{ano}",
            )
        )
        topo.addStretch()
        self.exec_status = QLabel("●  PROCESSANDO")
        self.exec_status.setObjectName("statusBadgeRunning")
        topo.addWidget(self.exec_status, alignment=Qt.AlignmentFlag.AlignTop)
        layout.addLayout(topo)

        info = Card()
        ilayout = QHBoxLayout(info)
        ilayout.setContentsMargins(22, 17, 22, 17)

        self.exec_periodo = QLabel(f"{mes.capitalize()} / {ano}  •  {tipo}")
        self.exec_periodo.setObjectName("periodValue")
        ilayout.addWidget(self.exec_periodo)

        ilayout.addStretch()

        self.exec_hint = QLabel("Download direto e headless — sem navegador, sem CAPTCHA.")
        self.exec_hint.setObjectName("hint")
        ilayout.addWidget(self.exec_hint)

        layout.addWidget(info)

        log_card = Card()
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(18, 18, 18, 18)

        log_header = QHBoxLayout()
        label = QLabel("LOG DA EXECUÇÃO")
        label.setObjectName("eyebrow")
        log_header.addWidget(label)
        log_header.addStretch()

        self.log_counter = QLabel("0 eventos")
        self.log_counter.setObjectName("logCounter")
        log_header.addWidget(self.log_counter)
        log_layout.addLayout(log_header)

        self.text_log = QTextEdit()
        self.text_log.setReadOnly(True)
        self.text_log.setObjectName("logConsole")
        log_layout.addWidget(self.text_log)
        layout.addWidget(log_card, 1)

        bottom = QHBoxLayout()
        self.botao_voltar_execucao = Botao("← Voltar ao painel", compacto=True)
        self.botao_voltar_execucao.setEnabled(False)
        self.botao_voltar_execucao.clicked.connect(self.tela_principal)
        bottom.addWidget(self.botao_voltar_execucao)
        bottom.addStretch()

        self.arquivo_resultado = QLabel("")
        self.arquivo_resultado.setObjectName("resultPath")
        bottom.addWidget(self.arquivo_resultado)

        # Botão "Parar", no canto inferior direito do rodapé.
        self.botao_parar = Botao("■  Parar", compacto=True, perigo=True)
        self.botao_parar.clicked.connect(self.parar_execucao)
        bottom.addWidget(self.botao_parar)

        layout.addLayout(bottom)

        if novo:
            self.log_execucao_atual = []
            self.execucao_status_final = None

            self.worker = Worker(ano, mes, tipo)
            self.worker.log_signal.connect(self._log_gui)
            self.worker.done_signal.connect(self._execucao_concluida)
            self.worker.error_signal.connect(self._execucao_erro)
            self.worker.cancelado_signal.connect(self._execucao_cancelada)
            self.worker.start()

        # Repovoa o console de log com tudo o que já foi registrado até
        # agora (histórico acumulado desde o início desta execução),
        # inclusive se o usuário passou um tempo em outra tela.
        if self.log_execucao_atual:
            self.text_log.setPlainText("\n".join(self.log_execucao_atual))
        total = len(self.log_execucao_atual)
        self.log_counter.setText(f"{total} evento" + ("" if total == 1 else "s"))

        self.execucao_visivel = True
        self._atualizar_estado_tela_execucao()

    def _atualizar_estado_tela_execucao(self):
        """Ajusta o indicativo de status, o rótulo de resultado e os
        botões (Voltar/Parar) da tela de execução, de acordo com
        self.execucao_status_final: None enquanto o robô ainda está
        rodando, ou "sucesso"/"erro"/"cancelado" quando já terminou.

        Só produz efeito quando a tela de execução está de fato visível
        (self.execucao_visivel) — assim, se o Worker terminar (ou for
        cancelado) enquanto o usuário está no painel, esta função não
        tenta atualizar widgets que já não existem mais.
        """
        if not self.execucao_visivel:
            return

        estados_finais = {
            "sucesso": ("●  CONCLUÍDO", "statusBadge"),
            "erro": ("●  ERRO", "statusBadgeError"),
            "cancelado": ("●  CANCELADO", "statusBadgeError"),
        }

        if self.execucao_status_final in estados_finais:
            texto, nome = estados_finais[self.execucao_status_final]
            self.exec_status.setText(texto)
            self.exec_status.setObjectName(nome)
            self.botao_voltar_execucao.setEnabled(True)
            self.botao_parar.setEnabled(False)
            if self.execucao_status_final == "sucesso":
                self.arquivo_resultado.setText("RESULTADO GERADO")
        else:
            self.exec_status.setText("●  PROCESSANDO")
            self.exec_status.setObjectName("statusBadgeRunning")
            self.botao_voltar_execucao.setEnabled(False)
            self.botao_parar.setEnabled(bool(self.worker and self.worker.isRunning()))

        self.exec_status.style().unpolish(self.exec_status)
        self.exec_status.style().polish(self.exec_status)

    def parar_execucao(self):
        """Chamado ao clicar no botão "Parar" da tela de execução.

        Pede confirmação ao usuário e, se confirmado, sinaliza ao
        Worker (via solicitar_parada) que a execução deve ser
        interrompida assim que atingir o próximo ponto seguro dentro do
        pipeline. A parada é cooperativa, não instantânea: pode levar
        alguns segundos até o cancelamento de fato se efetivar (o robô
        nunca é interrompido no meio de uma operação sensível, como a
        escrita de um arquivo).
        """
        if not (self.worker and self.worker.isRunning()):
            return

        resposta = QMessageBox.question(
            self,
            "Parar execução",
            "Tem certeza que deseja interromper a execução em andamento?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if resposta != QMessageBox.StandardButton.Yes:
            return

        self.worker.solicitar_parada()
        if self.execucao_visivel:
            self.botao_parar.setEnabled(False)
        self._log_gui("Solicitação de parada enviada — aguardando um ponto seguro para interromper...")

    def _log_gui(self, texto):
        """Recebe uma nova linha de log (emitida pelo Worker via
        log_signal, ou registrada internamente, ex: ao solicitar
        parada) e a acrescenta ao histórico desta execução
        (self.log_execucao_atual), que é preservado mesmo que o usuário
        esteja em outra tela.

        Se a tela de execução estiver visível no momento, também
        atualiza o console de log e o contador de eventos em tela; caso
        contrário, a atualização visual fica só registrada no
        histórico — ao voltar para a tela de execução, o console é
        repovoado com tudo o que aconteceu nesse meio-tempo.

        Parâmetros:
            texto (str): linha de log a ser exibida.
        """
        self.log_execucao_atual.append(texto)

        if not self.execucao_visivel:
            return

        self.text_log.append(texto)
        total = len(self.log_execucao_atual)
        self.log_counter.setText(f"{total} evento" + ("" if total == 1 else "s"))

    def _execucao_concluida(self, caminho, caminho_log):
        """Manipulador chamado quando o Worker emite done_signal (a
        execução terminou com sucesso).

        Registra linhas finais no log (caminhos da planilha e do log
        gerados), marca o estado final da execução como "sucesso" e
        atualiza a tela de execução (se estiver visível). Registra a
        execução no histórico (registrar_execucao) e, conforme as
        preferências do usuário, abre a pasta de saída automaticamente
        e/ou exibe uma notificação de conclusão — isso acontece mesmo
        que o usuário esteja no painel em vez da tela de execução.

        Parâmetros:
            caminho (str): caminho da planilha Excel gerada.
            caminho_log (str): caminho do arquivo de log da execução.
        """
        self.execucao_status_final = "sucesso"

        self._log_gui("")
        self._log_gui("✓ Processo finalizado com sucesso.")
        self._log_gui(f"Planilha: {caminho}")
        self._log_gui(f"Log: {caminho_log}")

        self._atualizar_estado_tela_execucao()

        registrar_execucao(self.worker.ano, self.worker.mes, "sucesso", caminho)

        if self.config.get("abrir_pasta_automatico", False):
            self.abrir_pasta_saida()

        if self.config.get("notificar_conclusao", True):
            self._mensagem(
                "Execução concluída",
                f"Planilha gerada com sucesso.\n\n{caminho}\n\nLog:\n{caminho_log}",
                "success",
            )

    def _execucao_erro(self, erro):
        """Manipulador chamado quando o Worker emite error_signal (a
        execução falhou com uma exceção).

        Registra a mensagem de erro no log, marca o estado final da
        execução como "erro" e atualiza a tela de execução (se estiver
        visível). Registra a falha no histórico (registrar_execucao com
        status "erro") e, se a preferência de notificação estiver
        ativa, exibe uma mensagem de erro ao usuário — mesmo que ele
        esteja no painel em vez da tela de execução.

        Parâmetros:
            erro (str): mensagem de erro (texto da exceção capturada).
        """
        self.execucao_status_final = "erro"

        self._log_gui("")
        self._log_gui(f"✕ Erro: {erro}")

        self._atualizar_estado_tela_execucao()

        registrar_execucao(self.worker.ano, self.worker.mes, "erro")

        if self.config.get("notificar_conclusao", True):
            self._mensagem("Falha na execução", str(erro), "error")

    def _execucao_cancelada(self):
        """Manipulador chamado quando o Worker emite cancelado_signal
        (o usuário pediu para parar, através do botão "Parar", e o
        pipeline foi interrompido no próximo ponto seguro).

        Registra a interrupção no log, marca o estado final da execução
        como "cancelado" e atualiza a tela de execução (se estiver
        visível). Registra a interrupção no histórico
        (registrar_execucao com status "cancelado") e, se a preferência
        de notificação estiver ativa, avisa o usuário.
        """
        self.execucao_status_final = "cancelado"

        self._log_gui("")
        self._log_gui("■ Execução interrompida pelo usuário.")

        self._atualizar_estado_tela_execucao()

        registrar_execucao(self.worker.ano, self.worker.mes, "cancelado")

        if self.config.get("notificar_conclusao", True):
            self._mensagem(
                "Execução interrompida",
                "A execução foi interrompida a seu pedido.",
                "info",
            )

    def tela_execucao_placeholder(self):
        """Método auxiliar/legado: se já houver uma execução em
        andamento, leva o usuário de volta ao acompanhamento dela (em
        vez de simplesmente não fazer nada, o que o deixaria sem
        feedback); caso contrário, encaminha para a tela de
        configuração de nova execução (tela_config_robo).
        """
        if self.worker and self.worker.isRunning():
            self.tela_execucao(self.worker.ano, self.worker.mes, self.worker.tipo, novo=False)
            return
        self.tela_config_robo()

    def closeEvent(self, event):
        """Sobrescreve o evento de fechamento da janela principal.

        Se houver uma execução do robô em andamento, pergunta ao usuário
        se ele realmente deseja fechar a interface mesmo assim; se a
        resposta for "Não", cancela o fechamento (event.ignore()). Caso
        não haja execução em andamento, ou se o usuário confirmar,
        permite o fechamento normal (event.accept()).

        Parâmetros:
            event (QCloseEvent): evento de fechamento fornecido pelo Qt.
        """
        if self.worker and self.worker.isRunning():
            resposta = QMessageBox.question(
                self,
                "Execução em andamento",
                "O robô ainda está executando. Deseja fechar a interface?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if resposta == QMessageBox.StandardButton.No:
                event.ignore()
                return
        event.accept()


def aplicar_tema(app):
    """Aplica a folha de estilos (QSS) global da aplicação, definindo a
    aparência visual (cores, bordas, tipografia) de todos os
    componentes: janela principal, sidebar, títulos, cartões (normais e
    "escuros"), lista de arquivos, badges de status, botões, campos de
    texto/combo, telas de login/cadastro, console de log, barra de
    rolagem e diálogos/QMessageBox.

    Parâmetros:
        app (QApplication): instância da aplicação Qt à qual o estilo
            será aplicado (via setStyleSheet).
    """
    app.setStyleSheet(
        f"""
        * {{
            font-family: "Segoe UI", "Arial";
        }}

        /* =========================================================
           BASE — claro / bege / futurista
           ========================================================= */
        QMainWindow {{
            background: #F3EFE5;
        }}

        QWidget {{
            background: transparent;
            color: {TEXT};
        }}

        QFrame#contentArea {{
            background: #F3EFE5;
        }}

        QScrollArea#mainScrollArea, QScrollArea#mainScrollArea > QWidget > QWidget {{
            background: transparent;
            border: none;
        }}

        /* =========================================================
           SIDEBAR — painel lateral minimalista
           ========================================================= */
        QFrame#sidebar {{
            background: #FBF9F4;
            border-right: 1px solid #DDD5C7;
        }}

        QLabel#brandName {{
            color: #171612;
            font-size: 14px;
            font-weight: 900;
            letter-spacing: 1.1px;
        }}

        QLabel#brandSub {{
            color: #B17E18;
            font-size: 8px;
            font-weight: 900;
            letter-spacing: 1.5px;
        }}

        QFrame#separator {{
            background: #E2DACB;
            border: none;
            max-height: 1px;
        }}

        QLabel#sideSection,
        QLabel#sideCaption {{
            color: #9B927F;
            font-size: 9px;
            font-weight: 900;
            letter-spacing: 1.6px;
        }}

        QLabel#sideUser {{
            color: #25221C;
            font-size: 13px;
            font-weight: 900;
        }}

        QFrame#userCard {{
            background: #FFFFFF;
            border: 1px solid #DED6C8;
            border-radius: 12px;
        }}

        QPushButton#navButton {{
            background: transparent;
            color: #777062;
            border: 1px solid transparent;
            border-radius: 10px;
            text-align: left;
            padding: 0 13px;
            font-size: 12px;
            font-weight: 800;
        }}

        QPushButton#navButton:hover {{
            background: #F1EBDD;
            border-color: #E5DCCB;
            color: #27231C;
        }}

        QPushButton#navButton:pressed {{
            background: #E9D5A1;
            color: #5B4215;
        }}

        /* =========================================================
           TÍTULOS
           ========================================================= */
        QLabel#pageTitle {{
            color: #191814;
            font-size: 24px;
            font-weight: 900;
        }}

        QLabel#pageSubtitle {{
            color: #746D61;
            font-size: 11px;
        }}

        /* =========================================================
           CARDS — flat, sem sombra, bordas técnicas
           ========================================================= */
        QFrame#card {{
            background: #FFFEFB;
            border: 1px solid #DED6C8;
            border-radius: 14px;
        }}

        QLabel#cardNumber {{
            color: #B17E18;
            background: transparent;
            font-size: 11px;
            font-weight: 900;
            letter-spacing: 2px;
        }}

        QLabel#cardTitle {{
            color: #1E1B16;
            background: transparent;
            font-size: 15px;
            font-weight: 900;
        }}

        QLabel#cardText {{
            color: #777062;
            background: transparent;
            font-size: 11px;
            line-height: 1.4;
        }}

        QLabel#eyebrow {{
            color: #B17E18;
            background: transparent;
            font-size: 9px;
            font-weight: 900;
            letter-spacing: 2px;
        }}

        QLabel#fieldLabel {{
            color: #71695C;
            background: transparent;
            font-size: 9px;
            font-weight: 900;
            letter-spacing: 1.5px;
        }}

        QLabel#hint {{
            color: #7D7568;
            background: transparent;
            font-size: 10px;
            padding: 8px 0;
        }}

        /* =========================================================
           CARTÃO ESCURO — status do robô, ao estilo da referência
           ========================================================= */
        QFrame#darkCard {{
            background: #23201A;
            border: 1px solid #23201A;
            border-radius: 14px;
        }}

        QLabel#darkEyebrow {{
            color: {YELLOW_LIGHT};
            background: transparent;
            font-size: 9px;
            font-weight: 900;
            letter-spacing: 2px;
        }}

        QLabel#darkValue {{
            color: #FFFFFF;
            background: transparent;
            font-size: 21px;
            font-weight: 900;
        }}

        QLabel#darkCaption {{
            color: #C9C2B4;
            background: transparent;
            font-size: 11px;
        }}

        QLabel#darkFoot {{
            color: #8B8478;
            background: transparent;
            font-size: 10px;
            font-weight: 700;
        }}

        /* =========================================================
           LISTA DE ARQUIVOS — linhas clicáveis
           ========================================================= */
        QFrame#fileRow {{
            background: #FBF8F1;
            border: 1px solid #E7E0D2;
            border-radius: 10px;
        }}

        QFrame#fileRow:hover {{
            background: #F6EFDD;
            border-color: {YELLOW};
        }}

        QLabel#fileName {{
            color: #25221D;
            background: transparent;
            font-size: 12px;
            font-weight: 700;
        }}

        QLabel#fileMeta {{
            color: #8A8274;
            background: transparent;
            font-size: 10px;
            font-weight: 700;
        }}

        QPushButton#toggleSwitch {{
            border: none;
            background: transparent;
        }}

        /* =========================================================
           STATUS
           ========================================================= */
        QLabel#statusBadge,
        QLabel#statusBadgeRunning,
        QLabel#statusBadgeError {{
            background: #FFF8E5;
            color: #916A18;
            border: 1px solid #DDBA62;
            border-radius: 9px;
            padding: 6px 11px;
            font-size: 9px;
            font-weight: 900;
            letter-spacing: 0.8px;
        }}

        QLabel#statusBadgeRunning {{
            background: #FFF4D1;
            color: #8B6415;
        }}

        QLabel#statusBadgeError {{
            color: #B94D4D;
            border-color: #DFAAAA;
            background: #FFF1F1;
        }}

        /* =========================================================
           BOTÕES
           ========================================================= */
        QPushButton[principal="true"] {{
            background: #D5A63A;
            color: #201A0F;
            border: 1px solid #C2942D;
            border-radius: 10px;
            padding: 0 16px;
            font-size: 11px;
            font-weight: 900;
            letter-spacing: 0.8px;
        }}

        QPushButton[principal="true"]:hover {{
            background: #E2B74D;
            border-color: #B98720;
        }}

        QPushButton[principal="true"]:pressed {{
            background: #BD8C27;
        }}

        QPushButton[principal="true"]:disabled {{
            background: #D9CCAA;
            color: #8E8368;
            border-color: #D0C3A3;
        }}

        QPushButton[compacto="true"] {{
            background: #FFFEFB;
            color: #4D473D;
            border: 1px solid #D7CEBE;
            border-radius: 9px;
            padding: 0 12px;
            font-size: 10px;
            font-weight: 900;
        }}

        QPushButton[compacto="true"]:hover {{
            background: #F8F0DF;
            color: #6C4F16;
            border-color: #CFA544;
        }}

        QPushButton[compacto="true"]:disabled {{
            color: #B5AD9E;
            border-color: #E1DBCF;
        }}

        QPushButton[compacto="true"][grande="true"] {{
            font-size: 13px;
            padding: 0 20px;
        }}

        QPushButton[perigo="true"] {{
            background: #FBEDEC;
            color: #A03D3D;
            border: 1px solid #E2B3B0;
            border-radius: 9px;
            padding: 0 14px;
            font-size: 10px;
            font-weight: 900;
        }}

        QPushButton[perigo="true"]:hover {{
            background: #F6DBD9;
            color: #832E2E;
            border-color: #D89C98;
        }}

        QPushButton[perigo="true"]:pressed {{
            background: #EFC7C4;
        }}

        QPushButton[perigo="true"]:disabled {{
            background: #F8F3E8;
            color: #B5AD9E;
            border-color: #E1DBCF;
        }}

        QPushButton#linkButton {{
            background: transparent;
            border: none;
            color: #A47519;
            font-size: 10px;
            font-weight: 800;
            padding: 7px;
        }}

        QPushButton#linkButton:hover {{
            color: #765310;
        }}

        /* =========================================================
           CAMPOS
           ========================================================= */
        QLineEdit {{
            background: #FFFFFF;
            color: #25221D;
            border: 1px solid #D9D1C3;
            border-radius: 10px;
            padding: 0 14px;
            font-size: 13px;
        }}

        QLineEdit:hover,
        QLineEdit:focus {{
            background: #FFFEFB;
            border: 1px solid #C99A35;
        }}

        QComboBox {{
            background: #FFFFFF;
            color: #25221D;
            border: 1px solid #D9D1C3;
            border-radius: 10px;
            padding: 0 14px;
            font-size: 13px;
        }}

        QComboBox:hover,
        QComboBox:focus {{
            background: #FFFEFB;
            border: 1px solid #C99A35;
        }}

        QComboBox::drop-down {{
            border: none;
            width: 34px;
        }}

        QComboBox QAbstractItemView {{
            background: #FFFFFF;
            color: #25221D;
            border: 1px solid #D9D1C3;
            selection-background-color: #D5A63A;
            selection-color: #201A0F;
            padding: 5px;
        }}

        /* =========================================================
           LOGIN / CADASTRO
           ========================================================= */
        QLabel#authTitle {{
            color: #191814;
            background: transparent;
            font-size: 22px;
            font-weight: 900;
        }}

        QLabel#authSubtitle {{
            color: #756D61;
            background: transparent;
            font-size: 11px;
        }}

        /* =========================================================
           EXECUÇÃO / LOG
           ========================================================= */
        QTextEdit#logConsole {{
            background: #FAF8F2;
            color: #49443B;
            border: 1px solid #D9D1C3;
            border-radius: 10px;
            padding: 12px;
            font-family: "Cascadia Mono", "Consolas", monospace;
            font-size: 11px;
            selection-background-color: #D5A63A;
            selection-color: #201A0F;
        }}

        QLabel#logCounter {{
            color: #81786B;
            background: transparent;
            font-size: 9px;
        }}

        QLabel#periodValue {{
            color: #9A6D15;
            background: transparent;
            font-size: 16px;
            font-weight: 900;
        }}

        QLabel#resultPath {{
            color: #4B916C;
            background: transparent;
            font-size: 10px;
            font-weight: 900;
        }}

        QLabel#storagePath {{
            color: #777063;
            background: transparent;
            font-size: 9px;
        }}

        QLabel#flowNumber {{
            color: #B17E18;
            background: transparent;
            font-size: 10px;
            font-weight: 900;
        }}

        QLabel#flowText {{
            color: #332F28;
            background: transparent;
            font-size: 10px;
        }}

        /* =========================================================
           SCROLLBAR
           ========================================================= */
        QScrollBar:vertical {{
            background: #EEE8DC;
            width: 7px;
            margin: 0;
            border: none;
        }}

        QScrollBar::handle:vertical {{
            background: #CDBFA8;
            border-radius: 3px;
            min-height: 28px;
        }}

        QScrollBar::handle:vertical:hover {{
            background: #D5A63A;
        }}

        /* =========================================================
           DIÁLOGOS
           ========================================================= */
        QDialog {{
            background: #FBF9F4;
        }}

        QLabel#dialogTitle {{
            color: #191814;
            background: transparent;
        }}

        QLabel#dialogText {{
            color: #756D61;
            background: transparent;
        }}

        QPushButton#dialogPrimary {{
            background: #D5A63A;
            color: #201A0F;
            border: 1px solid #C2942D;
            border-radius: 9px;
            padding: 8px 20px;
            min-height: 42px;
            font-weight: 900;
        }}

        QPushButton#dialogSecundario {{
            background: #FFFEFB;
            color: #4D473D;
            border: 1px solid #D7CEBE;
            border-radius: 9px;
            padding: 8px 20px;
            min-height: 42px;
            font-weight: 800;
        }}

        QPushButton#dialogSecundario:hover {{
            background: #F8F0DF;
            color: #6C4F16;
            border-color: #CFA544;
        }}

        QPushButton#dialogSecundario:disabled {{
            color: #B5AD9E;
            border-color: #E1DBCF;
        }}

        QMessageBox {{
            background: #FFFEFB;
        }}

        QMessageBox QLabel {{
            color: #25221D;
            background: transparent;
            font-size: 13px;
        }}

        QMessageBox QPushButton {{
            background: #D5A63A;
            color: #201A0F;
            border: 1px solid #C2942D;
            border-radius: 8px;
            padding: 8px 18px;
            min-width: 80px;
            font-weight: 900;
        }}
        """
    )


def main():
    """Ponto de entrada da aplicação: cria a QApplication, define nome
    do app/organização, aplica o tema visual global (aplicar_tema),
    instancia e exibe a janela principal (App) e inicia o loop de
    eventos do Qt (app.exec()).
    """
    app = QApplication([])
    app.setApplicationName("Robô SIAPE")
    app.setOrganizationName("Robô SIAPE")
    # Força o estilo "Fusion": sem isso, o Windows usa o estilo nativo
    # (windowsvista) para desenhar os botões, que só respeita parte da
    # folha de estilos (QSS) quando a janela está maximizada — fora da
    # tela cheia, botões e outros widgets ficam com aparência quebrada/
    # inconsistente. O Fusion garante que a QSS seja sempre respeitada,
    # em qualquer estado da janela (maximizada ou em modo janela).
    app.setStyle("Fusion")
    aplicar_tema(app)

    janela = App()
    janela.show()

    app.exec()


if __name__ == "__main__":
    main()