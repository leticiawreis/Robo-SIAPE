import hashlib
import json
import logging
import os
import queue
import re
import threading
from collections import OrderedDict
from datetime import date, datetime

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QRectF, QUrl
from PyQt6.QtGui import QDesktopServices, QFont, QColor, QPainter, QPen
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

from robo_siape import INDICES_ANOS, meses_disponiveis, executar_pipeline_completo, PASTA_SAIDA

TIPO_PACOTE = "Servidores_SIAPE"


ARQUIVO_USUARIOS = "usuarios.json"
ARQUIVO_HISTORICO = "historico_execucoes.json"
ARQUIVO_CONFIG = "config.json"

def preparar_pastas():
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
    if not os.path.exists(ARQUIVO_USUARIOS):
        return {}
    with open(ARQUIVO_USUARIOS, "r", encoding="utf-8") as arquivo:
        return json.load(arquivo)


def salvar_usuarios(usuarios):
    with open(ARQUIVO_USUARIOS, "w", encoding="utf-8") as arquivo:
        json.dump(usuarios, arquivo, indent=2, ensure_ascii=False)


def hash_senha(senha):

    return hashlib.sha256(senha.encode("utf-8")).hexdigest()


def usuario_valido(usuario):
    return re.fullmatch(r"[A-Za-z0-9_.-]{3,30}", usuario) is not None


# =====================================================================
# HISTÓRICO DE EXECUÇÕES — alimenta o gráfico e os indicadores do painel
# =====================================================================
def carregar_historico():
    if not os.path.exists(ARQUIVO_HISTORICO):
        return []
    try:
        with open(ARQUIVO_HISTORICO, "r", encoding="utf-8") as arquivo:
            return json.load(arquivo)
    except (json.JSONDecodeError, OSError):
        return []


def salvar_historico(historico):
    with open(ARQUIVO_HISTORICO, "w", encoding="utf-8") as arquivo:
        json.dump(historico[-200:], arquivo, indent=2, ensure_ascii=False)


def registrar_execucao(ano, mes, status, arquivo_gerado=None):
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
    historico = carregar_historico()
    total = len(historico)
    sucesso = sum(1 for item in historico if item.get("status") == "sucesso")
    taxa = round((sucesso / total) * 100) if total else 0
    return total, sucesso, taxa


def contagem_mensal(qtd_meses=6):
    historico = carregar_historico()
    contagem = OrderedDict()
    for item in historico:
        mes = str(item.get("mes", "")).strip()
        ano = str(item.get("ano", "")).strip()
        chave = f"{mes[:3].capitalize()}/{ano[-2:]}" if mes and ano else "—"
        contagem[chave] = contagem.get(chave, 0) + 1
    chaves = list(contagem.keys())[-qtd_meses:]
    valores = [contagem[c] for c in chaves]
    while len(valores) < qtd_meses:
        valores.insert(0, 0)
        chaves.insert(0, "")
    return chaves, valores


def ultima_execucao():
    historico = carregar_historico()
    if not historico:
        return None
    return historico[-1]


# =====================================================================
# CONFIGURAÇÕES DO USUÁRIO — preferências reais aplicadas no painel
# =====================================================================
def carregar_config():
    padrao = {"notificar_conclusao": True, "abrir_pasta_automatico": False}
    if os.path.exists(ARQUIVO_CONFIG):
        try:
            with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as arquivo:
                padrao.update(json.load(arquivo))
        except (json.JSONDecodeError, OSError):
            pass
    return padrao


def salvar_config(config):
    with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as arquivo:
        json.dump(config, arquivo, indent=2, ensure_ascii=False)


# =====================================================================
# ARQUIVOS DE SAÍDA — lista real do que o robô já gerou em disco
# =====================================================================
def formatar_tamanho(num_bytes):
    tamanho = float(num_bytes)
    for unidade in ("B", "KB", "MB", "GB"):
        if tamanho < 1024 or unidade == "GB":
            return f"{tamanho:.0f} {unidade}" if unidade == "B" else f"{tamanho:.1f} {unidade}"
        tamanho /= 1024
    return f"{tamanho:.1f} GB"


def listar_arquivos_saida(limite=6):
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



def estilizar_combo(combo):
    combo.setMinimumHeight(48)
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
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(54, 54)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setBrush(QColor(YELLOW))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(3, 3, 48, 48, 14, 14)
        painter.setPen(QPen(QColor("#0D0D0D"), 2))
        painter.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "RS")


class Campo(QLineEdit):
    def __init__(self, placeholder="", senha=False):
        super().__init__()
        self.setPlaceholderText(placeholder)
        self.setMinimumHeight(48)
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
    def __init__(self, texto, principal=False, compacto=False):
        super().__init__(texto)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(44 if compacto else 50)
        self.setProperty("principal", principal)
        self.setProperty("compacto", compacto)


class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")


class GaugeWidget(QWidget):
    """Indicador circular com um valor real (contagem ou percentual)."""

    def __init__(self, valor=0, maximo=100, sufixo="", rotulo="", parent=None):
        super().__init__(parent)
        self.valor = valor
        self.maximo = max(maximo, 1)
        self.sufixo = sufixo
        self.rotulo = rotulo
        self.setMinimumSize(120, 130)

    def definir_valor(self, valor, maximo=None, rotulo=None):
        self.valor = valor
        if maximo is not None:
            self.maximo = max(maximo, 1)
        if rotulo is not None:
            self.rotulo = rotulo
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        lado = min(self.width(), self.height() - 26) - 16
        x = (self.width() - lado) / 2
        y = 8
        area = QRectF(x, y, lado, lado)

        caneta_fundo = QPen(QColor(BORDER))
        caneta_fundo.setWidth(10)
        caneta_fundo.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(caneta_fundo)
        painter.drawArc(area, 225 * 16, -270 * 16)

        proporcao = min(self.valor / self.maximo, 1.0) if self.maximo else 0
        caneta_valor = QPen(QColor(YELLOW))
        caneta_valor.setWidth(10)
        caneta_valor.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(caneta_valor)
        painter.drawArc(area, 225 * 16, int(-270 * proporcao * 16))

        painter.setPen(QColor(TEXT))
        painter.setFont(QFont("Segoe UI", 20, QFont.Weight.Black))
        texto_valor = f"{self.valor}{self.sufixo}"
        painter.drawText(area, Qt.AlignmentFlag.AlignCenter, texto_valor)

        rotulo_rect = QRectF(0, y + lado - 6, self.width(), 26)
        painter.setPen(QColor(MUTED))
        painter.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
        painter.drawText(rotulo_rect, Qt.AlignmentFlag.AlignHCenter, self.rotulo.upper())


class BarChartWidget(QWidget):
    """Mini gráfico de barras com execuções reais por mês."""

    def __init__(self, valores=None, rotulos=None, parent=None):
        super().__init__(parent)
        self.valores = valores or []
        self.rotulos = rotulos or []
        self.setMinimumHeight(120)

    def definir_dados(self, valores, rotulos):
        self.valores = valores
        self.rotulos = rotulos
        self.update()

    def paintEvent(self, event):
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
    """Interruptor liga/desliga real, usado para preferências do usuário."""

    def __init__(self, marcado=False, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setChecked(marcado)
        self.setFixedSize(44, 24)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setObjectName("toggleSwitch")

    def paintEvent(self, event):
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
    log_signal = pyqtSignal(str)
    done_signal = pyqtSignal(str, str)
    error_signal = pyqtSignal(str)

    def __init__(self, ano, mes, tipo):
        super().__init__()
        self.ano = ano
        self.mes = mes
        self.tipo = tipo
        self.log_queue = queue.Queue()

    def run(self):
        thread_fila = threading.Thread(target=self._processar_fila, daemon=True)
        thread_fila.start()

        try:
            caminho, caminho_log = executar_pipeline_completo(
                self.ano,
                self.mes,
                self.tipo,
                self.log_queue,
            )
            self.log_queue.put(("done", (caminho, caminho_log)))
        except Exception as erro:
            logging.getLogger("robo_siape").exception("Falha na execução.")
            self.log_queue.put(("error", str(erro)))

        thread_fila.join(timeout=2)

    def _processar_fila(self):
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


class App(QMainWindow):
    def __init__(self):
        super().__init__()

        self.usuario_logado = None
        self.worker = None
        self.config = carregar_config()
        self.tela = QStackedWidget()

        self.setWindowTitle("Robô SIAPE")
        preparar_pastas()
        self.setMinimumSize(1120, 720)
        self.resize(1180, 760)

        self._montar_interface()
        self.tela_login()

    def _montar_interface(self):
        central = QWidget()
        central_layout = QHBoxLayout(central)
        central_layout.setContentsMargins(0, 0, 0, 0)
        central_layout.setSpacing(0)

        self.sidebar = self._criar_sidebar()
        central_layout.addWidget(self.sidebar)

        self.conteudo = QFrame()
        self.conteudo.setObjectName("contentArea")
        content_layout = QVBoxLayout(self.conteudo)
        content_layout.setContentsMargins(42, 34, 42, 32)
        content_layout.addWidget(self.tela)

        central_layout.addWidget(self.conteudo, 1)
        self.setCentralWidget(central)

    def _criar_sidebar(self):
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(238)

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
        self.nav_robot = self._nav_button("◈", "Executar robô")

        self.nav_home.clicked.connect(self.tela_principal)
        self.nav_robot.clicked.connect(self.tela_config_robo)

        layout.addWidget(self.nav_home)
        layout.addWidget(self.nav_robot)

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
        botao = QPushButton()
        botao.setText(f"{icon}   {texto}")
        botao.setCursor(Qt.CursorShape.PointingHandCursor)
        botao.setMinimumHeight(46)
        botao.setObjectName("navButton")
        return botao

    def _definir_usuario(self):
        self.side_user.setText(self.usuario_logado or "—")

    def limpar_tela(self):
        while self.tela.count():
            widget = self.tela.widget(0)
            self.tela.removeWidget(widget)
            widget.deleteLater()

    def _pagina(self):
        pagina = QWidget()
        layout = QVBoxLayout(pagina)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(20)
        self.tela.addWidget(pagina)
        return pagina, layout

    def _label(self, texto):
        label = QLabel(texto)
        label.setObjectName("fieldLabel")
        return label

    def _mensagem(self, titulo, texto, tipo="info"):
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
        self.limpar_tela()
        self.sidebar.show()
        self._definir_usuario()
        self.config = carregar_config()

        em_execucao = bool(self.worker and self.worker.isRunning())

        pagina, layout = self._pagina()
        layout.setSpacing(16)

        # ---------------------------------------------------------------
        # Cabeçalho
        # ---------------------------------------------------------------
        topo = QHBoxLayout()
        topo.addWidget(
            criar_titulo(
                "Painel de controle",
                "Automação de dados de remuneração do Portal da Transparência.",
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
        linha1.setSpacing(14)
        linha1.addWidget(self._card_status_robo(em_execucao), 3)
        linha1.addWidget(self._card_atividade(), 3)
        linha1.addWidget(self._card_acoes_rapidas(), 3)
        layout.addLayout(linha1)

        # ---------------------------------------------------------------
        # Linha 2 — barra de atalhos + preferências (toggles reais)
        # ---------------------------------------------------------------
        linha2 = QHBoxLayout()
        linha2.setSpacing(14)
        linha2.addWidget(self._barra_atalhos(), 3)
        linha2.addWidget(self._painel_preferencias(), 1)
        layout.addLayout(linha2)

        # ---------------------------------------------------------------
        # Linha 3 — arquivos gerados + indicadores (gauges reais)
        # ---------------------------------------------------------------
        linha3 = QHBoxLayout()
        linha3.setSpacing(14)
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
            "O robô baixa o pacote escolhido diretamente do Portal da Transparência "
            "(modo headless, sem abrir navegador), remove dados desnecessários "
            "e gera a planilha tratada automaticamente."
        )
        desc.setObjectName("cardText")
        desc.setWordWrap(True)
        left.addWidget(desc)
        dlayout.addLayout(left, 1)

        iniciar = Botao(
            "ACOMPANHAR EXECUÇÃO  →" if em_execucao else "CONFIGURAR ROBÔ  →",
            principal=True,
        )
        iniciar.setMinimumWidth(230)
        iniciar.clicked.connect(
            (lambda: self.tela_execucao(self.worker.ano, self.worker.mes, self.worker.tipo))
            if em_execucao
            else self.tela_config_robo
        )
        dlayout.addWidget(iniciar)

        layout.addWidget(destaque)

    # ---------------------------------------------------------------
    # Cartões do painel
    # ---------------------------------------------------------------
    def _card_status_robo(self, em_execucao):
        card = Card()
        card.setObjectName("darkCard")
        card.setMinimumHeight(190)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(10)

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
        card = Card()
        card.setMinimumHeight(190)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(6)

        tag = QLabel("ATIVIDADE")
        tag.setObjectName("eyebrow")
        lay.addWidget(tag)

        total, sucesso, taxa = estatisticas_execucoes()
        titulo = QLabel(f"{total} execuções registradas")
        titulo.setObjectName("cardTitle")
        lay.addWidget(titulo)

        rotulos, valores = contagem_mensal()
        self.grafico_atividade = BarChartWidget(valores, rotulos)
        lay.addWidget(self.grafico_atividade, 1)

        legenda = QLabel("Execuções por mês (últimos 6 meses)")
        legenda.setObjectName("cardText")
        lay.addWidget(legenda)
        return card

    def _card_acoes_rapidas(self):
        card = Card()
        card.setMinimumHeight(190)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(10)

        tag = QLabel("AÇÕES RÁPIDAS")
        tag.setObjectName("eyebrow")
        lay.addWidget(tag)

        nova = Botao("+  Nova execução", principal=True, compacto=True)
        nova.clicked.connect(self.tela_config_robo)
        lay.addWidget(nova)

        abrir = Botao("📂  Abrir pasta de saída", compacto=True)
        abrir.clicked.connect(self.abrir_pasta_saida)
        lay.addWidget(abrir)

        atualizar = Botao("⟳  Atualizar painel", compacto=True)
        atualizar.clicked.connect(self.tela_principal)
        lay.addWidget(atualizar)

        ver_log = Botao("📜  Ver último log", compacto=True)
        ver_log.clicked.connect(self.abrir_ultimo_log)
        lay.addWidget(ver_log)

        lay.addStretch()
        return card

    def abrir_ultimo_log(self):
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

    def _barra_atalhos(self):
        card = Card()
        lay = QHBoxLayout(card)
        lay.setContentsMargins(18, 12, 18, 12)
        lay.setSpacing(10)

        atalhos = [
            ("📄  Resumo", self.tela_principal),
            ("▶  Nova execução", self.tela_config_robo),
            ("📂  Pasta de saída", self.abrir_pasta_saida),
            ("👤  Perfil", lambda: self._mensagem(
                "Usuário", f"Conectado como: {self.usuario_logado or '—'}", "info"
            )),
            ("⟳  Atualizar", self.tela_principal),
        ]
        for texto, acao in atalhos:
            botao = Botao(texto, compacto=True)
            botao.clicked.connect(acao)
            lay.addWidget(botao)

        lay.addStretch()
        return card

    def _painel_preferencias(self):
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
        self.config[chave] = valor
        salvar_config(self.config)

    def _card_arquivos(self):
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
            QDesktopServices.openUrl(QUrl.fromLocalFile(caminho))

        linha.mousePressEvent = abrir
        return linha

    def _card_gauges(self):
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
        preparar_pastas()
        QDesktopServices.openUrl(QUrl.fromLocalFile(PASTA_SAIDA))

    def tela_config_robo(self):
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
        ano = self.combo_ano.currentText()
        self.combo_mes.clear()
        self.combo_mes.addItems(meses_disponiveis(ano))

    def validar_periodo(self, ano, mes, tipo):
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

        self.tela_execucao(ano, mes, tipo)

    def tela_execucao(self, ano, mes, tipo):
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

        layout.addLayout(bottom)

        self.log_queue = queue.Queue()
        self.log_count = 0

        self.worker = Worker(ano, mes, tipo)
        self.worker.log_signal.connect(self._log_gui)
        self.worker.done_signal.connect(self._execucao_concluida)
        self.worker.error_signal.connect(self._execucao_erro)
        self.worker.start()

    def _log_gui(self, texto):
        self.log_count += 1
        self.text_log.append(texto)
        self.log_counter.setText(
            f"{self.log_count} evento" + ("" if self.log_count == 1 else "s")
        )

    def _execucao_concluida(self, caminho, caminho_log):
        self.exec_status.setText("●  CONCLUÍDO")
        self.exec_status.setObjectName("statusBadge")
        self.exec_status.style().unpolish(self.exec_status)
        self.exec_status.style().polish(self.exec_status)

        self._log_gui("")
        self._log_gui("✓ Processo finalizado com sucesso.")
        self._log_gui(f"Planilha: {caminho}")
        self._log_gui(f"Log: {caminho_log}")

        self.arquivo_resultado.setText("RESULTADO GERADO")
        self.botao_voltar_execucao.setEnabled(True)

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
        self.exec_status.setText("●  ERRO")
        self.exec_status.setObjectName("statusBadgeError")
        self.exec_status.style().unpolish(self.exec_status)
        self.exec_status.style().polish(self.exec_status)

        self._log_gui("")
        self._log_gui(f"✕ Erro: {erro}")
        self.botao_voltar_execucao.setEnabled(True)

        registrar_execucao(self.worker.ano, self.worker.mes, "erro")

        if self.config.get("notificar_conclusao", True):
            self._mensagem("Falha na execução", str(erro), "error")

    def tela_execucao_placeholder(self):
        if self.worker and self.worker.isRunning():
            return
        self.tela_config_robo()

    def closeEvent(self, event):
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

        /* =========================================================
           SIDEBAR — painel lateral minimalista
           ========================================================= */
        QFrame#sidebar {{
            background: #FBF9F4;
            border-right: 1px solid #DDD5C7;
        }}

        QLabel#brandName {{
            color: #171612;
            font-size: 16px;
            font-weight: 900;
            letter-spacing: 1.2px;
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
            font-size: 31px;
            font-weight: 900;
        }}

        QLabel#pageSubtitle {{
            color: #746D61;
            font-size: 12px;
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
            font-size: 17px;
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
            font-size: 26px;
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
            padding: 0 20px;
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
            padding: 0 17px;
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
            font-size: 28px;
            font-weight: 900;
        }}

        QLabel#authSubtitle {{
            color: #756D61;
            background: transparent;
            font-size: 12px;
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
    app = QApplication([])
    app.setApplicationName("Robô SIAPE")
    app.setOrganizationName("Robô SIAPE")
    aplicar_tema(app)

    janela = App()
    janela.show()

    app.exec()


if __name__ == "__main__":
    main()