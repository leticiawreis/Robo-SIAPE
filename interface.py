import hashlib
import json
import logging
import os
import queue
import re
import threading
from datetime import date

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QColor, QPainter, QPen
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from robo_siape import INDICES_ANOS, meses_disponiveis, executar_pipeline_completo


ARQUIVO_USUARIOS = "usuarios.json"
PASTA_SAIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "saida")

def preparar_pastas():
    os.makedirs(PASTA_SAIDA, exist_ok=True)

BG = "#090909"
SURFACE = "#111111"
SURFACE_2 = "#151515"
SURFACE_3 = "#1B1B1B"
BORDER = "#292929"
YELLOW = "#D6A84F"
YELLOW_LIGHT = "#F0C96A"
TEXT = "#F5F1E8"
MUTED = "#9B978E"
SUCCESS = "#74C69D"
ERROR = "#E57373"
WARNING = "#E5B85C"


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


def aplicar_sombra(widget, blur=28, y=8, alpha=90):
    sombra = QGraphicsDropShadowEffect()
    sombra.setBlurRadius(blur)
    sombra.setOffset(0, y)
    sombra.setColor(QColor(0, 0, 0, alpha))
    widget.setGraphicsEffect(sombra)


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
        aplicar_sombra(self)


class CaptchaDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ação necessária")
        self.setModal(True)
        self.setMinimumWidth(470)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 28, 30, 25)
        layout.setSpacing(14)

        icon = QLabel("!")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setFixedSize(48, 48)
        icon.setStyleSheet(
            f"background:{YELLOW}; color:#111; border-radius:24px; font-size:24px; font-weight:700;"
        )
        layout.addWidget(icon, alignment=Qt.AlignmentFlag.AlignHCenter)

        titulo = QLabel("CAPTCHA detectado")
        titulo.setObjectName("dialogTitle")
        titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(titulo)

        texto = QLabel(
            "O navegador precisa de uma confirmação manual.\n\n"
            "Resolva o CAPTCHA na janela do Chrome e, quando terminar, "
            "clique em “Continuar” abaixo."
        )
        texto.setObjectName("dialogText")
        texto.setWordWrap(True)
        texto.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(texto)

        botoes = QDialogButtonBox()
        continuar = botoes.addButton(
            "Continuar", QDialogButtonBox.ButtonRole.AcceptRole
        )
        continuar.setObjectName("dialogPrimary")
        botoes.accepted.connect(self.accept)
        layout.addWidget(botoes)

        self.setStyleSheet(
            f"""
            QDialog {{
                background: {SURFACE};
                color: {TEXT};
            }}
            QLabel#dialogTitle {{
                color: {TEXT};
                font-size: 20px;
                font-weight: 700;
            }}
            QLabel#dialogText {{
                color: {MUTED};
                font-size: 13px;
                line-height: 1.5;
            }}
            QPushButton#dialogPrimary {{
                background: {YELLOW};
                color: #111111;
                border: none;
                border-radius: 9px;
                min-height: 44px;
                padding: 0 24px;
                font-weight: 700;
            }}
            QPushButton#dialogPrimary:hover {{
                background: {YELLOW_LIGHT};
            }}
            """
        )


class Worker(QThread):
    log_signal = pyqtSignal(str)
    captcha_signal = pyqtSignal()
    done_signal = pyqtSignal(str, str)
    error_signal = pyqtSignal(str)

    def __init__(self, ano, mes):
        super().__init__()
        self.ano = ano
        self.mes = mes
        self.log_queue = queue.Queue()
        self.captcha_event = threading.Event()

    def run(self):
        thread_fila = threading.Thread(target=self._processar_fila, daemon=True)
        thread_fila.start()

        try:
            caminho, caminho_log = executar_pipeline_completo(
                self.ano,
                self.mes,
                self.log_queue,
                self.captcha_event,
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
            elif tipo == "captcha":
                self.captcha_signal.emit()
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
        content_layout.setContentsMargins(48, 38, 48, 35)
        content_layout.addWidget(self.tela)

        central_layout.addWidget(self.conteudo, 1)
        self.setCentralWidget(central)

    def _criar_sidebar(self):
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(250)

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(24, 30, 24, 28)
        layout.setSpacing(10)

        topo = QHBoxLayout()
        topo.setSpacing(12)
        topo.addWidget(LogoMark())

        marca = QVBoxLayout()
        marca.setSpacing(0)
        nome = QLabel("ROBÔ SIAPE")
        nome.setObjectName("brandName")
        marca.addWidget(nome)
        subtitulo = QLabel("AUTOMAÇÃO INTELIGENTE")
        subtitulo.setObjectName("brandSub")
        marca.addWidget(subtitulo)
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
        self.nav_log = self._nav_button("≡", "Execução")

        self.nav_home.clicked.connect(self.tela_principal)
        self.nav_robot.clicked.connect(self.tela_config_robo)
        self.nav_log.clicked.connect(self.tela_execucao_placeholder)

        layout.addWidget(self.nav_home)
        layout.addWidget(self.nav_robot)
        layout.addWidget(self.nav_log)

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
        layout.setSpacing(24)
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

        pagina, layout = self._pagina()

        topo = QHBoxLayout()
        topo.addWidget(
            criar_titulo(
                "Painel de controle",
                "Automação de dados de remuneração do Portal da Transparência.",
            )
        )
        topo.addStretch()

        badge = QLabel("●  PRONTO")
        badge.setObjectName("statusBadge")
        topo.addWidget(badge, alignment=Qt.AlignmentFlag.AlignTop)
        layout.addLayout(topo)

        cards = QHBoxLayout()
        cards.setSpacing(16)

        cards.addWidget(
            self._info_card("01", "Selecionar período", "Escolha ano e mês para iniciar.")
        )
        cards.addWidget(
            self._info_card("02", "Processamento", "Download, tratamento e validação.")
        )
        cards.addWidget(
            self._info_card("03", "Resultado", "Excel final e log em /saida.")
        )
        layout.addLayout(cards)

        armazenamento = Card()
        armazenamento_layout = QHBoxLayout(armazenamento)
        armazenamento_layout.setContentsMargins(22, 18, 22, 18)
        armazenamento_layout.setSpacing(20)

        armazenamento_info = QVBoxLayout()
        armazenamento_info.setSpacing(3)

        armazenamento_titulo = QLabel("ARMAZENAMENTO")
        armazenamento_titulo.setObjectName("eyebrow")
        armazenamento_info.addWidget(armazenamento_titulo)

        armazenamento_texto = QLabel("Arquivos finais da execução")
        armazenamento_texto.setObjectName("cardTitle")
        armazenamento_info.addWidget(armazenamento_texto)

        armazenamento_layout.addLayout(armazenamento_info, 1)

        pastas = QVBoxLayout()
        pastas.setSpacing(2)

        pasta_bruta = QLabel(f"Planilha bruta  •  {os.path.join(PASTA_SAIDA, 'base_bruta_AAAA_MM.csv')}")
        pasta_bruta.setObjectName("storagePath")
        pastas.addWidget(pasta_bruta)

        pasta_formatada = QLabel(f"Planilha formatada  •  {os.path.join(PASTA_SAIDA, 'base_tratada_AAAA_MM.xlsx')}")
        pasta_formatada.setObjectName("storagePath")
        pastas.addWidget(pasta_formatada)

        pasta_log = QLabel(f"Log  •  {os.path.join(PASTA_SAIDA, 'execucao_AAAA_MM.log')}")
        pasta_log.setObjectName("storagePath")
        pastas.addWidget(pasta_log)

        armazenamento_layout.addLayout(pastas, 1)
        layout.addWidget(armazenamento)

        destaque = Card()
        dlayout = QHBoxLayout(destaque)
        dlayout.setContentsMargins(26, 24, 26, 24)

        left = QVBoxLayout()
        tag = QLabel("AUTOMAÇÃO SIAPE")
        tag.setObjectName("eyebrow")
        left.addWidget(tag)

        titulo = QLabel("Pronto para iniciar uma nova execução?")
        titulo.setObjectName("cardTitle")
        left.addWidget(titulo)

        desc = QLabel(
            "O robô navega pelo portal, baixa a base, remove dados desnecessários "
            "e gera a planilha tratada automaticamente."
        )
        desc.setObjectName("cardText")
        desc.setWordWrap(True)
        left.addWidget(desc)
        dlayout.addLayout(left, 1)

        iniciar = Botao("CONFIGURAR ROBÔ  →", principal=True)
        iniciar.setMinimumWidth(220)
        iniciar.clicked.connect(self.tela_config_robo)
        dlayout.addWidget(iniciar)

        layout.addWidget(destaque)
        layout.addStretch()

    def _info_card(self, numero, titulo, texto):
        card = Card()
        card.setMinimumHeight(155)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(20, 20, 20, 20)
        num = QLabel(numero)
        num.setObjectName("cardNumber")
        lay.addWidget(num)
        t = QLabel(titulo)
        t.setObjectName("cardTitle")
        lay.addWidget(t)
        d = QLabel(texto)
        d.setObjectName("cardText")
        d.setWordWrap(True)
        lay.addWidget(d)
        lay.addStretch()
        return card

    def tela_config_robo(self):
        self.limpar_tela()
        self.sidebar.show()
        self._definir_usuario()

        pagina, layout = self._pagina()
        layout.addWidget(
            criar_titulo(
                "Nova execução",
                "Defina o período que será consultado e processado.",
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

        texto = QLabel("Selecione um período disponível no Portal da Transparência.")
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
            "O sistema bloqueia automaticamente períodos futuros e mantém "
            "as regras de disponibilidade do portal."
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
            ("01", "Abrir Portal da Transparência"),
            ("02", "Selecionar ano e mês"),
            ("03", "Baixar e validar ZIP"),
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

    def validar_periodo(self, ano, mes):
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

        if not ano or not mes:
            self._mensagem("Atenção", "Selecione o ano e o mês.", "warning")
            return

        valido, mensagem = self.validar_periodo(ano, mes)

        if not valido:
            self._mensagem("Entrada inválida", mensagem, "error")
            return

        self.tela_execucao(ano, mes)

    def tela_execucao(self, ano, mes):
        self.limpar_tela()
        self.sidebar.show()
        self._definir_usuario()

        pagina, layout = self._pagina()

        topo = QHBoxLayout()
        topo.addWidget(
            criar_titulo(
                "Execução em andamento",
                f"Processando remuneração — {mes.capitalize()}/{ano}",
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

        self.exec_periodo = QLabel(f"{mes.capitalize()} / {ano}")
        self.exec_periodo.setObjectName("periodValue")
        ilayout.addWidget(self.exec_periodo)

        ilayout.addStretch()

        self.exec_hint = QLabel("O navegador pode solicitar uma ação manual.")
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
        self.captcha_event = threading.Event()
        self.log_count = 0

        self.worker = Worker(ano, mes)
        self.worker.log_signal.connect(self._log_gui)
        self.worker.captcha_signal.connect(self._mostrar_captcha)
        self.worker.done_signal.connect(self._execucao_concluida)
        self.worker.error_signal.connect(self._execucao_erro)
        self.worker.start()

    def _log_gui(self, texto):
        self.log_count += 1
        self.text_log.append(texto)
        self.log_counter.setText(
            f"{self.log_count} evento" + ("" if self.log_count == 1 else "s")
        )

    def _mostrar_captcha(self):
        dialogo = CaptchaDialog(self)
        dialogo.exec()
        if self.worker:
            self.worker.captcha_event.set()

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
        QMainWindow, QWidget {{
            background: {BG};
            color: {TEXT};
        }}
        QFrame#sidebar {{
            background: #0C0C0C;
            border-right: 1px solid {BORDER};
        }}
        QFrame#contentArea {{
            background: {BG};
        }}
        QLabel#brandName {{
            color: {TEXT};
            font-size: 16px;
            font-weight: 800;
            letter-spacing: 1px;
        }}
        QLabel#brandSub {{
            color: {YELLOW};
            font-size: 8px;
            font-weight: 700;
            letter-spacing: 1px;
        }}
        QFrame#separator {{
            color: {BORDER};
            background: {BORDER};
            max-height: 1px;
        }}
        QLabel#sideSection, QLabel#sideCaption {{
            color: {MUTED};
            font-size: 10px;
            font-weight: 700;
            letter-spacing: 1px;
        }}
        QLabel#sideUser {{
            color: {TEXT};
            font-size: 13px;
            font-weight: 700;
        }}
        QFrame#userCard {{
            background: {SURFACE};
            border: 1px solid {BORDER};
            border-radius: 12px;
        }}
        QPushButton#navButton {{
            background: transparent;
            color: {MUTED};
            border: none;
            border-radius: 9px;
            text-align: left;
            padding: 0 14px;
            font-size: 13px;
            font-weight: 600;
        }}
        QPushButton#navButton:hover {{
            background: {SURFACE_2};
            color: {TEXT};
        }}
        QPushButton#navButton:pressed {{
            background: {SURFACE_3};
            color: {YELLOW_LIGHT};
        }}
        QLabel#pageTitle {{
            color: {TEXT};
            font-size: 30px;
            font-weight: 800;
        }}
        QLabel#pageSubtitle {{
            color: {MUTED};
            font-size: 13px;
        }}
        QFrame#card {{
            background: {SURFACE};
            border: 1px solid {BORDER};
            border-radius: 15px;
        }}
        QLabel#cardNumber {{
            color: {YELLOW};
            font-size: 12px;
            font-weight: 800;
            letter-spacing: 1px;
        }}
        QLabel#cardTitle {{
            color: {TEXT};
            font-size: 17px;
            font-weight: 700;
        }}
        QLabel#cardText {{
            color: {MUTED};
            font-size: 12px;
            line-height: 1.4;
        }}
        QLabel#eyebrow {{
            color: {YELLOW};
            font-size: 10px;
            font-weight: 800;
            letter-spacing: 1.5px;
        }}
        QLabel#fieldLabel {{
            color: {MUTED};
            font-size: 10px;
            font-weight: 800;
            letter-spacing: 1px;
        }}
        QLabel#hint {{
            color: {MUTED};
            font-size: 11px;
            padding: 10px 0;
        }}
        QLabel#statusBadge, QLabel#statusBadgeRunning, QLabel#statusBadgeError {{
            background: rgba(214, 168, 79, 0.12);
            color: {YELLOW_LIGHT};
            border: 1px solid rgba(214, 168, 79, 0.35);
            border-radius: 14px;
            padding: 7px 12px;
            font-size: 10px;
            font-weight: 800;
        }}
        QLabel#statusBadgeRunning {{
            color: {YELLOW_LIGHT};
        }}
        QLabel#statusBadgeError {{
            color: {ERROR};
            border-color: rgba(229, 115, 115, 0.35);
            background: rgba(229, 115, 115, 0.08);
        }}
        QLabel#authTitle {{
            color: {TEXT};
            font-size: 28px;
            font-weight: 800;
        }}
        QLabel#authSubtitle {{
            color: {MUTED};
            font-size: 12px;
        }}
        QPushButton[principal="true"] {{
            background: {YELLOW};
            color: #111111;
            border: none;
            border-radius: 10px;
            padding: 0 20px;
            font-size: 12px;
            font-weight: 800;
            letter-spacing: 0.5px;
        }}
        QPushButton[principal="true"]:hover {{
            background: {YELLOW_LIGHT};
        }}
        QPushButton[principal="true"]:pressed {{
            background: #B88D3B;
        }}
        QPushButton[principal="true"]:disabled {{
            background: #4C402A;
            color: #8D826A;
        }}
        QPushButton[compacto="true"] {{
            background: {SURFACE_2};
            color: {TEXT};
            border: 1px solid {BORDER};
            border-radius: 9px;
            padding: 0 18px;
            font-size: 11px;
            font-weight: 700;
        }}
        QPushButton[compacto="true"]:hover {{
            border-color: {YELLOW};
            color: {YELLOW_LIGHT};
        }}
        QPushButton[compacto="true"]:disabled {{
            color: #5C5A55;
        }}
        QPushButton#linkButton {{
            background: transparent;
            border: none;
            color: {YELLOW};
            font-size: 11px;
            font-weight: 600;
            padding: 8px;
        }}
        QPushButton#linkButton:hover {{
            color: {YELLOW_LIGHT};
        }}
        QTextEdit#logConsole {{
            background: #080808;
            color: #D9D3C8;
            border: 1px solid #202020;
            border-radius: 10px;
            padding: 12px;
            font-family: "Cascadia Mono", "Consolas", monospace;
            font-size: 11px;
            selection-background-color: {YELLOW};
            selection-color: #111111;
        }}
        QLabel#logCounter {{
            color: {MUTED};
            font-size: 10px;
        }}
        QLabel#periodValue {{
            color: {YELLOW_LIGHT};
            font-size: 17px;
            font-weight: 800;
        }}
        QLabel#resultPath {{
            color: {SUCCESS};
            font-size: 10px;
            font-weight: 700;
        }}
        QLabel#storagePath {{
            color: {MUTED};
            font-size: 10px;
        }}
        QLabel#flowNumber {{
            color: {YELLOW};
            font-size: 11px;
            font-weight: 800;
        }}
        QLabel#flowText {{
            color: {TEXT};
            font-size: 11px;
        }}
        QScrollBar:vertical {{
            background: #0C0C0C;
            width: 8px;
            margin: 0;
        }}
        QScrollBar::handle:vertical {{
            background: #3A3833;
            border-radius: 4px;
            min-height: 30px;
        }}
        QScrollBar::handle:vertical:hover {{
            background: {YELLOW};
        }}
        QMessageBox {{
            background: {SURFACE};
        }}
        QMessageBox QLabel {{
            color: {TEXT};
            font-size: 13px;
        }}
        QMessageBox QPushButton {{
            background: {YELLOW};
            color: #111111;
            border: none;
            border-radius: 8px;
            padding: 8px 18px;
            min-width: 80px;
            font-weight: 700;
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
