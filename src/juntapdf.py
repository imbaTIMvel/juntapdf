import sys, os
from pathlib import Path
import win32com.client
from pypdf import PdfWriter
from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QLabel,
    QLineEdit,
    QVBoxLayout,
    QHBoxLayout,
    QProgressBar,
    QFileDialog,
    QFrame,
    QMessageBox
)
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import Qt

# ==============================
# FUNÇÕES AUXILIARES
# ==============================

def atualizar_progresso(valor, total, texto=""):
    progresso = int((valor / total) * 100) if total > 0 else 0
    progress.setValue(progresso)
    label_status.setText(texto)
    app.processEvents()

def selecionar_pasta():
    pasta = QFileDialog.getExistingDirectory(
        window,
        "Selecione a pasta"
    )
    if pasta:
        entry_pasta.setText(pasta)
        btn_compilar.setEnabled(True)

def selecionar_saida():
    pasta = QFileDialog.getExistingDirectory(
        window,
        "Selecione a pasta de saída"
    )
    return pasta

def converter_lote_docs(lista_arquivos, callback=None, progresso_inicial=0, total=1):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    convertidos = []

    for i, arquivo in enumerate(lista_arquivos, start=1):
        pdf_path = arquivo.with_suffix(".pdf")

        doc = word.Documents.Open(str(arquivo))
        doc.SaveAs(str(pdf_path), FileFormat=17) # 17 = PDF
        doc.Close()

        convertidos.append(pdf_path)

        if callback:
            callback(progresso_inicial + i, total, f"Convertendo: {arquivo.name}")

    word.Quit()

    return convertidos

def extrair_base_e_indice(nome_arquivo):
    """
    Separa usando o ÚLTIMO underscore.
    Ex: relatorio_final_3.pdf → ("relatorio_final", 3)
    """
    nome = nome_arquivo.stem  # sem .pdf

    if "_" not in nome:
        return None, None

    base, indice = nome.rsplit("_", 1)

    if not indice.isdigit():
        return None, None

    return base, int(indice)

# ==============================
# FUNÇÃO PRINCIPAL
# ==============================

def compilar_pdfs():
    btn_compilar.setEnabled(False)
    arquivos_convertidos = []
    
    try:
        pasta = Path(entry_pasta.text())

        if not pasta.exists():
            QMessageBox.critical(window, "Erro", "Pasta inválida.")
            return

        arquivos_pdf = list(pasta.glob("*.pdf"))
        arquivos_docx = list(pasta.glob("*.docx"))
        arquivos_doc = list(pasta.glob("*.doc"))

        docs = arquivos_docx + arquivos_doc
        arquivos_convertidos = converter_lote_docs(
            docs,
            callback=atualizar_progresso,
            progresso_inicial=0,
            total=1 if len(docs) == 0 else len(docs)
        )
        arquivos_pdf.extend(arquivos_convertidos)

        grupos = {}
        for arquivo in arquivos_pdf:
            base, indice = extrair_base_e_indice(arquivo)

            if base is None:
                continue

            grupos.setdefault(base, []).append((indice, arquivo))

        if not grupos:
            QMessageBox.warning(window, "Aviso", "Nenhum arquivo válido encontrado.")
            return
        
        total_passos = len(docs) + len(grupos)
        passo_atual = len(docs)

        if not grupos:
            QMessageBox.warning(window, "Aviso", "Nenhum arquivo válido encontrado.")
            return

        pasta_saida = QFileDialog.getExistingDirectory(
            window,
            "Selecione a pasta para salvar os PDFs compilados"
        )

        if not pasta_saida:
            return  # usuário cancelou

        pasta_saida = Path(pasta_saida)

        # Processar e salvar
        for base, lista in grupos.items():
            lista_ordenada = sorted(lista, key=lambda x: x[0])

            merger = PdfWriter()

            for _, arquivo in lista_ordenada:
                merger.append(str(arquivo))

            saida = pasta_saida / f"{base}.pdf"

            if saida.exists():
                resposta = QMessageBox.question(
                    window,
                    "Sobrescrever",
                    f"O arquivo {base}.pdf já existe. Deseja sobrescrevê-lo?",
                    QMessageBox.StandardButton.Yes |
                    QMessageBox.StandardButton.No
                )
                if resposta == QMessageBox.StandardButton.No:
                    continue
                
                saida.unlink()

            try:
                with open(saida, "wb") as f:
                    merger.write(f)
                merger.close()
            except Exception as e:
                QMessageBox.information(
                    window,
                    "Erro",
                    f"Erro ao gerar {base}.pdf\n{e}"
                )
                return
            
            passo_atual += 1
            atualizar_progresso(passo_atual, total_passos, f"Gerando: {base}.pdf")
        
        QMessageBox.information(window, "Sucesso", f"{len(grupos)} PDF(s) compilados com sucesso!")
    
    finally:
        for pdf in arquivos_convertidos:
            try:
                pdf.unlink()
            except Exception as e:
                pass
        
        progress.setValue(0)
        label_status.setText("")
        btn_compilar.setEnabled(True)

# ==============================
# RESOURCE PATH
# ==============================

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# ==============================
# WINDOW
# ==============================

app = QApplication(sys.argv)

window = QWidget()
window.setWindowTitle("JuntaPDF")
window.setWindowIcon(QIcon(resource_path("juntapdf.ico")))
window.resize(700, 400)

# ==============================
# BACKGROUND
# ==============================

window.setStyleSheet("""
QWidget {
    background-color: #1e1e1e;
}
""")

# ==============================
# BACKGROUND IMAGE
# ==============================

bg_label = QLabel(window)
bg_pixmap = QPixmap(resource_path("bg_hbr.png"))

bg_label.setPixmap(bg_pixmap)
bg_label.setScaledContents(True)

# ==============================
# CARD CENTRAL
# ==============================

card = QFrame(window)
card.setObjectName("card")

card.setStyleSheet("""
#card {
    background-color: rgba(30, 30, 30, 220);
    border-radius: 16px;
}
""")

# ==============================
# TITLE
# ==============================

titulo = QLabel("JuntaPDF")
titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)

titulo.setStyleSheet("""
font-family: "Bahnschrift Condensed";
font-size: 42px;
font-weight: bold;
color: white;
padding-bottom: 10px;
""")

# ==============================
# INPUT
# ==============================

entry_pasta = QLineEdit()

entry_pasta.setPlaceholderText("Selecione uma pasta...")

entry_pasta.setStyleSheet("""
QLineEdit {
    background-color: #2b2b2b;
    color: white;
    border: 1px solid #3a3a3a;
    border-radius: 8px;
    padding: 8px;
    font-size: 14px;
}
""")

# ==============================
# BUTTON STYLE
# ==============================

button_style = """
QPushButton {
    background-color: #f9b02e;
    color: black;
    border: none;
    border-radius: 8px;
    padding: 10px;
    font-weight: bold;
    font-size: 14px;
}

QPushButton:hover {
    background-color: #ffd166;
}

QPushButton:pressed {
    background-color: #e69500;
}

QPushButton:disabled {
    background-color: #666666;
    color: #aaaaaa;
}
"""

# ==============================
# BUTTONS
# ==============================

btn_selecionar = QPushButton("Selecionar Pasta")
btn_selecionar.setStyleSheet(button_style)
btn_selecionar.clicked.connect(selecionar_pasta)

btn_compilar = QPushButton("Juntar PDFs")
btn_compilar.setStyleSheet(button_style)
btn_compilar.setEnabled(False)
btn_compilar.clicked.connect(compilar_pdfs)

# ==============================
# PROGRESS BAR
# ==============================

progress = QProgressBar()

progress.setStyleSheet("""
QProgressBar {
    background-color: #2b2b2b;
    border-radius: 6px;
    text-align: center;
    color: white;
    height: 18px;
}

QProgressBar::chunk {
    background-color: #f9b02e;
    border-radius: 6px;
}
""")

# ==============================
# STATUS
# ==============================

label_status = QLabel("")
label_status.setAlignment(Qt.AlignmentFlag.AlignCenter)

label_status.setStyleSheet("""
color: #cccccc;
font-size: 12px;
padding-top: 5px;
""")

# ==============================
# FOOTER / ASSINATURA
# ==============================

github_label = QLabel()
github_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

github_label.setText(
    '<a href="https://github.com/imbaTIMvel/juntapdf">'
    'JuntaPDF v0.1.0 - GitHub'
    '</a>'
)

github_label.setOpenExternalLinks(True)

github_label.setStyleSheet("""
QLabel {
    background-color: transparent;
    color: rgba(255,255,255,120);
    font-size: 11px;
}

QLabel:hover {
    color: #f9b02e;
}
""")

footer = QLabel(
    "Desenvolvido por: Diretoria Administrativa Financeira - DAF"
)

footer.setAlignment(Qt.AlignmentFlag.AlignCenter)

footer.setStyleSheet("""
QLabel {
    background-color: transparent;
    color: rgba(255,255,255,120);
    font-size: 10px;
    padding-bottom: 4px;
}
""")

# ==============================
# LAYOUTS
# ==============================

input_layout = QHBoxLayout()
input_layout.addWidget(entry_pasta)
input_layout.addWidget(btn_selecionar)

card_layout = QVBoxLayout(card)

card_layout.addWidget(titulo)
card_layout.addLayout(input_layout)
card_layout.addWidget(btn_compilar)
card_layout.addWidget(progress)
card_layout.addWidget(label_status)

card_layout.setSpacing(15)
card_layout.setContentsMargins(25, 25, 25, 25)

main_layout = QVBoxLayout(window)

main_layout.addStretch()
main_layout.addWidget(card, alignment=Qt.AlignmentFlag.AlignCenter)
main_layout.addStretch()
main_layout.addWidget(github_label, alignment=Qt.AlignmentFlag.AlignBottom)
main_layout.addWidget(footer, alignment=Qt.AlignmentFlag.AlignBottom)

main_layout.setContentsMargins(40, 40, 40, 40)

window.setLayout(main_layout)

# ==============================
# RESPONSIVE BACKGROUND
# ==============================

def resize_event(event):
    bg_label.resize(window.size())
    card.setMaximumWidth(600)

window.resizeEvent = resize_event

# ==============================
# SHOW
# ==============================

window.show()

sys.exit(app.exec())