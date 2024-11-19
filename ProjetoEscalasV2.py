import json
import sys
import win32com.client  # Importação para interagir com o Outlook
from PyQt5.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QVBoxLayout, QLabel, QComboBox, QDateTimeEdit,
    QPushButton, QLineEdit, QCompleter, QTableWidget, QTableWidgetItem, QHeaderView,
    QSizePolicy, QFileDialog, QDialog, QDateEdit, QMessageBox, QListWidget, QListWidgetItem,
    QAbstractItemView, QRadioButton, QButtonGroup, QScrollArea, QFormLayout, QCheckBox
)
from PyQt5.QtCore import (
    Qt, QDateTime, QDate, QTime, QStringListModel, QTimer, QSortFilterProxyModel,
    QRegularExpression, QPoint, pyqtSignal, QEvent
)
from PyQt5.QtGui import QFont, QIcon, QMouseEvent, QColor, QLinearGradient, QBrush
import pandas as pd  # Biblioteca para manipulação de planilhas Excel
from openpyxl import load_workbook
from openpyxl.styles import numbers
from datetime import timedelta
import datetime  # Importação do módulo datetime
import locale  # Importação para configurar a localidade

# Dicionário com e-mails dos gestores das unidades
unit_manager_emails = {
    "AMA Zaio": "allef.barbosa@libertyti.com.br",
    # Adicione as outras unidades e seus respectivos e-mails
}

# Dicionário com e-mails dos técnicos
technician_emails = {
    "Allef Barbosa": "allef.barbosa@libertyti.com.br",
    "Vinicius Oliveira": "vinicius.oliveira@libertyti.com.br",
    "Eduardo Lima": "eduardo.lima@libertyti.com.br",
    "Ivaldo Junior": "ivaldo.junior@libertyti.com.br",
    "Kaue Rodrigues": "kaue.rodrigues@libertyti.com.br",
    "Geovanna Oliveira": "geovanna.oliveira@libertyti.com.br",
    "Gustavo Silva": "gustavo.silva@libertyti.com.br",
    "Vitor Martins": "vitor.martins@libertyti.com.br",
    "Mateus Marinho": "vinicius.santos@libertyti.com.br",
    "Joao Marinho": "joao.marinho@libertyti.com.br",
    "Andre Assis": "andre.assis@libertyti.com.br"
}

# Classe ClickableLabel
class ClickableLabel(QLabel):
    clicked = pyqtSignal()

    def mousePressEvent(self, event):
        self.clicked.emit()

# Classe SubstringFilterProxyModel
class SubstringFilterProxyModel(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.setFilterRole(Qt.DisplayRole)

    def filterAcceptsRow(self, sourceRow, sourceParent):
        if not self.filterRegularExpression().pattern():
            return True  # Sem filtro, aceita todas as linhas
        index = self.sourceModel().index(sourceRow, self.filterKeyColumn(), sourceParent)
        data = self.sourceModel().data(index, self.filterRole())
        if self.filterRegularExpression().match(data).hasMatch():
            return True
        return False

# Classe FilteredComboBox
class FilteredComboBox(QComboBox):
    def __init__(self, items, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)

        # Modelo base com todos os itens
        self.model_base = QStringListModel(items)

        # Proxy model personalizado para filtragem por substring
        self.proxy_model = SubstringFilterProxyModel(self)
        self.proxy_model.setSourceModel(self.model_base)

        # Configuração do QCompleter com o proxy model
        self.completer = QCompleter(self.proxy_model, self)
        self.completer.setCompletionMode(QCompleter.PopupCompletion)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.setCompleter(self.completer)

        # Conectar ao método de edição
        self.lineEdit().textEdited.connect(self.filter_items)

        # Evento de tecla para detectar Enter
        self.lineEdit().installEventFilter(self)

        # Estilo personalizado
        self.setStyleSheet("""
            QComboBox {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 6px;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
                background-color: #fff;
            }
            QComboBox QAbstractItemView {
                background-color: #fff;
                selection-background-color: #007bff;
                selection-color: #fff;
            }
            QComboBox::down-arrow {
                image: url('icons/down-arrow.png');
                width: 14px;
                height: 14px;
            }
        """)

    def filter_items(self, text):
        # Atualiza o filtro do proxy model para considerar qualquer substring
        pattern = f".*{QRegularExpression.escape(text)}.*"
        regex = QRegularExpression(pattern, QRegularExpression.CaseInsensitiveOption)
        self.proxy_model.setFilterRegularExpression(regex)
        self.completer.complete()

    def showPopup(self):
        # Atualiza o filtro ao exibir o popup
        self.filter_items(self.lineEdit().text())
        super().showPopup()

    def eventFilter(self, source, event):
        if event.type() == QEvent.KeyPress and source is self.lineEdit():
            if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                self.parent().handle_enter_press()
        return super().eventFilter(source, event)

# Classe SelectionDialog
class SelectionDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon('JT.ico'))
        self.planilha_path = None
        self.periodo_inicio = None
        self.periodo_fim = None
        self.choice = None  # Adicionado para armazenar a escolha do usuário
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Seleção de Planilha e Período")
        self.resize(400, 250)
        layout = QVBoxLayout()

        # Botão para selecionar a planilha
        self.planilha_label = QLabel("Nenhuma planilha selecionada")
        self.planilha_label.setAlignment(Qt.AlignCenter)
        self.planilha_button = QPushButton("Selecionar Planilha")
        self.planilha_button.clicked.connect(self.select_planilha)
        self.planilha_button.setStyleSheet(self.get_button_style())

        # Seleção do período
        periodo_layout = QHBoxLayout()
        self.data_inicio = QDateEdit()
        self.data_inicio.setCalendarPopup(True)
        self.data_inicio.setDate(QDate.currentDate())
        self.data_fim = QDateEdit()
        self.data_fim.setCalendarPopup(True)
        self.data_fim.setDate(QDate.currentDate())
        periodo_layout.addWidget(QLabel("Data Início:"))
        periodo_layout.addWidget(self.data_inicio)
        periodo_layout.addWidget(QLabel("Data Fim:"))
        periodo_layout.addWidget(self.data_fim)

        # Botões de opção para "Incluir Escala" e "Consultar Escala"
        self.radio_incluir_escala = QRadioButton("Incluir Escala")
        self.radio_consultar_escala = QRadioButton("Consultar Escala")
        self.radio_incluir_escala.setChecked(True)  # Define "Incluir Escala" como padrão

        # Agrupar os botões para que apenas um possa ser selecionado
        self.choice_group = QButtonGroup()
        self.choice_group.addButton(self.radio_incluir_escala)
        self.choice_group.addButton(self.radio_consultar_escala)

        # Layout para os botões de opção
        choice_layout = QHBoxLayout()
        choice_layout.addWidget(self.radio_incluir_escala)
        choice_layout.addWidget(self.radio_consultar_escala)

        # Botão de confirmar
        self.confirm_button = QPushButton("Confirmar")
        self.confirm_button.clicked.connect(self.confirm_selection)
        self.confirm_button.setEnabled(False)  # Desabilitado até que a planilha seja selecionada
        self.confirm_button.setStyleSheet(self.get_button_style())

        layout.addWidget(self.planilha_label)
        layout.addWidget(self.planilha_button)
        layout.addLayout(periodo_layout)
        layout.addLayout(choice_layout)  # Adicionado os botões de opção ao layout
        layout.addWidget(self.confirm_button)

        self.setLayout(layout)

    def select_planilha(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Planilha", "", "Planilhas Excel (*.xlsx *.xls)", options=options
        )
        if file_path:
            self.planilha_path = file_path
            self.planilha_label.setText(f"Planilha selecionada:\n{file_path}")
            self.confirm_button.setEnabled(True)

    def confirm_selection(self):
        self.periodo_inicio = self.data_inicio.date()
        self.periodo_fim = self.data_fim.date()
        if self.periodo_fim < self.periodo_inicio:
            # Validação básica
            QMessageBox.warning(self, "Erro", "Data fim não pode ser antes da data início.")
        else:
            # Captura a escolha do usuário
            if self.radio_incluir_escala.isChecked():
                self.choice = 'incluir_escala'
            elif self.radio_consultar_escala.isChecked():
                self.choice = 'consultar_escala'
            self.accept()

    def get_button_style(self):
        return """
            QPushButton {
                background-color: #007bff;
                color: white;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:disabled {
                background-color: #6c757d;
            }
        """

# Classe PeriodoConsultaDialog
class PeriodoConsultaDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon('JT.ico'))
        self.periodo_inicio = None
        self.periodo_fim = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Período de Consulta")
        self.resize(300, 150)
        layout = QVBoxLayout()

        periodo_layout = QHBoxLayout()
        self.data_inicio = QDateEdit()
        self.data_inicio.setCalendarPopup(True)
        self.data_inicio.setDate(QDate.currentDate())
        self.data_fim = QDateEdit()
        self.data_fim.setCalendarPopup(True)
        self.data_fim.setDate(QDate.currentDate())
        periodo_layout.addWidget(QLabel("Data Início:"))
        periodo_layout.addWidget(self.data_inicio)
        periodo_layout.addWidget(QLabel("Data Fim:"))
        periodo_layout.addWidget(self.data_fim)

        self.confirm_button = QPushButton("Confirmar")
        self.confirm_button.clicked.connect(self.confirm_selection)
        self.confirm_button.setStyleSheet(self.get_button_style())

        layout.addLayout(periodo_layout)
        layout.addWidget(self.confirm_button)

        self.setLayout(layout)

    def confirm_selection(self):
        self.periodo_inicio = self.data_inicio.date()
        self.periodo_fim = self.data_fim.date()
        if self.periodo_fim < self.periodo_inicio:
            QMessageBox.warning(self, "Erro", "Data fim não pode ser antes da data início.")
        else:
            self.accept()

    def get_button_style(self):
        return """
            QPushButton {
                background-color: #007bff;
                color: white;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """

# Classe TechnicianSelectionDialog
class TechnicianSelectionDialog(QDialog):
    def __init__(self, tecnicos, technician_schedules):
        super().__init__()
        self.setWindowIcon(QIcon('JT.ico'))
        self.selected_tecnicos = {}
        self.technician_schedules = technician_schedules
        self.init_ui(tecnicos)

    def init_ui(self, tecnicos):
        self.setWindowTitle("Selecionar Técnicos")
        self.resize(400, 500)
        layout = QVBoxLayout()

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        widget = QWidget()
        self.scroll_area.setWidget(widget)
        form_layout = QFormLayout(widget)

        self.tecnico_widgets = {}

        for tecnico in tecnicos:
            tecnico_info = self.technician_schedules.get(tecnico, {})
            escala = tecnico_info.get('escala')
            hbox = QHBoxLayout()
            checkbox = QCheckBox()
            hbox.addWidget(checkbox)
            label = QLabel(tecnico)
            hbox.addWidget(label)

            if escala == '12X36':
                # Adicionar opções de Pares ou Ímpares sem seleção prévia
                radiobutton_pares = QRadioButton("Pares")
                radiobutton_impares = QRadioButton("Ímpares")

                # Agrupar os botões de rádio para permitir apenas uma seleção
                button_group = QButtonGroup(self)
                button_group.addButton(radiobutton_pares)
                button_group.addButton(radiobutton_impares)

                hbox.addWidget(radiobutton_pares)
                hbox.addWidget(radiobutton_impares)
                self.tecnico_widgets[tecnico] = (checkbox, radiobutton_pares, radiobutton_impares)
            else:
                # Para técnicos 5x2, adicionar espaçamento para alinhamento
                hbox.addStretch()
                self.tecnico_widgets[tecnico] = (checkbox, None, None)

            form_layout.addRow(hbox)

        self.confirm_button = QPushButton("Confirmar")
        self.confirm_button.clicked.connect(self.confirm_selection)
        self.confirm_button.setStyleSheet(self.get_button_style())

        layout.addWidget(self.scroll_area)
        layout.addWidget(self.confirm_button)

        self.setLayout(layout)

    def confirm_selection(self):
        self.selected_tecnicos = {}
        for tecnico, widgets in self.tecnico_widgets.items():
            checkbox, radiobutton_pares, radiobutton_impares = widgets
            if checkbox.isChecked():
                if radiobutton_pares and radiobutton_impares:
                    if radiobutton_pares.isChecked():
                        dias_trabalho = 'pares'
                    elif radiobutton_impares.isChecked():
                        dias_trabalho = 'impares'
                    else:
                        QMessageBox.warning(
                            self,
                            "Aviso",
                            f"Selecione 'Pares' ou 'Ímpares' para o técnico {tecnico}."
                        )
                        return  # Impede de continuar sem selecionar
                    self.selected_tecnicos[tecnico] = dias_trabalho
                else:
                    self.selected_tecnicos[tecnico] = None  # Técnicos 5x2 não precisam dessa informação
        if not self.selected_tecnicos:
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos um técnico.")
        else:
            self.accept()

    def get_button_style(self):
        return """
            QPushButton {
                background-color: #007bff;
                color: white;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """

# Classe EmailSelectionDialog
class EmailSelectionDialog(QDialog):
    def __init__(self, tecnicos, periodo_inicio, periodo_fim):
        super().__init__()
        self.setWindowIcon(QIcon('JT.ico'))
        self.selected_tecnicos = []
        self.periodo_inicio = periodo_inicio
        self.periodo_fim = periodo_fim
        self.init_ui(tecnicos)

    def init_ui(self, tecnicos):
        self.setWindowTitle("Selecionar Técnicos e Período para Envio")
        self.resize(400, 500)
        layout = QVBoxLayout()

        # Lista de seleção de técnicos
        self.tecnico_list_widget = QListWidget()
        self.tecnico_list_widget.setSelectionMode(QAbstractItemView.MultiSelection)
        for tecnico in tecnicos:
            item = QListWidgetItem(tecnico)
            self.tecnico_list_widget.addItem(item)

        layout.addWidget(QLabel("Selecione os Técnicos:"))
        layout.addWidget(self.tecnico_list_widget)

        # Seleção do período
        periodo_layout = QHBoxLayout()
        self.data_inicio = QDateEdit()
        self.data_inicio.setCalendarPopup(True)
        self.data_inicio.setDate(self.periodo_inicio)
        self.data_fim = QDateEdit()
        self.data_fim.setCalendarPopup(True)
        self.data_fim.setDate(self.periodo_fim)
        periodo_layout.addWidget(QLabel("Data Início:"))
        periodo_layout.addWidget(self.data_inicio)
        periodo_layout.addWidget(QLabel("Data Fim:"))
        periodo_layout.addWidget(self.data_fim)

        layout.addWidget(QLabel("Selecione o Período:"))
        layout.addLayout(periodo_layout)

        # Botão Confirmar
        self.confirm_button = QPushButton("Confirmar")
        self.confirm_button.clicked.connect(self.confirm_selection)
        self.confirm_button.setStyleSheet(self.get_button_style())

        layout.addWidget(self.confirm_button)
        self.setLayout(layout)

    def confirm_selection(self):
        selected_items = self.tecnico_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos um técnico.")
            return

        self.selected_tecnicos = [item.text() for item in selected_items]
        self.periodo_inicio = self.data_inicio.date()
        self.periodo_fim = self.data_fim.date()
        if self.periodo_fim < self.periodo_inicio:
            QMessageBox.warning(self, "Erro", "Data fim não pode ser antes da data início.")
            return
        self.accept()

    def get_button_style(self):
        return """
            QPushButton {
                background-color: #007bff;
                color: white;
                border-radius: 5px;
                padding: 6px 12px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """

# Classe ScheduleForm
class ScheduleForm(QWidget):
    def __init__(self, planilha_path, periodo_inicio, periodo_fim):
        super().__init__()
        self.setWindowIcon(QIcon('JT.ico'))
        self.planilha_path = planilha_path
        self.periodo_inicio = periodo_inicio
        self.periodo_fim = periodo_fim
        self.editing_row = None  # Para rastrear se estamos editando uma linha
        self.enter_pressed_once = False  # Variável para controlar o primeiro clique de Enter
        self.enter_timer = QTimer(self)  # Timer para gerenciar o tempo entre dois Enters
        self.enter_timer.setInterval(500)  # Intervalo de 500ms entre os dois Enters
        self.enter_timer.timeout.connect(self.reset_enter)  # Conecta o timeout ao método reset_enter
        self.is_editing_entry = False  # Flag para controlar o modo de edição
        self.tecnicos_12x36_dias = {}  # Dicionário para armazenar os dias de trabalho dos técnicos 12x36

        # Carregar as escalas dos técnicos a partir do arquivo JSON
        try:
            with open('escalas_tecnicos.json', 'r', encoding='utf-8') as f:
                self.technician_schedules = json.load(f)
        except FileNotFoundError:
            QMessageBox.critical(self, "Erro", "Arquivo 'escalas_tecnicos.json' não encontrado.")
            sys.exit()
        except json.JSONDecodeError as e:
            QMessageBox.critical(self, "Erro", f"Erro ao decodificar o arquivo JSON: {e}")
            sys.exit()

        self.init_ui()

    def reset_enter(self):
        """
        Reseta o estado de "Enter pressionado" se o tempo limite for atingido.
        """
        self.enter_pressed_once = False
        self.enter_timer.stop()

    def handle_enter_press(self):
        """
        Lida com o pressionamento da tecla Enter. Se for o segundo pressionamento dentro do intervalo, executa a ação.
        """
        if self.enter_pressed_once:
            # Segundo pressionamento dentro do intervalo
            self.enter_pressed_once = False
            self.enter_timer.stop()
            # Executar ação de adicionar ou salvar
            self.add_entry()  # Tanto adiciona nova entrada quanto salva edição
        else:
            # Primeiro pressionamento
            self.enter_pressed_once = True
            self.enter_timer.start()

    def should_work(self, tecnico, date):
        """
        Verifica se o técnico deve trabalhar na data fornecida com base em sua escala.

        :param tecnico: Nome do técnico.
        :param date: Objeto QDate representando a data a ser verificada.
        :return: Booleano indicando se o técnico deve trabalhar nessa data.
        """
        tecnico_info = self.technician_schedules.get(tecnico, {})
        if not tecnico_info:
            return False  # Técnico não encontrado

        escala = tecnico_info.get('escala')

        if escala == '5X2':
            # Para escala 5X2, verificar se o dia da semana está na lista de dias de trabalho
            dia_semana = date.dayOfWeek()  # 1 (Segunda) a 7 (Domingo)
            return dia_semana in tecnico_info.get('dias_trabalho', [])

        elif escala == '12X36':
            # Para escala 12X36, utilizar a informação fornecida pelo usuário
            dias_trabalho = self.tecnicos_12x36_dias.get(tecnico)
            if dias_trabalho:
                dia_mes = date.day()
                if dias_trabalho == 'pares':
                    return dia_mes % 2 == 0
                elif dias_trabalho == 'impares':
                    return dia_mes % 2 != 0

        return False  # Escala não reconhecida ou dias_trabalho não especificado

    def does_on_call(self, tecnico):
        tecnico_info = self.technician_schedules.get(tecnico)
        if not tecnico_info:
            return False
        sobreaviso_info = tecnico_info.get('sobreaviso', {})
        return sobreaviso_info.get('faz_sobreaviso', False)

    def incluir_escala_semanal(self):
        tecnicos = list(self.technician_schedules.keys())
        dialog = TechnicianSelectionDialog(tecnicos, self.technician_schedules)
        if dialog.exec_() == QDialog.Accepted:
            selected_tecnicos = dialog.selected_tecnicos
            self.tecnicos_12x36_dias = {}  # Dicionário para armazenar os dias de trabalho dos técnicos 12x36
            for tecnico, dias_trabalho in selected_tecnicos.items():
                if self.technician_schedules[tecnico]['escala'] == '12X36':
                    self.tecnicos_12x36_dias[tecnico] = dias_trabalho
            for tecnico in selected_tecnicos:
                current_date = self.periodo_inicio
                while current_date <= self.periodo_fim:
                    # Verificar se o técnico deve trabalhar nesta data
                    if self.should_work(tecnico, current_date):
                        self.combo_box_tecnico.setCurrentText(tecnico)
                        self.date_time_edit_inicio.setDate(current_date)
                        self.update_fields_based_on_tecnico()
                        self.add_entry()
                    current_date = current_date.addDays(1)

    def add_entry(self):
        # Coleta os dados dos campos
        dia_semana = self.dia_semana.text()
        localizacao = self.combo_box_localizacao.currentText()
        unidade = self.combo_box_unidade.currentText()
        tecnico = self.combo_box_tecnico.currentText()
        turno = self.combo_box_turno.currentText()
        data_hora_inicio = self.date_time_edit_inicio.dateTime()
        data_hora_fim = self.date_time_edit_fim.dateTime()
        justificativa = self.justificativa.text()
        card_value = self.card_input.text()

        # Obter informações do técnico
        tecnico_info = self.technician_schedules.get(tecnico, {})
        escala = tecnico_info.get('escala', '')

        # Verificar se o técnico deve trabalhar nesta data apenas se a escala for 5X2 e a localização não for 'Sobreaviso'
        if escala == '5X2' and localizacao != 'Sobreaviso' and not self.should_work(tecnico, data_hora_inicio.date()):
            QMessageBox.warning(
                self,
                "Erro",
                f"{tecnico} não está programado para trabalhar em {data_hora_inicio.toString('dd/MM/yyyy')}."
            )
            return  # Interrompe a adição da entrada

        # Se for Sobreaviso, verificar se o técnico faz Sobreaviso
        if localizacao == 'Sobreaviso' and not self.does_on_call(tecnico):
            QMessageBox.warning(
                self,
                "Erro",
                f"{tecnico} não está configurado para Sobreaviso."
            )
            return  # Interrompe a adição da entrada

        # Converter datas para objetos datetime (mantém como datetime para o pandas)
        data_hora_inicio_dt = data_hora_inicio.toPyDateTime()
        data_hora_fim_dt = data_hora_fim.toPyDateTime()

        # Formatar as datas no formato desejado
        data_hora_inicio_str = data_hora_inicio_dt.strftime("%d/%m/%Y %H:%M:%S")
        data_hora_fim_str = data_hora_fim_dt.strftime("%d/%m/%Y %H:%M:%S")

        # Lista com os valores dos campos formatados
        fields = [
            dia_semana,
            localizacao,
            unidade,
            tecnico,
            escala,  # Adicionado aqui
            turno,
            data_hora_inicio_str,
            data_hora_fim_str,
            justificativa,
            card_value
        ]

        if self.editing_row is not None:
            # Se estivermos editando uma linha, atualiza os valores
            for column, value in enumerate(fields):
                self.table_widget.setItem(self.editing_row, column, QTableWidgetItem(value))
            self.original_data[self.editing_row] = fields  # Atualiza os dados originais
            self.editing_row = None
            self.add_button.setText(" Adicionar")
            self.add_button.setIcon(QIcon("icons/add.png"))
        else:
            # Adiciona uma nova linha à tabela
            row_position = self.table_widget.rowCount()
            self.table_widget.insertRow(row_position)
            for column, value in enumerate(fields):
                self.table_widget.setItem(row_position, column, QTableWidgetItem(value))
            self.original_data.append(fields)  # Armazena os dados originais

            # Se for Sobreaviso em um domingo, adicionar folga na segunda-feira
            if localizacao == 'Sobreaviso':
                day_of_week = data_hora_inicio.date().dayOfWeek()  # 7 para domingo
                if day_of_week == 7:
                    tecnico_info = self.technician_schedules.get(tecnico)
                    if tecnico_info:
                        # Adicionar folga na segunda-feira
                        folga_date = data_hora_inicio.addDays(1)
                        horario_inicio = tecnico_info['horario_inicio']
                        horario_fim = tecnico_info['horario_fim']
                        folga_inicio = QDateTime(folga_date.date(), QTime.fromString(horario_inicio, "HH:mm"))
                        folga_fim = QDateTime(folga_date.date(), QTime.fromString(horario_fim, "HH:mm"))

                        # Atualizar dia da semana sem data
                        folga_dia_semana = self.get_dia_semana_text(folga_inicio)

                        folga_fields = [
                            folga_dia_semana,
                            "Folga",
                            "",
                            tecnico,
                            escala,  # Adicionado aqui
                            turno,
                            folga_inicio.toString("dd/MM/yyyy HH:mm:ss"),
                            folga_fim.toString("dd/MM/yyyy HH:mm:ss"),
                            "Folga após sobreaviso",
                            ""  # Campo CARD vazio
                        ]

                        row_position = self.table_widget.rowCount()
                        self.table_widget.insertRow(row_position)
                        for column, value in enumerate(folga_fields):
                            self.table_widget.setItem(row_position, column, QTableWidgetItem(value))
                        self.original_data.append(folga_fields)

        # Limpa os campos após a adição
        self.clear_fields()

    def init_ui(self):
        # Layout principal vertical
        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.setSpacing(0)  # Remove o espaçamento vertical entre os widgets

        # Label com o período selecionado
        self.periodo_label = ClickableLabel(
            f"Período Selecionado: {self.periodo_inicio.toString('dd/MM/yyyy')} a {self.periodo_fim.toString('dd/MM/yyyy')}"
        )
        self.periodo_label.setAlignment(Qt.AlignCenter)
        self.periodo_label.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                margin: 10px;
                color: #343a40;
            }
            QLabel:hover {
                color: #007bff;
                text-decoration: underline;
                cursor: pointer;
            }
        """)
        self.periodo_label.clicked.connect(self.change_period)
        main_layout.addWidget(self.periodo_label)

        # Lista de títulos e opções
        self.labels = [
            'DIA DA SEMANA', 'LOCALIZAÇÃO', 'UNIDADE', 'TÉCNICO', 'ESCALA', 'TURNO',
            'DATA/HORA INICIO', 'DATA/HORA FIM', 'JUSTIFICATIVA', 'CARD'
        ]
        localizacao_options = ["", "Folga", "Férias", "Sobreaviso", "Unidade", "Escritório", "Home"]

        # Definir as listas com os valores especificados
        unidades = unidades = [ "AMA Zaio", "CRST Freguesia do Ó", "CRST Lapa", "CRST Leste", "CRST Mooca", "CRST Santo Amaro", "CRST Sé", "HD Brasilandia", "HM Alipio", "HM Benedicto", "HM Brasilandia", "HM Brigadeiro", "HM Cachoeirinha", "HM Campo Limpo", "H Cantareira", "HM Capela Do Socorro", "HM Hungria", "HM Ignacio", "HM Mario Degni", "HM Saboya",
"HM Sorocabana", "HM Tatuape", "HM Tide", "HM Waldomiro", "HM Zaio", "PA Sao Mateus", "PSM Balneario Sao Jose", "PSM Lapa", "UPA Dona Maria Antonieta", "UPA Elisa Maria", "UPA Parelheiros", "UPA Pedreira", "UPA Peri", "UPA 21 de Junho", "UPA Santo Amaro",
"UPA Parque Doroteia", "UPA 26 de Agosto", "UPA Campo Limpo", "UPA Tiradentes", "UPA City Jaragua", "UPA Ermelino Matarazzo", "UPA Carrao", "UPA Rio Pequeno", "UPA Jabaquara", "UPA Jacana", "UPA Jardim Angela", "UPA Julio Tupy", "UPA Mooca", "UPA Perus", "UPA Pirituba", "UPA Tatuape", "UPA Tito Lopes", "UPA Vera Cruz", "UPA Vergueiro", "UPA Vila Mariana",
"UPA Vila Santa Catarina", "CAPS AD II Cachoeirinha", "CAPS AD II Cangaiba", "CAPS AD II Cidade Ademar", "CAPS AD II Ermelino Matarazzo", "CAPS AD II Guaianases", "CAPS AD II Jabaquara", "CAPS AD II Jardim Nelia", "CAPS AD II Mooca", "CAPS AD II Pinheiros", "CAPS AD II Sacoma", "CAPS AD II Santo Amaro", "CAPS AD II Sapopemba", "CAPS AD II Vila Madalena Prosam", "CAPS AD II Vila Mariana", "CAPS AD III Armenia", "CAPS AD III Boracea", "CAPS AD III Butanta", "CAPS AD III Campo Limpo", "CAPS AD III Capela Do Socorro",
"CAPS AD III Centro", "CAPS AD III Complexo Prates", "CAPS AD III Freguesia Do O Brasilandia", "CAPS AD III Grajau", "CAPS AD III Heliopolis", "CAPS AD III Itaquera", "CAPS AD III Jardim Angela", "CAPS AD III Jardim Sao Luiz", "CAPS AD III Leopoldina", "CAPS AD III Paraisopolis", "CAPS AD III Penha", "CAPS AD III Pirituba Casa Azul", "CAPS AD III Santana", "CAPS AD III Sao Mateus Liberdade De Escolha", "CAPS AD III Sao Miguel", "CAPS AD IV Redencao", "CAPS Adulto II Aricanduva Formosa", "CAPS Adulto II Butanta", "CAPS Adulto II Casa Verde", "CAPS Adulto II Cidade Ademar",
"CAPS Adulto II Cidade Tiradentes", "CAPS Adulto II Ermelino Matarazzo", "CAPS Adulto II Guaianases Artur Bispo Do Rosario", "CAPS Adulto II Itaim Paulista", "CAPS Adulto II Itaquera", "CAPS Adulto II Jardim Lidia", "CAPS Adulto II Jabaquara", "CAPS Adulto II Jacana Dr Leonidio Galvao Dos Santos", "CAPS Adulto II Perdizes Manoel Munhoz", "CAPS Adulto II Perus", "CAPS Adulto II Sao Miguel", "CAPS Adulto II V Monumento", "CAPS Adulto II Vila Prudente", "CAPS Adulto III Capela Do Socorro", "CAPS Adulto III Freguesia Do O Brasilandia", "CAPS Adulto III Grajau", "CAPS Adulto III Itaim Bibi", "CAPS Adulto III Jardim Sao Luiz", "CAPS Adulto III Lapa", "CAPS Adulto III Largo 13",
"CAPS Adulto III M Boi Mirim", "CAPS Adulto III Mandaqui", "CAPS Adulto III Mooca", "CAPS Adulto III Paraisopolis", "CAPS Adulto III Parelheiros", "CAPS Adulto III Pirituba Jaragua", "CAPS Adulto III Sao Mateus", "CAPS Adulto III Sapopemba", "CAPS Adulto III Se", "CAPS Adulto III Vila Matilde", "CAPS IJ II Pirituba Jaragua", "CAPS IJ II Vila Mariana Quixote", "CAPS IJ II Butanta", "CAPS IJ II Campo Limpo", "CAPS IJ II Capela Do Socorro Piracao", "CAPS IJ II Casa Verde Nise Da Silveira", "CAPS IJ II Cidade Ademar", "CAPS IJ II Cidade Lider", "CAPS IJ II Cidade Tiradentes", "CAPS IJ II Ermelino Matarazzo",
"CAPS IJ II Freguesia Do O Brasilandia", "CAPS IJ II Guaianases Coloridamente", "CAPS IJ II Ipiranga", "CAPS IJ II Itaim Paulista", "CAPS IJ II Itaquera", "CAPS IJ II Jabaquara Casinha", "CAPS IJ II Lapa", "CAPS IJ II M Boi Mirim", "CAPS IJ II Mooca", "CAPS IJ II Parelheiros Aquarela", "CAPS IJ II Perus", "CAPS IJ II Santo Amaro", "CAPS IJ II Sao Mateus", "CAPS IJ II Sapopemba", "CAPS IJ II Vila Maria Vila Guilherme", "CAPS IJ II Vila Prudente", "CAPS IJ III Aricanduva", "CAPS IJ III Cidade Dutra", "CAPS IJ III Heliopolis", "CAPS IJ III Jardim Sao Luiz",
"CAPS IJ III Penha", "CAPS IJ III Santana", "CAPS IJ III Sao Miguel"
]
        tecnicos = list(self.technician_schedules.keys())
        turnos = ["Diurno", "Noturno"]

        # Criar os widgets de entrada
        self.dia_semana = QLabel("-")
        self.dia_semana.setAlignment(Qt.AlignCenter)
        self.dia_semana.setStyleSheet("""
            QLabel {
                border: 1px solid #ced4da;
                padding: 5px;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
                background-color: #fff;
                border-radius: 4px;
            }
        """)
        self.combo_box_localizacao = QComboBox()
        self.combo_box_localizacao.addItems(localizacao_options)
        self.combo_box_localizacao.setCurrentIndex(0)
        self.combo_box_localizacao.currentIndexChanged.connect(self.handle_localizacao_change)
        self.combo_box_localizacao.setStyleSheet("""
            QComboBox {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 5px;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
                background-color: #fff;
            }
            QComboBox QAbstractItemView {
                background-color: #fff;
                selection-background-color: #007bff;
                selection-color: #fff;
            }
        """)

        # Substituir sample_items pelas listas definidas
        self.combo_box_unidade = FilteredComboBox(unidades, parent=self)
        self.combo_box_tecnico = FilteredComboBox(tecnicos, parent=self)
        self.combo_box_turno = FilteredComboBox(turnos, parent=self)
        self.combo_box_tecnico.lineEdit().editingFinished.connect(self.update_fields_based_on_tecnico)
        self.combo_box_unidade.lineEdit().editingFinished.connect(self.update_fields_based_on_tecnico)

        # Criar QLabel para 'Escala'
        self.escala_label = QLabel("-")
        self.escala_label.setAlignment(Qt.AlignCenter)
        self.escala_label.setStyleSheet("""
            QLabel {
                border: 1px solid #ced4da;
                padding: 5px;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
                background-color: #fff;
                border-radius: 4px;
            }
        """)

        # Inicializa date_time_edit_inicio e date_time_edit_fim antes de conectar sinais
        self.date_time_edit_inicio = QDateTimeEdit()
        self.date_time_edit_fim = QDateTimeEdit()

        # Configurar date_time_edit_inicio
        self.date_time_edit_inicio.setCalendarPopup(True)
        self.date_time_edit_inicio.setDisplayFormat("dd/MM/yyyy HH:mm:ss")
        self.date_time_edit_inicio.setStyleSheet("""
            QDateTimeEdit {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 5px;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
                background-color: #fff;
            }
            QDateTimeEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #ced4da;
            }
        """)
        # Configurar limites de data/hora com base no período selecionado
        self.date_time_edit_inicio.setMinimumDateTime(QDateTime(self.periodo_inicio, QTime(0, 0, 0)))
        self.date_time_edit_inicio.setMaximumDateTime(QDateTime(self.periodo_fim, QTime(23, 59, 59)))
        self.date_time_edit_inicio.setDateTime(QDateTime(self.periodo_inicio, QTime.currentTime()))
        self.date_time_edit_inicio.setEnabled(True)  # Assegurar que está habilitado

        # Configurar date_time_edit_fim
        self.date_time_edit_fim.setCalendarPopup(True)
        self.date_time_edit_fim.setDisplayFormat("dd/MM/yyyy HH:mm:ss")
        self.date_time_edit_fim.setStyleSheet("""
            QDateTimeEdit {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 5px;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
                background-color: #fff;
            }
            QDateTimeEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #ced4da;
            }
        """)
        # Configurar limites de data/hora com base no período selecionado
        self.date_time_edit_fim.setMinimumDate(self.periodo_inicio)
        self.date_time_edit_fim.setMaximumDate(self.periodo_fim.addDays(1))  # Ajuste realizado aqui
        self.date_time_edit_fim.setDateTime(QDateTime(self.periodo_inicio, QTime.currentTime()))
        self.date_time_edit_fim.setEnabled(True)  # Assegurar que está habilitado

        # Conectar sinais
        # Remover a conexão que atualiza os campos de data/hora automaticamente
        # self.date_time_edit_inicio.dateTimeChanged.connect(self.on_datetime_changed)

        # Conectar para atualizar apenas o dia da semana ao modificar o date_time_edit_inicio
        self.date_time_edit_inicio.dateTimeChanged.connect(self.update_dia_semana_from_datetime)

        self.justificativa = QLineEdit()
        self.justificativa.setStyleSheet("""
            QLineEdit {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 5px;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
                background-color: #fff;
            }
        """)

        self.card_input = QLineEdit()
        self.card_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 5px;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
                background-color: #fff;
            }
        """)

        # Lista dos widgets de entrada
        self.input_widgets = [
            self.dia_semana,
            self.combo_box_localizacao,
            self.combo_box_unidade,
            self.combo_box_tecnico,
            self.escala_label,  # Adicionado aqui
            self.combo_box_turno,
            self.date_time_edit_inicio,
            self.date_time_edit_fim,
            self.justificativa,
            self.card_input
        ]

        # Criação da tabela do formulário
        self.form_table = QTableWidget()
        self.form_table.setRowCount(1)
        self.form_table.setColumnCount(len(self.labels))
        self.form_table.setHorizontalHeaderLabels(self.labels)
        self.form_table.verticalHeader().setVisible(False)
        self.form_table.horizontalHeader().setStretchLastSection(True)
        self.form_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.form_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.form_table.setSelectionMode(QTableWidget.NoSelection)
        self.form_table.setFocusPolicy(Qt.NoFocus)
        self.form_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # Não expandir verticalmente
        self.form_table.setMaximumHeight(100)  # Define uma altura máxima apropriada

        # Estilo com fundo degradê nos cabeçalhos
        self.form_table.setStyleSheet("""
            QTableWidget {
                border: none;
            }
            QHeaderView::section {
                background-color: qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:0,
                    stop:0 #007bff, stop:1 #6610f2
                );
                color: white;
                font-weight: bold;
                border: none;
                font-size: 15px;
                font-family: 'Segoe UI', sans-serif;
                height: 35px;
            }
        """)

        # Inserir os widgets de entrada nas células da tabela do formulário
        for col, widget in enumerate(self.input_widgets):
            self.form_table.setCellWidget(0, col, widget)

        main_layout.addWidget(self.form_table)

        # Habilitar a ordenação nos cabeçalhos do formulário
        self.form_table.horizontalHeader().setSectionsClickable(True)
        self.form_table.horizontalHeader().setSortIndicatorShown(True)
        self.form_table.horizontalHeader().sectionClicked.connect(self.handle_header_click)

        # Criação da tabela para exibir as entradas adicionadas
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(len(self.labels))
        # Remover os cabeçalhos das colunas
        self.table_widget.horizontalHeader().setVisible(False)
        self.table_widget.verticalHeader().setVisible(False)
        self.table_widget.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_widget.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_widget.setStyleSheet("""
            QTableWidget {
                border: none;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
            }
            QTableWidget::item {
                border-bottom: 1px solid #dee2e6;
            }
            QTableWidget::item:selected {
                background-color: #007bff;
                color: #fff;
            }
        """)
        # Mostrar linhas divisórias
        self.table_widget.setShowGrid(False)
        # Ajustar o redimensionamento das colunas
        self.table_widget.horizontalHeader().setStretchLastSection(True)
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        main_layout.addWidget(self.table_widget)

        # Inicializar estados de ordenação
        self.sort_states = {
            2: False,  # 'UNIDADE'
            3: False,  # 'TÉCNICO'
            6: False   # 'DATA/HORA INICIO' (índice atualizado)
        }

        # Armazenar dados originais
        self.original_data = []

        # Adicionar botões de EDITAR, EXCLUIR e FINALIZAR ESCALA (sem o botão "Adicionar")
        buttons_layout = QHBoxLayout()

        # Layout para os botões à esquerda
        left_buttons_layout = QHBoxLayout()

        # Botão Consultar Escala (alinhado à esquerda)
        self.consult_button = QPushButton(" Consultar Escala")
        self.consult_button.setFixedSize(180, 40)
        self.consult_button.setIcon(QIcon("icons/consult.png"))
        self.consult_button.setStyleSheet(self.get_primary_button_style())
        self.consult_button.clicked.connect(self.consultar_escala)
        left_buttons_layout.addWidget(self.consult_button)

        # Botão Incluir Escala Semanal (ao lado do Consultar Escala)
        self.incluir_escala_button = QPushButton(" Incluir Escala Pré-Preenchida")
        self.incluir_escala_button.setFixedSize(220, 40)
        self.incluir_escala_button.setIcon(QIcon("icons/add_schedule.png"))
        self.incluir_escala_button.setStyleSheet(self.get_success_button_style())
        self.incluir_escala_button.clicked.connect(self.incluir_escala_semanal)
        left_buttons_layout.addWidget(self.incluir_escala_button)

        # Botão para Enviar E-mails
        self.send_email_button = QPushButton(" Enviar Escala")
        self.send_email_button.setFixedSize(180, 40)
        self.send_email_button.setIcon(QIcon("icons/email.png"))
        self.send_email_button.setStyleSheet(self.get_primary_button_style())
        self.send_email_button.clicked.connect(self.send_emails)
        left_buttons_layout.addWidget(self.send_email_button)

        buttons_layout.addLayout(left_buttons_layout)

        buttons_layout.addStretch()  # Espaço entre os botões

        # Layout para os botões à direita
        right_buttons_layout = QHBoxLayout()
        right_buttons_layout.setSpacing(10)

        # Botão "Adicionar" (visível)
        self.add_button = QPushButton(" Adicionar")
        self.add_button.setFixedSize(140, 40)
        self.add_button.setIcon(QIcon("icons/add.png"))
        self.add_button.setStyleSheet(self.get_success_button_style())
        self.add_button.clicked.connect(self.add_entry)
        # Não esconder o botão para permitir salvar com Enter

        right_buttons_layout.addWidget(self.add_button)

        self.edit_button = QPushButton(" Editar")
        self.edit_button.setFixedSize(140, 40)
        self.edit_button.setIcon(QIcon("icons/edit.png"))
        self.edit_button.setStyleSheet(self.get_warning_button_style())
        self.edit_button.clicked.connect(self.edit_entry)
        right_buttons_layout.addWidget(self.edit_button)

        self.delete_button = QPushButton(" Excluir")
        self.delete_button.setFixedSize(140, 40)
        self.delete_button.setIcon(QIcon("icons/delete.png"))
        self.delete_button.setStyleSheet(self.get_danger_button_style())
        self.delete_button.clicked.connect(self.delete_entry)
        right_buttons_layout.addWidget(self.delete_button)

        self.finalize_button = QPushButton(" Finalizar Escala")
        self.finalize_button.setFixedSize(180, 40)
        self.finalize_button.setIcon(QIcon("icons/save.png"))
        self.finalize_button.setStyleSheet(self.get_primary_button_style())
        self.finalize_button.clicked.connect(self.finalize_schedule)
        right_buttons_layout.addWidget(self.finalize_button)

        buttons_layout.addLayout(right_buttons_layout)
        main_layout.addLayout(buttons_layout)

        # Estilizar o layout principal
        self.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
            }
            QPushButton {
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
            }
        """)

        self.setLayout(main_layout)
        self.setWindowTitle('Saints V')
        self.resize(1200, 600)
        self.show()

    def get_primary_button_style(self):
        return """
            QPushButton {
                background-color: #007bff;
                color: white;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """

    def get_success_button_style(self):
        return """
            QPushButton {
                background-color: #28a745;
                color: white;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #1e7e34;
            }
        """

    def get_warning_button_style(self):
        return """
            QPushButton {
                background-color: #ffc107;
                color: #212529;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #e0a800;
            }
        """

    def get_danger_button_style(self):
        return """
            QPushButton {
                background-color: #dc3545;
                color: white;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #bd2130;
            }
        """

    def change_period(self):
        """
        Permite ao usuário alterar o período selecionado.
        """
        # Exibir diálogo para selecionar o período
        periodo_dialog = PeriodoConsultaDialog()
        if periodo_dialog.exec_() == QDialog.Accepted:
            self.periodo_inicio = periodo_dialog.periodo_inicio
            self.periodo_fim = periodo_dialog.periodo_fim
            # Atualizar o label
            self.periodo_label.setText(
                f"Período Selecionado: {self.periodo_inicio.toString('dd/MM/yyyy')} a {self.periodo_fim.toString('dd/MM/yyyy')}"
            )
            # Atualizar os campos que dependem do período
            self.date_time_edit_inicio.setMinimumDate(self.periodo_inicio)
            self.date_time_edit_inicio.setMaximumDate(self.periodo_fim)
            self.date_time_edit_fim.setMinimumDate(self.periodo_inicio)
            self.date_time_edit_fim.setMaximumDate(self.periodo_fim.addDays(1))  # Ajuste realizado aqui
        else:
            # Usuário cancelou
            pass

    def clear_fields(self):
        # Limpa ou redefine os campos do formulário
        self.combo_box_localizacao.setCurrentIndex(0)
        self.combo_box_unidade.setCurrentIndex(0)
        # self.combo_box_tecnico.setCurrentIndex(0)  # Opcional: não limpar o técnico
        self.combo_box_turno.setCurrentIndex(0)
        self.date_time_edit_inicio.setDateTime(QDateTime(self.periodo_inicio, QTime.currentTime()))
        self.date_time_edit_fim.setDateTime(QDateTime(self.periodo_inicio, QTime.currentTime()))
        self.justificativa.clear()
        self.card_input.clear()
        self.dia_semana.setText("-")
        self.escala_label.setText("-")  # Adicionado aqui
        self.add_button.setText(" Adicionar")
        self.add_button.setIcon(QIcon("icons/add.png"))
        self.editing_row = None

    def edit_entry(self):
        selected_rows = self.table_widget.selectionModel().selectedRows()
        if selected_rows:
            selected_row = selected_rows[0].row()
            self.editing_row = selected_row

            self.is_editing_entry = True  # Inicia o modo de edição

            # Desconectar sinais que podem interferir
            try:
                self.combo_box_tecnico.lineEdit().editingFinished.disconnect(self.update_fields_based_on_tecnico)
                self.combo_box_unidade.lineEdit().editingFinished.disconnect(self.update_fields_based_on_tecnico)
            except TypeError:
                pass  # Sinal já desconectado

            # Pega os dados da linha selecionada
            dia_semana = self.table_widget.item(selected_row, 0).text()
            localizacao = self.table_widget.item(selected_row, 1).text()
            unidade = self.table_widget.item(selected_row, 2).text()
            tecnico = self.table_widget.item(selected_row, 3).text()
            escala = self.table_widget.item(selected_row, 4).text()  # 'ESCALA'
            turno = self.table_widget.item(selected_row, 5).text()
            data_hora_inicio = self.table_widget.item(selected_row, 6).text()
            data_hora_fim = self.table_widget.item(selected_row, 7).text()
            justificativa = self.table_widget.item(selected_row, 8).text()
            card_value = self.table_widget.item(selected_row, 9).text()

            # Preenche os campos com os dados para edição
            self.dia_semana.setText(dia_semana)
            index_localizacao = self.combo_box_localizacao.findText(localizacao)
            if index_localizacao >= 0:
                self.combo_box_localizacao.setCurrentIndex(index_localizacao)
            self.combo_box_unidade.setCurrentText(unidade)
            self.combo_box_tecnico.setCurrentText(tecnico)
            self.escala_label.setText(escala)  # Atualiza a escala
            self.combo_box_turno.setCurrentText(turno)

            # Parse das datas com o formato correto
            date_time_inicio = QDateTime.fromString(data_hora_inicio, "dd/MM/yyyy HH:mm:ss")
            date_time_fim = QDateTime.fromString(data_hora_fim, "dd/MM/yyyy HH:mm:ss")
            if date_time_inicio.isValid():
                self.date_time_edit_inicio.setDateTime(date_time_inicio)
            else:
                self.date_time_edit_inicio.setDateTime(QDateTime(self.periodo_inicio, QTime.currentTime()))
            if date_time_fim.isValid():
                self.date_time_edit_fim.setDateTime(date_time_fim)
            else:
                self.date_time_edit_fim.setDateTime(QDateTime(self.periodo_inicio, QTime.currentTime()))

            self.justificativa.setText(justificativa)
            self.card_input.setText(card_value)

            self.add_button.setText(" Gravar")
            self.add_button.setIcon(QIcon("icons/save.png"))

            # Reconectar o sinal
            self.combo_box_tecnico.lineEdit().editingFinished.connect(self.update_fields_based_on_tecnico)
            self.combo_box_unidade.lineEdit().editingFinished.connect(self.update_fields_based_on_tecnico)

            self.is_editing_entry = False  # Encerra o modo de edição
        else:
            # Nenhuma linha selecionada
            QMessageBox.warning(self, "Aviso", "Nenhuma linha selecionada para editar.")

    def delete_entry(self):
        selected_rows = self.table_widget.selectionModel().selectedRows()
        if selected_rows:
            selected_row = selected_rows[0].row()
            self.table_widget.removeRow(selected_row)
            del self.original_data[selected_row]  # Remove dos dados originais
            self.clear_fields()
        else:
            # Nenhuma linha selecionada
            QMessageBox.warning(self, "Aviso", "Nenhuma linha selecionada para excluir.")

    def handle_header_click(self, logicalIndex):
        if logicalIndex in self.sort_states:
            # Alterna o estado de ordenação
            self.sort_states[logicalIndex] = not self.sort_states[logicalIndex]
            order = Qt.AscendingOrder if self.sort_states[logicalIndex] else Qt.DescendingOrder
            # Define o indicador de ordenação
            self.form_table.horizontalHeader().setSortIndicator(logicalIndex, order)
            # Ordena a tabela
            self.sort_table(logicalIndex, order)
        else:
            # Para outras colunas, não faz nada
            pass

    def sort_table(self, column, order):
        # Copia os dados atuais da tabela
        data = []
        for row in range(self.table_widget.rowCount()):
            row_data = []
            for col in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row, col)
                if item:
                    row_data.append(item.text())
                else:
                    row_data.append('')

            data.append(row_data)

        # Ordena os dados com base na coluna
        reverse = (order == Qt.DescendingOrder)
        if column == 6:  # DATA/HORA INICIO (índice atualizado)
            data.sort(key=lambda x: QDateTime.fromString(x[column], "dd/MM/yyyy HH:mm:ss").toPyDateTime(), reverse=reverse)
        else:
            data.sort(key=lambda x: x[column], reverse=reverse)
        # Atualiza a tabela com os dados ordenados
        self.table_widget.setRowCount(0)
        for row_data in data:
            row_position = self.table_widget.rowCount()
            self.table_widget.insertRow(row_position)
            for column_index, value in enumerate(row_data):
                self.table_widget.setItem(row_position, column_index, QTableWidgetItem(value))
        # Atualiza os dados originais com a nova ordem
        self.original_data = data

    def finalize_schedule(self):
        if not self.original_data:
            # Nenhuma entrada para salvar
            QMessageBox.information(self, "Aviso", "Não há entradas para salvar.")
            return

        # Ler a planilha existente ou criar uma nova
        try:
            # Tentar ler a planilha existente
            df_existing = pd.read_excel(self.planilha_path)
            df_existing.columns = df_existing.columns.str.upper()  # Ajuste aqui
        except FileNotFoundError:
            # Se não existir, criar um DataFrame vazio
            df_existing = pd.DataFrame()

        # Garantir que a coluna 'SEQ' seja do tipo inteiro
        if not df_existing.empty and 'SEQ' in df_existing.columns:
            try:
                df_existing['SEQ'] = df_existing['SEQ'].astype(int)
            except ValueError:
                # Caso haja valores não inteiros, trate-os conforme necessário
                df_existing['SEQ'] = pd.to_numeric(df_existing['SEQ'], errors='coerce').fillna(0).astype(int)
        elif not df_existing.empty and 'SEQ' not in df_existing.columns:
            df_existing.insert(0, 'SEQ', range(1, len(df_existing) + 1))

        # Criar um DataFrame com os dados atuais
        df_new = pd.DataFrame(self.original_data, columns=self.labels)

        # Adicionar a coluna SEQ corretamente
        if not df_existing.empty and 'SEQ' in df_existing.columns:
            max_seq = df_existing['SEQ'].max()
            if pd.isna(max_seq):
                max_seq = 0
            else:
                max_seq = int(max_seq)  # Converter para inteiro
        else:
            max_seq = 0

        df_new.insert(0, 'SEQ', range(max_seq + 1, max_seq + len(df_new) + 1))

        # Garantir que as colunas de data/hora estejam no formato correto
        for col in ['DATA/HORA INICIO', 'DATA/HORA FIM']:
            df_new[col] = pd.to_datetime(df_new[col], format="%d/%m/%Y %H:%M:%S", errors='coerce')

        # Salvar na planilha usando pandas
        df_final = pd.concat([df_existing, df_new], ignore_index=True)
        df_final.to_excel(self.planilha_path, index=False, engine='openpyxl')

        # Abrir o arquivo Excel com openpyxl para aplicar formatação
        wb = load_workbook(self.planilha_path)
        ws = wb.active  # Considerando que os dados estão na primeira planilha

        # Encontrar os índices das colunas 'DATA/HORA INICIO' e 'DATA/HORA FIM'
        columns = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}
        inicio_col = columns.get('DATA/HORA INICIO')
        fim_col = columns.get('DATA/HORA FIM')

        # Definir o formato desejado
        date_format = 'dd/mm/yyyy hh:mm:ss'  # Correção aqui

        # Aplicar o formato na coluna 'DATA/HORA INICIO'
        if inicio_col:
            for cell in ws.iter_cols(min_col=inicio_col, max_col=inicio_col, min_row=2):
                for c in cell:
                    c.number_format = date_format

        # Aplicar o formato na coluna 'DATA/HORA FIM'
        if fim_col:
            for cell in ws.iter_cols(min_col=fim_col, max_col=fim_col, min_row=2):
                for c in cell:
                    c.number_format = date_format

        # Salvar as alterações no Excel
        wb.save(self.planilha_path)

        # Limpar a tabela e os dados originais
        self.table_widget.setRowCount(0)
        self.original_data.clear()

        # Mensagem de confirmação
        QMessageBox.information(self, "Sucesso", "Escala finalizada e salva com sucesso!")

    def consultar_escala(self, periodo_inicio=None, periodo_fim=None):
        if periodo_inicio is None or periodo_fim is None:
            # Exibir diálogo para selecionar o período
            periodo_dialog = PeriodoConsultaDialog()
            if periodo_dialog.exec_() == QDialog.Accepted:
                periodo_inicio = periodo_dialog.periodo_inicio
                periodo_fim = periodo_dialog.periodo_fim
            else:
                return  # Usuário cancelou

        # Ler a planilha existente
        try:
            df_existing = pd.read_excel(self.planilha_path)
            df_existing.columns = df_existing.columns.str.upper()  # Ajuste aqui
        except FileNotFoundError:
            QMessageBox.warning(self, "Erro", "A planilha selecionada não foi encontrada.")
            return

        # Verificar se há dados
        if df_existing.empty:
            QMessageBox.information(self, "Aviso", "Não há dados na planilha para consultar.")
            return

        # Converter a coluna 'DATA/HORA INICIO' para datetime
        df_existing['DATA/HORA INICIO'] = pd.to_datetime(
            df_existing['DATA/HORA INICIO'], format="%d/%m/%Y %H:%M:%S", errors='coerce'
        )

        # Filtrar pelo período selecionado
        start_date = QDateTime(periodo_inicio, QTime(0, 0)).toPyDateTime()
        end_date = QDateTime(periodo_fim, QTime(23, 59, 59)).toPyDateTime()
        df_filtered = df_existing[
            (df_existing['DATA/HORA INICIO'] >= start_date) &
            (df_existing['DATA/HORA INICIO'] <= end_date)
        ]

        if df_filtered.empty:
            QMessageBox.information(self, "Aviso", "Não há registros no período selecionado.")
            return

        # Abrir a tela de consulta
        self.consulta_dialog = ConsultaEscalaDialog(
            df_filtered, self.planilha_path, df_existing, periodo_inicio, periodo_fim, ['SEQ'] + self.labels
        )
        self.consulta_dialog.show()  # Usar show() em vez de exec_() para manter as janelas abertas simultaneamente

    def update_fields_based_on_tecnico(self):
        if self.is_editing_entry:
            return  # Não atualizar durante a edição

        tecnico_nome = self.combo_box_tecnico.currentText()
        tecnico_info = self.technician_schedules.get(tecnico_nome)
        if not tecnico_info:
            self.escala_label.setText('-')
            return  # Não faz nada se o técnico não for encontrado

        self.escala_label.setText(tecnico_info.get('escala', '-'))

        selected_localizacao = self.combo_box_localizacao.currentText()
        unidade_preenchida = bool(self.combo_box_unidade.currentText())

        # Obter a data/hora selecionada no date_time_edit_inicio
        selected_datetime = self.date_time_edit_inicio.dateTime()
        selected_date = selected_datetime.date()

        if selected_localizacao == 'Sobreaviso' and self.does_on_call(tecnico_nome):
            # Definir horários de Sobreaviso
            sobreaviso_info = tecnico_info.get('sobreaviso', {})
            if unidade_preenchida:
                horario_inicio = sobreaviso_info.get('horario_com_unidade', {}).get('inicio', tecnico_info['horario_inicio'])
                horario_fim = sobreaviso_info.get('horario_com_unidade', {}).get('fim', tecnico_info['horario_fim'])
            else:
                horario_inicio = sobreaviso_info.get('horario_sem_unidade', {}).get('inicio', tecnico_info['horario_inicio'])
                horario_fim = sobreaviso_info.get('horario_sem_unidade', {}).get('fim', tecnico_info['horario_fim'])
        else:
            # Usar horários normais
            horario_inicio = tecnico_info['horario_inicio']
            horario_fim = tecnico_info['horario_fim']

        # Montar datetime inicio e fim
        datetime_inicio_str = f"{selected_date.toString('dd/MM/yyyy')} {horario_inicio}"
        datetime_fim_str = f"{selected_date.toString('dd/MM/yyyy')} {horario_fim}"
        datetime_inicio = QDateTime.fromString(datetime_inicio_str, "dd/MM/yyyy HH:mm")
        datetime_fim = QDateTime.fromString(datetime_fim_str, "dd/MM/yyyy HH:mm")

        if datetime_fim <= datetime_inicio:
            datetime_fim = datetime_fim.addDays(1)

        self.date_time_edit_inicio.setDateTime(datetime_inicio)
        self.date_time_edit_fim.setDateTime(datetime_fim)

        self.update_dia_semana(datetime_inicio)

        hour_inicio = datetime_inicio.time().hour()
        if 6 <= hour_inicio < 18:
            self.combo_box_turno.setCurrentText("Diurno")
        else:
            self.combo_box_turno.setCurrentText("Noturno")

    def update_dia_semana_from_datetime(self):
        selected_datetime = self.date_time_edit_inicio.dateTime()
        self.update_dia_semana(selected_datetime)

    def get_dia_semana_text(self, selected_datetime):
        # Mapeamento dos dias da semana
        selected_date = selected_datetime.date()
        dias_semana = [
            "Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira",
            "Sexta-feira", "Sábado", "Domingo"
        ]
        dia_semana_idx = selected_date.dayOfWeek() - 1  # Retorna 1 para segunda-feira
        dia_semana_texto = dias_semana[dia_semana_idx]
        # Retornar apenas o dia da semana
        return dia_semana_texto

    def update_dia_semana(self, selected_datetime):
        dia_semana_sem_data = self.get_dia_semana_text(selected_datetime)
        self.dia_semana.setText(dia_semana_sem_data)

    def handle_localizacao_change(self):
        # Atualizar horários e desabilitar 'Unidade' se necessário
        self.update_fields_based_on_tecnico()

        selected_localizacao = self.combo_box_localizacao.currentText()
        if selected_localizacao in ["Folga", "Férias"]:
            self.combo_box_unidade.setDisabled(True)
            self.combo_box_unidade.setStyleSheet(
                self.combo_box_unidade.styleSheet() + "background-color: #e9ecef;"
            )
        else:
            self.combo_box_unidade.setDisabled(False)
            self.combo_box_unidade.setStyleSheet(
                self.combo_box_unidade.styleSheet().replace("background-color: #e9ecef;", "") +
                "background-color: #fff;"
            )

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.handle_enter_press()
        else:
            super().keyPressEvent(event)

    # Método para enviar e-mails usando Outlook (ajustado)
    def send_email(self, to_email, subject, body, send_time=None):
        try:
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = to_email
            mail.Subject = subject
            mail.Body = body

            if send_time:
                # Verifica se o horário agendado já passou
                if send_time <= datetime.datetime.now():
                    # Se o horário já passou, envia imediatamente
                    pass  # Não define DeferredDeliveryTime
                else:
                    # Se o horário está no futuro, agenda o envio
                    mail.DeferredDeliveryTime = send_time
            # Envia o e-mail
            mail.Send()

            # Força o envio/recebimento para garantir que o e-mail saia da caixa de saída
            namespace = outlook.GetNamespace("MAPI")
            namespace.SendAndReceive(False)

        except Exception as e:
            QMessageBox.warning(self, "Erro", f"Falha ao enviar e-mail para {to_email}: {e}")

    # Método para enviar e-mails para os técnicos selecionados
    def send_emails(self):
        tecnicos = list(self.technician_schedules.keys())
        dialog = EmailSelectionDialog(tecnicos, self.periodo_inicio, self.periodo_fim)
        if dialog.exec_() == QDialog.Accepted:
            selected_tecnicos = dialog.selected_tecnicos
            periodo_inicio = dialog.periodo_inicio
            periodo_fim = dialog.periodo_fim

            # Ler a planilha existente
            try:
                df_existing = pd.read_excel(self.planilha_path)
                df_existing.columns = df_existing.columns.str.upper()
            except FileNotFoundError:
                QMessageBox.warning(self, "Erro", "A planilha selecionada não foi encontrada.")
                return

            # Verificar se há dados
            if df_existing.empty:
                QMessageBox.information(self, "Aviso", "Não há dados na planilha para enviar.")
                return

            # Converter a coluna 'DATA/HORA INICIO' para datetime
            df_existing['DATA/HORA INICIO'] = pd.to_datetime(
                df_existing['DATA/HORA INICIO'], format="%d/%m/%Y %H:%M:%S", errors='coerce'
            )

            # Filtrar pelo período selecionado
            start_date = QDateTime(periodo_inicio, QTime(0, 0)).toPyDateTime()
            end_date = QDateTime(periodo_fim, QTime(23, 59, 59)).toPyDateTime()
            df_filtered = df_existing[
                (df_existing['DATA/HORA INICIO'] >= start_date) &
                (df_existing['DATA/HORA INICIO'] <= end_date) &
                (df_existing['TÉCNICO'].isin(selected_tecnicos))
            ]

            if df_filtered.empty:
                QMessageBox.information(self, "Aviso", "Não há registros para os técnicos e período selecionados.")
                return

            # Configurar a localidade para obter os dias da semana em português
            try:
                locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')  # Para sistemas Unix/Linux
            except:
                locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')  # Para Windows

            # Envio de e-mails para os técnicos
            grouped = df_filtered.groupby('TÉCNICO')

            for tecnico, group in grouped:
                email = technician_emails.get(tecnico)
                if not email:
                    QMessageBox.warning(self, "Aviso", f"E-mail do técnico {tecnico} não encontrado.")
                    continue

                message = f"Prezado(a) {tecnico},\n\nSegue sua escala:\n\n"
                for idx, row in group.iterrows():
                    dia_semana = row['DIA DA SEMANA'] if pd.notnull(row['DIA DA SEMANA']) else ''
                    data_hora_inicio = row['DATA/HORA INICIO']
                    data = data_hora_inicio.strftime('%d/%m/%Y') if pd.notnull(data_hora_inicio) else ''
                    unidade = row['UNIDADE'] if pd.notnull(row['UNIDADE']) and row['UNIDADE'] else ''
                    justificativa = row['JUSTIFICATIVA'] if pd.notnull(row['JUSTIFICATIVA']) and row['JUSTIFICATIVA'] else ''
                    card = row['CARD'] if pd.notnull(row['CARD']) and row['CARD'] else ''

                    # Ajuste realizado aqui
                    entry_message = ''
                    if dia_semana:
                        entry_message += f"Dia da Semana: {dia_semana}\n"
                    if data:
                        entry_message += f"Data: {data}\n"
                    if unidade:
                        entry_message += f"Unidade: {unidade}\n"
                    elif pd.notnull(row['LOCALIZAÇÃO']) and row['LOCALIZAÇÃO']:
                        localizacao = row['LOCALIZAÇÃO']
                        entry_message += f"Localização: {localizacao}\n"
                    if justificativa:
                        entry_message += f"Justificativa: {justificativa}\n"
                    if card:
                        entry_message += f"Card: {card}\n"

                    message += entry_message + '\n'  # Adiciona uma linha em branco entre as entradas

                message += "Atenciosamente,\nSua Equipe"

                self.send_email(email, "Sua Escala", message)

            # Envio agendado de e-mails para os gestores
            df_units = df_filtered.copy()
            df_units['DATA'] = df_units['DATA/HORA INICIO'].dt.date
            unit_grouped = df_units.groupby(['UNIDADE', 'DATA'])

            for (unidade, data_visita), group in unit_grouped:
                gestor_email = unit_manager_emails.get(unidade)
                if not gestor_email:
                    QMessageBox.warning(self, "Aviso", f"E-mail do gestor da unidade {unidade} não encontrado.")
                    continue

                # Lista de técnicos que estarão na unidade nesse dia
                tecnicos_na_unidade = group['TÉCNICO'].unique()
                tecnicos_lista = ', '.join(tecnicos_na_unidade)

                # Montar a mensagem
                mensagem = f"Prezado(a) Gestor(a),\n\nInformamos que o(s) técnico(s) {tecnicos_lista} estará(ão) presente(s) na unidade {unidade} no dia {data_visita.strftime('%d/%m/%Y')}.\n\nAtenciosamente,\nSua Equipe"

                # Agendar o envio para as 14:25 do dia da visita
                data_envio = datetime.datetime.combine(data_visita, datetime.time(14, 25))

                # Criar e enviar o e-mail agendado
                try:
                    self.send_email(gestor_email, f"Visita de Técnico - {data_visita.strftime('%d/%m/%Y')}", mensagem, send_time=data_envio)
                except Exception as e:
                    QMessageBox.warning(self, "Erro", f"Falha ao agendar e-mail para {gestor_email}: {e}")

            QMessageBox.information(self, "Sucesso", "E-mails enviados aos técnicos.")

# Classe ConsultaEscalaDialog
class ConsultaEscalaDialog(QDialog):
    def __init__(self, df_filtered, planilha_path, df_existing, periodo_inicio, periodo_fim, labels):
        super().__init__()
        self.setWindowIcon(QIcon('JT.ico'))
        self.df_filtered = df_filtered.copy()
        self.planilha_path = planilha_path
        self.df_existing = df_existing.copy()
        self.periodo_inicio = periodo_inicio
        self.periodo_fim = periodo_fim
        self.labels = labels  # Inclui 'SEQ' agora
        self.sort_columns = []  # Lista de colunas de ordenação
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Consulta de Escala")
        self.showMaximized()
        layout = QVBoxLayout()

        # Reordenar as colunas
        self.df_filtered = self.df_filtered[self.labels]

        # Converter colunas de data/hora para datetime
        for col in ['DATA/HORA INICIO', 'DATA/HORA FIM']:
            self.df_filtered[col] = pd.to_datetime(
                self.df_filtered[col], format="%d/%m/%Y %H:%M:%S", errors='coerce'
            )

        # Tabela para exibir os dados
        self.table_widget = QTableWidget()
        self.table_widget.setRowCount(len(self.df_filtered))
        self.table_widget.setColumnCount(len(self.df_filtered.columns))
        self.table_widget.setHorizontalHeaderLabels(self.df_filtered.columns.tolist())
        self.table_widget.horizontalHeader().setStretchLastSection(True)
        # Ajustar o tamanho das seções para que os títulos sejam completamente visíveis
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table_widget.horizontalHeader().setMinimumSectionSize(100)
        self.table_widget.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_widget.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_widget.setStyleSheet("""
            QTableWidget {
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
            }
            QTableWidget::item:selected {
                background-color: #007bff;
                color: #fff;
            }
            QHeaderView::section {
                background-color: #343a40;
                color: white;
                font-weight: bold;
                font-size: 14px;
                height: 30px;
            }
        """)

        # Habilitar a ordenação nos cabeçalhos
        self.table_widget.horizontalHeader().setSectionsClickable(True)
        self.table_widget.horizontalHeader().setSortIndicatorShown(True)
        self.table_widget.horizontalHeader().sectionClicked.connect(self.handle_header_click)

        # Preencher a tabela com os dados
        self.populate_table()

        layout.addWidget(self.table_widget)

        # Botões de editar, excluir, salvar e enviar escala
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()

        self.edit_button = QPushButton(" Editar")
        self.edit_button.setFixedSize(140, 40)
        self.edit_button.setIcon(QIcon("icons/edit.png"))
        self.edit_button.setStyleSheet(self.get_warning_button_style())
        self.edit_button.clicked.connect(self.enable_editing)
        buttons_layout.addWidget(self.edit_button)

        self.delete_button = QPushButton(" Excluir")
        self.delete_button.setFixedSize(140, 40)
        self.delete_button.setIcon(QIcon("icons/delete.png"))
        self.delete_button.setStyleSheet(self.get_danger_button_style())
        self.delete_button.clicked.connect(self.delete_entry)
        buttons_layout.addWidget(self.delete_button)

        self.save_button = QPushButton(" Salvar Alterações")
        self.save_button.setFixedSize(180, 40)
        self.save_button.setIcon(QIcon("icons/save.png"))
        self.save_button.setStyleSheet(self.get_primary_button_style())
        self.save_button.clicked.connect(self.save_changes)
        buttons_layout.addWidget(self.save_button)

        # Botão para Enviar E-mails
        self.send_email_button = QPushButton(" Enviar Escala")
        self.send_email_button.setFixedSize(180, 40)
        self.send_email_button.setIcon(QIcon("icons/email.png"))
        self.send_email_button.setStyleSheet(self.get_primary_button_style())
        self.send_email_button.clicked.connect(self.send_emails)
        buttons_layout.addWidget(self.send_email_button)

        layout.addLayout(buttons_layout)
        self.setLayout(layout)

    def get_primary_button_style(self):
        return """
            QPushButton {
                background-color: #007bff;
                color: white;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """

    def get_warning_button_style(self):
        return """
            QPushButton {
                background-color: #ffc107;
                color: #212529;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #e0a800;
            }
        """

    def get_danger_button_style(self):
        return """
            QPushButton {
                background-color: #dc3545;
                color: white;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #bd2130;
            }
        """

    def handle_header_click(self, logicalIndex):
        column_name = self.df_filtered.columns[logicalIndex]
        existing = next((item for item in self.sort_columns if item[0] == column_name), None)
        if existing:
            self.sort_columns.remove(existing)
            ascending = not existing[1]
        else:
            ascending = True

        self.sort_columns.insert(0, (column_name, ascending))
        order = Qt.AscendingOrder if ascending else Qt.DescendingOrder
        self.table_widget.horizontalHeader().setSortIndicator(logicalIndex, order)
        self.sort_table()

    def sort_table(self):
        if not self.sort_columns:
            return

        sort_by = [col for col, asc in self.sort_columns]
        ascending = [asc for col, asc in self.sort_columns]

        self.df_filtered.sort_values(
            by=sort_by, ascending=ascending, inplace=True,
            key=lambda col: col.str.lower() if col.dtype == object else col
        )

        self.df_filtered.reset_index(drop=True, inplace=True)
        self.populate_table()

    def populate_table(self):
        self.table_widget.setRowCount(len(self.df_filtered))
        for row in range(len(self.df_filtered)):
            for col in range(len(self.df_filtered.columns)):
                value = self.df_filtered.iloc[row, col]
                if pd.isna(value):
                    display_value = ''
                else:
                    if isinstance(value, pd.Timestamp):
                        display_value = value.strftime("%d/%m/%Y %H:%M:%S")
                    else:
                        display_value = str(value)
                item = QTableWidgetItem(display_value)

                if self.df_filtered.columns[col] == 'SEQ':
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)

                if self.df_filtered.columns[col] == 'LOCALIZAÇÃO':
                    location = display_value
                    color_map = {
                        'Unidade': '#17a2b8',
                        'Escritório': '#ffc107',
                        'Sobreaviso': '#fd7e14',
                        'Folga': '#6c757d',
                        'Home': '#20c997'
                    }
                    color = color_map.get(location, None)
                    if color:
                        item.setBackground(QColor(color))

                self.table_widget.setItem(row, col, item)

    def enable_editing(self):
        self.table_widget.setEditTriggers(QTableWidget.AllEditTriggers)

    def delete_entry(self):
        selected_rows = self.table_widget.selectionModel().selectedRows()
        if selected_rows:
            selected_row = selected_rows[0].row()
            self.table_widget.removeRow(selected_row)
            self.df_filtered = self.df_filtered.drop(self.df_filtered.index[selected_row]).reset_index(drop=True)
            QMessageBox.information(self, "Sucesso", "Entrada excluída com sucesso.")
        else:
            QMessageBox.warning(self, "Aviso", "Nenhuma linha selecionada para excluir.")

    def save_changes(self):
        # Atualizar o DataFrame filtrado com os valores da tabela
        for row_index in range(self.table_widget.rowCount()):
            for col_index in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row_index, col_index)
                if item:
                    column_name = self.df_filtered.columns[col_index]
                    value = item.text()
                    # Converter valor para o tipo de dado apropriado
                    if value == '':
                        if column_name in ['DATA/HORA INICIO', 'DATA/HORA FIM']:
                            value = pd.NaT
                        else:
                            value = pd.NA
                    else:
                        if column_name in ['DATA/HORA INICIO', 'DATA/HORA FIM']:
                            try:
                                # Usar dayfirst=True para aceitar formatos DD/MM/AAAA
                                value = pd.to_datetime(value, dayfirst=True, errors='raise')
                            except ValueError:
                                QMessageBox.warning(
                                    self,
                                    "Data/Hora Inválida",
                                    f"Formato de data/hora inválido para {column_name} na linha {row_index + 1}. Por favor, use o formato DD/MM/AAAA HH:MM:SS."
                                )
                                value = self.df_filtered.iloc[row_index, col_index]
                            except Exception as e:
                                value = self.df_filtered.iloc[row_index, col_index]
                        elif column_name == 'SEQ':
                            try:
                                value = int(value)
                            except ValueError:
                                QMessageBox.warning(self, "Valor Inválido", f"Valor inválido para SEQ na linha {row_index + 1}.")
                                value = self.df_filtered.iloc[row_index, col_index]
                        else:
                            value = str(value)
                    self.df_filtered.iloc[row_index, col_index] = value

        # Identificar as SEQs que foram excluídas
        # Primeiro, obter todas as SEQs do período no df_existing
        start_date = pd.to_datetime(self.periodo_inicio.toString("dd/MM/yyyy"), dayfirst=True)
        end_date = pd.to_datetime(self.periodo_fim.toString("dd/MM/yyyy"), dayfirst=True) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        df_existing_period = self.df_existing[
            (self.df_existing['DATA/HORA INICIO'] >= start_date) &
            (self.df_existing['DATA/HORA INICIO'] <= end_date)
        ]
        existing_seqs_in_period = set(df_existing_period['SEQ'])

        # SEQs no df_filtered
        filtered_seqs = set(self.df_filtered['SEQ'])

        # SEQs a serem excluídas
        seqs_to_delete = existing_seqs_in_period - filtered_seqs

        # Remover entradas do df_existing com essas SEQs
        if seqs_to_delete:
            self.df_existing = self.df_existing[~self.df_existing['SEQ'].isin(seqs_to_delete)]

        # Atualizar ou adicionar as entradas restantes
        for idx in self.df_filtered.index:
            seq = self.df_filtered.loc[idx, 'SEQ']
            if pd.isna(seq):
                max_seq = self.df_existing['SEQ'].max()
                if pd.isna(max_seq):
                    max_seq = 0
                else:
                    max_seq = int(max_seq)
                seq = max_seq + 1
                self.df_filtered.at[idx, 'SEQ'] = seq

            if (self.df_existing['SEQ'] == seq).any():
                for col in self.df_filtered.columns:
                    if col != 'SEQ':
                        self.df_existing.loc[self.df_existing['SEQ'] == seq, col] = self.df_filtered.loc[idx, col]
            else:
                self.df_existing = pd.concat([self.df_existing, self.df_filtered.loc[[idx]]], ignore_index=True)

        # Garantir que as colunas de data/hora estejam no formato correto (mantendo como datetime)
        for col in ['DATA/HORA INICIO', 'DATA/HORA FIM']:
            self.df_existing[col] = pd.to_datetime(self.df_existing[col], errors='coerce')

        try:
            with pd.ExcelWriter(self.planilha_path, engine='openpyxl') as writer:
                self.df_existing.to_excel(writer, index=False)

            wb = load_workbook(self.planilha_path)
            ws = wb.active

            columns = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}
            inicio_col = columns.get('DATA/HORA INICIO')
            fim_col = columns.get('DATA/HORA FIM')

            date_format = 'dd/mm/yyyy hh:mm:ss'

            if inicio_col:
                for cell in ws.iter_cols(min_col=inicio_col, max_col=inicio_col, min_row=2):
                    for c in cell:
                        c.number_format = date_format

            if fim_col:
                for cell in ws.iter_cols(min_col=fim_col, max_col=fim_col, min_row=2):
                    for c in cell:
                        c.number_format = date_format

            wb.save(self.planilha_path)

            QMessageBox.information(self, "Sucesso", f"Alterações salvas com sucesso em {self.planilha_path}!")
            self.table_widget.setEditTriggers(QTableWidget.NoEditTriggers)
        except Exception as e:
            QMessageBox.warning(self, "Erro", f"Falha ao salvar as alterações: {e}")

        self.populate_table()

    def send_emails(self):
        tecnicos = list(self.df_filtered['TÉCNICO'].unique())
        dialog = EmailSelectionDialog(tecnicos, self.periodo_inicio, self.periodo_fim)
        if dialog.exec_() == QDialog.Accepted:
            selected_tecnicos = dialog.selected_tecnicos
            periodo_inicio = dialog.periodo_inicio
            periodo_fim = dialog.periodo_fim

            # Filtrar df_filtered com base nos técnicos e período selecionados
            start_date = pd.to_datetime(periodo_inicio.toString("dd/MM/yyyy"), dayfirst=True)
            end_date = pd.to_datetime(periodo_fim.toString("dd/MM/yyyy"), dayfirst=True) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

            df_to_send = self.df_filtered[
                (self.df_filtered['DATA/HORA INICIO'] >= start_date) &
                (self.df_filtered['DATA/HORA INICIO'] <= end_date) &
                (self.df_filtered['TÉCNICO'].isin(selected_tecnicos))
            ]

            if df_to_send.empty:
                QMessageBox.information(self, "Aviso", "Não há registros para os técnicos e período selecionados.")
                return

            # Configurar a localidade para obter os dias da semana em português
            try:
                locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')  # Para sistemas Unix/Linux
            except:
                locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')  # Para Windows

            # Envio de e-mails para os técnicos
            grouped = df_to_send.groupby('TÉCNICO')

            for tecnico, group in grouped:
                email = technician_emails.get(tecnico)
                if not email:
                    QMessageBox.warning(self, "Aviso", f"E-mail do técnico {tecnico} não encontrado.")
                    continue

                message = f"Prezado(a) {tecnico},\n\nSegue sua escala:\n\n"
                for idx, row in group.iterrows():
                    dia_semana = row['DIA DA SEMANA'] if pd.notnull(row['DIA DA SEMANA']) else ''
                    data_hora_inicio = row['DATA/HORA INICIO']
                    data = data_hora_inicio.strftime('%d/%m/%Y') if pd.notnull(data_hora_inicio) else ''
                    unidade = row['UNIDADE'] if pd.notnull(row['UNIDADE']) and row['UNIDADE'] else ''
                    justificativa = row['JUSTIFICATIVA'] if pd.notnull(row['JUSTIFICATIVA']) and row['JUSTIFICATIVA'] else ''
                    card = row['CARD'] if pd.notnull(row['CARD']) and row['CARD'] else ''

                    # Ajuste realizado aqui
                    entry_message = ''
                    if dia_semana:
                        entry_message += f"Dia da Semana: {dia_semana}\n"
                    if data:
                        entry_message += f"Data: {data}\n"
                    if unidade:
                        entry_message += f"Unidade: {unidade}\n"
                    elif pd.notnull(row['LOCALIZAÇÃO']) and row['LOCALIZAÇÃO']:
                        localizacao = row['LOCALIZAÇÃO']
                        entry_message += f"Localização: {localizacao}\n"
                    if justificativa:
                        entry_message += f"Justificativa: {justificativa}\n"
                    if card:
                        entry_message += f"Card: {card}\n"

                    message += entry_message + '\n'  # Adiciona uma linha em branco entre as entradas

                message += "Atenciosamente,\nSua Equipe"

                self.send_email(email, "Sua Escala", message)

            QMessageBox.information(self, "Sucesso", "E-mails enviados aos técnicos.")

    # Método para enviar e-mails usando Outlook (ajustado)
    def send_email(self, to_email, subject, body, send_time=None):
        try:
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = to_email
            mail.Subject = subject
            mail.Body = body

            if send_time:
                # Verifica se o horário agendado já passou
                if send_time <= datetime.datetime.now():
                    # Se o horário já passou, envia imediatamente
                    pass  # Não define DeferredDeliveryTime
                else:
                    # Se o horário está no futuro, agenda o envio
                    mail.DeferredDeliveryTime = send_time
            # Envia o e-mail
            mail.Send()

            # Força o envio/recebimento para garantir que o e-mail saia da caixa de saída
            namespace = outlook.GetNamespace("MAPI")
            namespace.SendAndReceive(False)

        except Exception as e:
            QMessageBox.warning(self, "Erro", f"Falha ao enviar e-mail para {to_email}: {e}")

# Função main
def main():
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('JT.ico'))

    # Exibe a tela de seleção de planilha e período
    selection_dialog = SelectionDialog()
    if selection_dialog.exec_() == QDialog.Accepted:
        planilha_path = selection_dialog.planilha_path
        periodo_inicio = selection_dialog.periodo_inicio
        periodo_fim = selection_dialog.periodo_fim
        choice = selection_dialog.choice  # Obtém a escolha do usuário

        # Inicia a tela principal com os parâmetros selecionados
        form = ScheduleForm(planilha_path, periodo_inicio, periodo_fim)

        # Dependendo da escolha, abre a consulta de escala
        if choice == 'consultar_escala':
            form.consultar_escala(periodo_inicio, periodo_fim)

        sys.exit(app.exec_())
    else:
        sys.exit()

# Execução do script
if __name__ == '__main__':
    main()
