from __future__ import annotations

import os
import re
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from application.artifact_locator_service import ArtifactLocatorService
from application.manual_execution_service import (
    build_manual_company,
    has_manual_credentials,
    has_partial_manual_credentials,
    write_manual_company_csv,
)
from core.app_info import APP_DISPLAY_NAME, APP_VERSION, default_output_base_dir, project_dir as runtime_project_dir
from application.spreadsheet_validation_service import SpreadsheetValidationService
from core.config_runtime import (
    ExecutionProfile,
    RuntimeConfig,
    competencia_alvo_dir_name,
    competencias_alvo,
    mes_atual_str,
)
from core.paths import build_runtime_paths
from ui.controller import ExecutionController
from ui.models import CompanyRowState, EtapaExecucao, ResultadoFinal


PROJECT_DIR = runtime_project_dir()
ORQUESTRADOR_PATH = PROJECT_DIR / "orquestrador_empresas.py"


def only_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def format_competencia(ano: int, mes: int) -> str:
    return f"{mes:02d}/{ano}"


def open_in_os(path: Path) -> None:
    path = path.resolve()
    if not path.exists():
        raise FileNotFoundError(str(path))

    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
        return

    import subprocess

    if sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class MainWindow(QMainWindow):
    TABLE_HEADERS = [
        "Codigo",
        "Razao social",
        "CNPJ",
        "Resultado",
        "Etapa atual",
        "Ultima mensagem",
        "Tentativas",
        "Acao recomendada",
        "Inicio",
        "Fim",
    ]
    COL_CODIGO = 0
    COL_RAZAO = 1
    COL_CNPJ = 2
    COL_RESULTADO = 3
    COL_ETAPA = 4
    COL_MENSAGEM = 5
    COL_TENTATIVAS = 6
    COL_ACAO = 7
    COL_INICIO = 8
    COL_FIM = 9

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_DISPLAY_NAME} {APP_VERSION}")
        self.resize(1300, 850)

        self.company_row_by_cnpj: dict[str, int] = {}
        self.company_data_by_cnpj: dict[str, dict] = {}
        self.runtime_paths = build_runtime_paths(PROJECT_DIR, default_output_base_dir())
        self.last_manual_alert_key: str | None = None
        self.manual_runtime_file: Path | None = None

        self.validation_service = SpreadsheetValidationService()
        self.artifact_service = ArtifactLocatorService()
        self.controller = ExecutionController(ORQUESTRADOR_PATH, PROJECT_DIR, self.runtime_paths)

        self.controller.log_received.connect(self.append_log)
        self.controller.company_updated.connect(self.on_company_updated)
        self.controller.summary_updated.connect(self.on_summary_updated)
        self.controller.process_started.connect(self.on_process_started)
        self.controller.process_finished.connect(self.on_process_finished)
        self.controller.process_failed.connect(self.on_process_failed)
        self.controller.manual_wait_changed.connect(self.on_manual_wait_changed)

        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)

        outer_layout = QVBoxLayout(root)
        outer_layout.setContentsMargins(8, 8, 8, 8)
        outer_layout.setSpacing(8)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        layout.addWidget(self._build_group_arquivo())
        layout.addWidget(self._build_group_competencia_saida())
        layout.addWidget(self._build_group_etapas())
        layout.addWidget(self._build_group_execucao())
        layout.addWidget(self._build_group_manual())
        layout.addWidget(self._build_group_resumo())
        layout.addWidget(self._build_group_grade())
        layout.addWidget(self._build_group_log())
        layout.addWidget(self._build_group_acoes())
        layout.addStretch(1)

        scroll.setWidget(content)
        outer_layout.addWidget(scroll)

    def _build_group_arquivo(self) -> QGroupBox:
        group = QGroupBox("Arquivo")
        form = QFormLayout(group)

        row = QHBoxLayout()
        self.planilha_edit = QLineEdit()
        self.planilha_edit.setPlaceholderText("Selecione a planilha (.xlsx ou .csv)")
        self.btn_planilha = QPushButton("Selecionar planilha")
        self.btn_validar = QPushButton("Validar entrada")

        self.btn_planilha.clicked.connect(self.select_planilha)
        self.btn_validar.clicked.connect(self.validate_planilha)

        row.addWidget(self.planilha_edit)
        row.addWidget(self.btn_planilha)
        row.addWidget(self.btn_validar)

        form.addRow("Planilha:", row)

        self.manual_cnpj_edit = QLineEdit()
        self.manual_cnpj_edit.setPlaceholderText("Digite o CNPJ para executar uma unica empresa")
        self.manual_senha_edit = QLineEdit()
        self.manual_senha_edit.setPlaceholderText("Digite a senha da prefeitura")
        self.manual_senha_edit.setEchoMode(QLineEdit.EchoMode.Password)

        self.manual_cnpj_edit.textChanged.connect(self.update_input_mode_hint)
        self.manual_senha_edit.textChanged.connect(self.update_input_mode_hint)

        self.lbl_input_mode = QLabel()
        self.lbl_input_mode.setWordWrap(True)

        form.addRow("CNPJ manual:", self.manual_cnpj_edit)
        form.addRow("Senha manual:", self.manual_senha_edit)
        form.addRow("Modo:", self.lbl_input_mode)

        self.update_input_mode_hint()
        return group

    def _build_group_competencia_saida(self) -> QGroupBox:
        group = QGroupBox("Competencia e saida")
        form = QFormLayout(group)

        self.mes_atual_edit = QLineEdit(mes_atual_str())
        self.mes_atual_edit.setReadOnly(True)

        self.apuracao_edit = QLineEdit(mes_atual_str())
        self.apuracao_edit.setPlaceholderText("MM/AAAA")
        self.apuracao_edit.textChanged.connect(self.sync_execution_profile_widgets)

        row_saida = QHBoxLayout()
        self.output_edit = QLineEdit(str(default_output_base_dir()))
        self.btn_output = QPushButton("Selecionar pasta")
        self.btn_output_abrir = QPushButton("Abrir pasta de saida")
        self.btn_output_report = QPushButton("Abrir report")
        self.btn_output_checkpoint = QPushButton("Abrir checkpoint")

        self.btn_output.clicked.connect(self.select_output_dir)
        self.btn_output_abrir.clicked.connect(self.open_output_dir)
        self.btn_output_report.clicked.connect(self.open_report_file)
        self.btn_output_checkpoint.clicked.connect(self.open_checkpoint_file)

        row_saida.addWidget(self.output_edit)
        row_saida.addWidget(self.btn_output)
        row_saida.addWidget(self.btn_output_abrir)
        row_saida.addWidget(self.btn_output_report)
        row_saida.addWidget(self.btn_output_checkpoint)

        form.addRow("Mes atual:", self.mes_atual_edit)
        form.addRow("Mes de apuracao:", self.apuracao_edit)
        form.addRow("Pasta base:", row_saida)
        return group

    def _build_group_etapas(self) -> QGroupBox:
        group = QGroupBox("Perfil de execucao")
        layout = QVBoxLayout(group)

        grid = QGridLayout()

        self.chk_prestados = QCheckBox("Prestados")
        self.chk_tomados = QCheckBox("Tomados")
        self.chk_xml = QCheckBox("XML")
        self.chk_livros = QCheckBox("Livros")
        self.chk_iss = QCheckBox("ISS")
        self.chk_pausa_manual_final = QCheckBox("Pausa manual no final")
        self.chk_modo_automatico = QCheckBox("Modo automatico")
        self.chk_apurar_completo = QCheckBox("Apurar completo")

        self.chk_prestados.setChecked(True)
        self.chk_tomados.setChecked(True)
        self.chk_xml.setChecked(True)
        self.chk_livros.setChecked(True)
        self.chk_iss.setChecked(True)
        self.chk_pausa_manual_final.setChecked(True)
        self.chk_modo_automatico.setChecked(False)
        self.chk_apurar_completo.setChecked(False)

        profile_widgets = [
            self.chk_prestados,
            self.chk_tomados,
            self.chk_xml,
            self.chk_livros,
            self.chk_iss,
            self.chk_pausa_manual_final,
            self.chk_modo_automatico,
            self.chk_apurar_completo,
        ]
        for widget in profile_widgets:
            widget.toggled.connect(self.sync_execution_profile_widgets)

        grid.addWidget(self.chk_prestados, 0, 0)
        grid.addWidget(self.chk_tomados, 0, 1)
        grid.addWidget(self.chk_xml, 1, 0)
        grid.addWidget(self.chk_livros, 1, 1)
        grid.addWidget(self.chk_iss, 2, 0)
        grid.addWidget(self.chk_pausa_manual_final, 2, 1)
        grid.addWidget(self.chk_modo_automatico, 3, 0)
        grid.addWidget(self.chk_apurar_completo, 3, 1)

        self.lbl_execution_profile_hint = QLabel()
        self.lbl_execution_profile_hint.setWordWrap(True)

        layout.addLayout(grid)
        layout.addWidget(self.lbl_execution_profile_hint)

        self.sync_execution_profile_widgets()
        return group

    def _build_group_execucao(self) -> QGroupBox:
        group = QGroupBox("Execucao")
        row = QHBoxLayout(group)

        self.btn_iniciar = QPushButton("Iniciar lote")
        self.btn_parar = QPushButton("Parar")
        self.btn_parar.setEnabled(False)

        self.btn_iniciar.clicked.connect(self.start_worker)
        self.btn_parar.clicked.connect(self.stop_worker)

        row.addWidget(self.btn_iniciar)
        row.addWidget(self.btn_parar)
        row.addStretch(1)
        return group

    def _build_group_manual(self) -> QGroupBox:
        group = QGroupBox("Acao manual")
        row = QHBoxLayout(group)

        self.lbl_manual_status = QLabel("Nenhuma etapa manual aguardando.")
        self.lbl_manual_status.setWordWrap(True)
        self.lbl_manual_status.setStyleSheet("color: #7a1f1f; font-weight: 600;")

        self.btn_manual_continue = QPushButton("Continuar etapa manual")
        self.btn_manual_continue.setEnabled(False)
        self.btn_manual_continue.clicked.connect(self.continue_manual_stage)

        row.addWidget(self.lbl_manual_status, 1)
        row.addWidget(self.btn_manual_continue)
        return group

    def _build_group_resumo(self) -> QGroupBox:
        group = QGroupBox("Resumo")
        grid = QGridLayout(group)

        self.lbl_total = QLabel("0")
        self.lbl_execucao = QLabel("0")
        self.lbl_sucesso = QLabel("0")
        self.lbl_sem_comp = QLabel("0")
        self.lbl_sem_serv = QLabel("0")
        self.lbl_revisao_manual = QLabel("0")
        self.lbl_falha = QLabel("0")

        labels = [
            ("Total de empresas", self.lbl_total),
            ("Em execucao", self.lbl_execucao),
            ("SUCESSO", self.lbl_sucesso),
            ("SUCESSO_SEM_COMPETENCIA", self.lbl_sem_comp),
            ("SUCESSO_SEM_SERVICOS", self.lbl_sem_serv),
            ("REVISAO_MANUAL", self.lbl_revisao_manual),
            ("FALHA", self.lbl_falha),
        ]

        for i, (titulo, valor) in enumerate(labels):
            box = QGroupBox(titulo)
            box_layout = QVBoxLayout(box)
            valor.setAlignment(Qt.AlignmentFlag.AlignCenter)
            box_layout.addWidget(valor)
            grid.addWidget(box, 0, i)

        return group

    def _build_group_grade(self) -> QGroupBox:
        group = QGroupBox("Grade operacional")
        layout = QVBoxLayout(group)

        self.table = QTableWidget(0, len(self.TABLE_HEADERS))
        self.table.setHorizontalHeaderLabels(self.TABLE_HEADERS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setMinimumHeight(220)

        layout.addWidget(self.table)
        return group

    def _build_group_log(self) -> QGroupBox:
        group = QGroupBox("Log ao vivo")
        layout = QVBoxLayout(group)

        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setMinimumHeight(140)

        layout.addWidget(self.log_edit)
        return group

    def _build_group_acoes(self) -> QGroupBox:
        group = QGroupBox("Acoes da empresa selecionada")
        row = QHBoxLayout(group)

        self.btn_empresa_pasta = QPushButton("Abrir pasta")
        self.btn_empresa_log_nfse = QPushButton("Abrir log NFSe")
        self.btn_empresa_log_tomados = QPushButton("Abrir log Tomados")
        self.btn_empresa_log_manual = QPushButton("Abrir log manual")
        self.btn_empresa_evidencias = QPushButton("Abrir evidencias")
        self.btn_empresa_debug = QPushButton("Abrir debug")

        self.btn_empresa_pasta.clicked.connect(lambda: self.open_selected_company_target("pasta"))
        self.btn_empresa_log_nfse.clicked.connect(lambda: self.open_selected_company_target("log_nfse"))
        self.btn_empresa_log_tomados.clicked.connect(lambda: self.open_selected_company_target("log_tomados"))
        self.btn_empresa_log_manual.clicked.connect(lambda: self.open_selected_company_target("log_manual"))
        self.btn_empresa_evidencias.clicked.connect(lambda: self.open_selected_company_target("evidencias"))
        self.btn_empresa_debug.clicked.connect(lambda: self.open_selected_company_target("debug"))

        row.addWidget(self.btn_empresa_pasta)
        row.addWidget(self.btn_empresa_log_nfse)
        row.addWidget(self.btn_empresa_log_tomados)
        row.addWidget(self.btn_empresa_log_manual)
        row.addWidget(self.btn_empresa_evidencias)
        row.addWidget(self.btn_empresa_debug)
        row.addStretch(1)

        return group

    def append_log(self, text: str) -> None:
        self.log_edit.appendPlainText(text)
        self.log_edit.verticalScrollBar().setValue(self.log_edit.verticalScrollBar().maximum())

    def select_planilha(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar planilha",
            str(PROJECT_DIR),
            "Planilhas (*.xlsx *.csv)",
        )
        if path:
            self.planilha_edit.setText(path)

    def select_output_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self,
            "Selecionar pasta de saida",
            self.output_edit.text() or str(PROJECT_DIR),
        )
        if path:
            self.output_edit.setText(path)
            self._refresh_runtime_paths()

    def sync_execution_profile_widgets(self) -> None:
        apurar_completo = self.chk_apurar_completo.isChecked()

        if apurar_completo and not self.chk_xml.isChecked():
            self.chk_xml.blockSignals(True)
            self.chk_xml.setChecked(True)
            self.chk_xml.blockSignals(False)

        if apurar_completo and self.chk_livros.isChecked():
            self.chk_livros.blockSignals(True)
            self.chk_livros.setChecked(False)
            self.chk_livros.blockSignals(False)

        livros_ativos = self.chk_livros.isChecked()
        automatico = self.chk_modo_automatico.isChecked()

        if (apurar_completo or not livros_ativos) and self.chk_iss.isChecked():
            self.chk_iss.blockSignals(True)
            self.chk_iss.setChecked(False)
            self.chk_iss.blockSignals(False)

        pausa_manual_permitida = (not apurar_completo) and livros_ativos and not automatico
        if not pausa_manual_permitida and self.chk_pausa_manual_final.isChecked():
            self.chk_pausa_manual_final.blockSignals(True)
            self.chk_pausa_manual_final.setChecked(False)
            self.chk_pausa_manual_final.blockSignals(False)

        self.chk_xml.setEnabled(not apurar_completo)
        self.chk_livros.setEnabled(not apurar_completo)
        self.chk_iss.setEnabled((not apurar_completo) and livros_ativos)
        self.chk_pausa_manual_final.setEnabled(pausa_manual_permitida)

        partes = []
        modulos = []
        if self.chk_prestados.isChecked():
            modulos.append("Prestados")
        if self.chk_tomados.isChecked():
            modulos.append("Tomados")
        if modulos:
            partes.append("Modulos: " + ", ".join(modulos))
        else:
            partes.append("Modulos: nenhum selecionado")

        if apurar_completo:
            partes.append("Modo anual: XML somente")

        saidas = []
        if self.chk_xml.isChecked():
            saidas.append("XML")
        if not apurar_completo and self.chk_livros.isChecked():
            saidas.append("Livros")
        if not apurar_completo and self.chk_iss.isChecked():
            saidas.append("ISS")
        if not apurar_completo and self.chk_pausa_manual_final.isChecked():
            saidas.append("Pausa final")
        if self.chk_modo_automatico.isChecked():
            saidas.append("Modo automatico")

        if saidas:
            partes.append("Saidas: " + ", ".join(saidas))
        else:
            partes.append("Saidas: nenhuma selecionada")

        if apurar_completo:
            try:
                competencias = competencias_alvo(self.apuracao_edit.text().strip(), apurar_completo=True)
                inicio = format_competencia(*competencias[0])
                fim = format_competencia(*competencias[-1])
                partes.append(
                    f"Apuracao completa ativa: vai processar {len(competencias)} competencias, de {inicio} ate {fim}."
                )
            except ValueError as exc:
                partes.append(f"Apuracao completa ativa, mas a referencia esta invalida: {exc}")
            partes.append("Nunca inclui a competencia do mes atual em aberto.")

        partes.append("Tomados sempre mantem os PDFs individuais quando o modulo estiver ativo.")
        self.lbl_execution_profile_hint.setText(" | ".join(partes))

    def build_execution_profile(self) -> ExecutionProfile:
        profile = ExecutionProfile(
            executar_prestados=self.chk_prestados.isChecked(),
            executar_tomados=self.chk_tomados.isChecked(),
            executar_xml=self.chk_xml.isChecked(),
            executar_livros=self.chk_livros.isChecked(),
            executar_iss=self.chk_iss.isChecked(),
            pausa_manual_final=self.chk_pausa_manual_final.isChecked(),
            modo_automatico=self.chk_modo_automatico.isChecked(),
            apurar_completo=self.chk_apurar_completo.isChecked(),
        )
        profile.validate()
        return profile

    def _is_manual_mode(self) -> bool:
        return has_manual_credentials(self.manual_cnpj_edit.text(), self.manual_senha_edit.text())

    def _has_partial_manual_credentials(self) -> bool:
        return has_partial_manual_credentials(self.manual_cnpj_edit.text(), self.manual_senha_edit.text())

    def update_input_mode_hint(self) -> None:
        if self._is_manual_mode():
            self.lbl_input_mode.setText(
                "Modo manual ativo. Esta execucao vai ignorar a planilha e rodar somente o CNPJ informado."
            )
            self.lbl_input_mode.setStyleSheet("color: #1f4d2e; font-weight: 600;")
            return

        if self._has_partial_manual_credentials():
            self.lbl_input_mode.setText(
                "Modo manual incompleto. Preencha CNPJ e Senha para ignorar a planilha, "
                "ou deixe ambos vazios para usar o lote por planilha."
            )
            self.lbl_input_mode.setStyleSheet("color: #8a5a00; font-weight: 600;")
            return

        self.lbl_input_mode.setText(
            "Modo planilha ativo. Preencha CNPJ e Senha para executar uma unica empresa sem depender da planilha."
        )
        self.lbl_input_mode.setStyleSheet("color: #4a4a4a;")

    def _cleanup_manual_runtime_file(self, force: bool = False) -> None:
        if self.manual_runtime_file is None:
            return

        if self.controller.process is not None and not force:
            return

        try:
            self.manual_runtime_file.unlink(missing_ok=True)
        except OSError:
            pass
        finally:
            self.manual_runtime_file = None

    def _resolve_empresas_arquivo(self) -> Path:
        if self._is_manual_mode():
            empresa = build_manual_company(self.manual_cnpj_edit.text(), self.manual_senha_edit.text())
            self._cleanup_manual_runtime_file(force=True)
            self.manual_runtime_file = write_manual_company_csv(empresa)
            return self.manual_runtime_file

        if self._has_partial_manual_credentials():
            raise ValueError(
                "Para usar o modo manual, preencha CNPJ e Senha. "
                "Se quiser usar a planilha, deixe ambos vazios."
            )

        self._cleanup_manual_runtime_file(force=True)

        planilha_path = self.planilha_edit.text().strip()
        if not planilha_path:
            raise ValueError("Selecione a planilha ou informe CNPJ e Senha para o modo manual.")

        return Path(planilha_path)

    def _set_execution_inputs_enabled(self, enabled: bool) -> None:
        widgets = [
            self.planilha_edit,
            self.btn_planilha,
            self.btn_validar,
            self.manual_cnpj_edit,
            self.manual_senha_edit,
            self.apuracao_edit,
            self.output_edit,
            self.btn_output,
            self.chk_prestados,
            self.chk_tomados,
            self.chk_xml,
            self.chk_livros,
            self.chk_iss,
            self.chk_pausa_manual_final,
            self.chk_modo_automatico,
            self.chk_apurar_completo,
        ]
        for widget in widgets:
            widget.setEnabled(enabled)

    def build_config(self) -> RuntimeConfig:
        output_text = self.output_edit.text().strip()
        if not output_text:
            raise ValueError("Pasta base de saida nao informada.")

        cfg = RuntimeConfig(
            empresas_arquivo=self._resolve_empresas_arquivo(),
            apuracao_referencia=self.apuracao_edit.text().strip(),
            output_base_dir=Path(output_text),
            execution_profile=self.build_execution_profile(),
        )
        cfg.validate()
        self.runtime_paths = build_runtime_paths(PROJECT_DIR, cfg.output_base_dir)
        self.controller.runtime_paths = self.runtime_paths
        return cfg

    def validate_planilha(self, show_success_message: bool = True) -> bool:
        manual_mode = self._is_manual_mode()
        try:
            cfg = self.build_config()
            result = self.validation_service.validate(cfg)

            erros = [issue for issue in result.issues if issue.severity == "ERROR"]
            if erros:
                raise ValueError("\n".join(issue.message for issue in erros))

            self.table.setRowCount(0)
            self.company_row_by_cnpj.clear()
            self.company_data_by_cnpj.clear()

            for empresa in result.companies:
                cnpj = only_digits(empresa.get("cnpj", ""))
                row_state = CompanyRowState(
                    codigo=str(empresa.get("codigo", "")),
                    razao_social=str(empresa.get("razao_social", "")),
                    cnpj=str(empresa.get("cnpj", "")),
                    resultado=ResultadoFinal.AGUARDANDO.value,
                    etapa=EtapaExecucao.AGUARDANDO.value,
                )

                row = self.table.rowCount()
                self.table.insertRow(row)

                values = [
                    row_state.codigo,
                    row_state.razao_social,
                    row_state.cnpj,
                    row_state.resultado,
                    row_state.etapa,
                    row_state.ultima_mensagem,
                    row_state.tentativas,
                    row_state.acao_recomendada,
                    row_state.inicio,
                    row_state.fim,
                ]

                for col, value in enumerate(values):
                    self.table.setItem(row, col, QTableWidgetItem(str(value)))

                self.company_row_by_cnpj[cnpj] = row
                self.company_data_by_cnpj[cnpj] = empresa

            self.controller.load_companies(result.companies)

            self.lbl_total.setText(str(len(result.companies)))
            self.lbl_sucesso.setText("0")
            self.lbl_sem_comp.setText("0")
            self.lbl_sem_serv.setText("0")
            self.lbl_revisao_manual.setText("0")
            self.lbl_falha.setText("0")
            self.update_execution_count()

            avisos = []
            for issue in result.issues:
                if issue.severity != "WARNING":
                    continue
                prefix = f"Linha {issue.row_number}: " if issue.row_number is not None else ""
                avisos.append(prefix + issue.message)

            origem = "Entrada manual" if manual_mode else "Planilha"
            mensagem = f"{origem} validada com sucesso.\nEmpresas carregadas: {len(result.companies)}"
            if avisos:
                mensagem += "\n\nAvisos:\n- " + "\n- ".join(avisos[:10])

            if show_success_message:
                QMessageBox.information(self, "Validacao", mensagem)
            return True
        except Exception as e:
            self._cleanup_manual_runtime_file(force=True)
            self.reset_company_state()
            QMessageBox.critical(self, "Erro na validacao", str(e))
            return False
        finally:
            if manual_mode:
                self._cleanup_manual_runtime_file(force=True)

    def start_worker(self) -> None:
        try:
            if not self.validate_planilha(show_success_message=False):
                return

            cfg = self.build_config()
            if self.table.rowCount() == 0:
                self._cleanup_manual_runtime_file(force=True)
                return

            self.log_edit.clear()
            self.append_log("Iniciando lote...")
            self.controller.runtime_paths = self.runtime_paths
            self.controller.start(cfg)
        except Exception as e:
            self._cleanup_manual_runtime_file(force=True)
            QMessageBox.critical(self, "Erro ao iniciar", str(e))

    def stop_worker(self) -> None:
        self.append_log("Solicitacao de parada enviada.")
        self.controller.stop()

    def on_process_started(self) -> None:
        self.btn_iniciar.setEnabled(False)
        self.btn_parar.setEnabled(True)
        self._set_execution_inputs_enabled(False)
        self.lbl_manual_status.setText("Nenhuma etapa manual aguardando.")
        self.btn_manual_continue.setEnabled(False)
        self.append_log(f"Controle manual: {self.controller.control_dir}")
        self.append_log(f"Report esperado: {self.runtime_paths.report_path}")

    def on_process_finished(self, exit_code: int, exit_status: int) -> None:
        was_stop_requested = self.controller.stop_requested or exit_code == 130
        self.controller.stop_requested = False

        if was_stop_requested:
            self.mark_running_rows_as_interrupted()

        self.btn_iniciar.setEnabled(True)
        self.btn_parar.setEnabled(False)
        self._set_execution_inputs_enabled(True)
        self.lbl_manual_status.setText("Nenhuma etapa manual aguardando.")
        self.btn_manual_continue.setEnabled(False)
        self._cleanup_manual_runtime_file(force=True)
        self.update_execution_count()
        self.append_log(f"Processo finalizado. exit_code={exit_code} exit_status={exit_status}")

    def on_process_failed(self, message: str) -> None:
        self.controller.stop_requested = False
        self.btn_iniciar.setEnabled(True)
        self.btn_parar.setEnabled(False)
        self._set_execution_inputs_enabled(True)
        self.lbl_manual_status.setText("Nenhuma etapa manual aguardando.")
        self.btn_manual_continue.setEnabled(False)
        self._cleanup_manual_runtime_file(force=True)
        QMessageBox.critical(self, "Falha na execucao", message)

    def on_company_updated(self, cnpj: str, payload: dict) -> None:
        row = self.company_row_by_cnpj.get(cnpj)
        if row is None:
            return

        if "resultado" in payload:
            self.set_table_value(row, self.COL_RESULTADO, payload["resultado"])
        if "etapa" in payload:
            self.set_table_value(row, self.COL_ETAPA, payload["etapa"])
        if "ultima_mensagem" in payload:
            self.set_table_value(row, self.COL_MENSAGEM, payload["ultima_mensagem"])
        if "tentativas" in payload:
            self.set_table_value(row, self.COL_TENTATIVAS, payload["tentativas"])
        if "acao_recomendada" in payload:
            self.set_table_value(row, self.COL_ACAO, payload["acao_recomendada"])
        if "inicio" in payload:
            self.set_table_value(row, self.COL_INICIO, payload["inicio"])
        if "fim" in payload:
            self.set_table_value(row, self.COL_FIM, payload["fim"])

        self.update_execution_count()

    def on_summary_updated(self, resumo: dict) -> None:
        self.lbl_sucesso.setText(str(resumo.get("sucesso", 0)))
        self.lbl_sem_comp.setText(str(resumo.get("sem_comp", 0)))
        self.lbl_sem_serv.setText(str(resumo.get("sem_serv", 0)))
        self.lbl_revisao_manual.setText(str(resumo.get("revisao_manual", 0)))
        self.lbl_falha.setText(str(resumo.get("falha", 0)))
        self.update_execution_count()

    def on_manual_wait_changed(self, payload: dict) -> None:
        if not payload.get("active"):
            self.lbl_manual_status.setText("Nenhuma etapa manual aguardando.")
            self.btn_manual_continue.setEnabled(False)
            self.last_manual_alert_key = None
            return

        codigo = str(payload.get("codigo", "")).strip()
        razao = str(payload.get("razao_social", "")).strip()
        etapa = str(payload.get("etapa", "")).strip()
        mensagem = str(payload.get("mensagem", "")).strip()
        tag = str(payload.get("tag", "")).strip()
        empresa = " - ".join(part for part in [codigo, razao] if part)
        if not empresa:
            empresa = "Empresa em execucao"

        detalhes = [empresa]
        if etapa:
            detalhes.append(f"Etapa: {etapa}")
        if mensagem:
            detalhes.append(mensagem)
        detalhes.append("Depois de concluir no navegador, clique em 'Continuar etapa manual'.")
        self.lbl_manual_status.setText("Aguardando acao manual\n" + "\n".join(detalhes))
        self.btn_manual_continue.setEnabled(True)

        alert_key = "|".join([empresa, etapa, mensagem, tag])
        if alert_key != self.last_manual_alert_key:
            self.last_manual_alert_key = alert_key
            self.raise_()
            self.activateWindow()
            QMessageBox.information(
                self,
                "Etapa manual aguardando",
                "A automacao pausou para uma acao manual.\n\n" + "\n".join(detalhes),
            )

    def continue_manual_stage(self) -> None:
        self.controller.release_manual_wait()

    def open_output_dir(self) -> None:
        try:
            output_dir = Path(self.output_edit.text().strip())
            output_dir.mkdir(parents=True, exist_ok=True)
            open_in_os(output_dir)
        except Exception as e:
            QMessageBox.critical(self, "Erro ao abrir pasta", str(e))

    def open_report_file(self) -> None:
        try:
            self._refresh_runtime_paths()
            open_in_os(self.runtime_paths.report_path)
        except FileNotFoundError as e:
            QMessageBox.warning(self, "Report indisponivel", f"Report ainda nao foi gerado:\n{e}")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao abrir report", str(e))

    def open_checkpoint_file(self) -> None:
        try:
            self._refresh_runtime_paths()
            open_in_os(self.runtime_paths.checkpoint_path)
        except FileNotFoundError as e:
            QMessageBox.warning(self, "Checkpoint indisponivel", f"Checkpoint ainda nao foi gerado:\n{e}")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao abrir checkpoint", str(e))

    def open_selected_company_target(self, target: str) -> None:
        selected = self.table.currentRow()
        if selected < 0:
            QMessageBox.warning(self, "Selecao", "Selecione uma empresa na grade.")
            return

        cnpj_item = self.table.item(selected, self.COL_CNPJ)
        if cnpj_item is None:
            QMessageBox.warning(self, "Selecao", "Linha selecionada sem CNPJ.")
            return

        cnpj = only_digits(cnpj_item.text())
        empresa = self.company_data_by_cnpj.get(cnpj)
        if not empresa:
            QMessageBox.warning(self, "Selecao", "Empresa nao encontrada na memoria da interface.")
            return

        self._refresh_runtime_paths()

        try:
            competencia_dir = competencia_alvo_dir_name(self.apuracao_edit.text().strip())
        except ValueError:
            competencia_dir = ""
        artifacts = self.artifact_service.get_company_artifacts(
            self.runtime_paths,
            empresa,
            competencia_dir_name=competencia_dir,
        )

        try:
            if target == "pasta":
                open_in_os(artifacts.competencia_dir)
                return
            if target == "log_nfse":
                open_in_os(artifacts.log_downloads)
                return
            if target == "log_tomados":
                open_in_os(artifacts.log_tomados)
                return
            if target == "log_manual":
                open_in_os(artifacts.log_manual)
                return
            if target == "evidencias":
                open_in_os(artifacts.evidencias_dir)
                return
            if target == "debug":
                open_in_os(artifacts.debug_dir)
                return
        except FileNotFoundError as e:
            QMessageBox.warning(self, "Artefato indisponivel", f"Artefato ainda nao foi gerado:\n{e}")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao abrir item", str(e))

    def set_table_value(self, row: int, col: int, value: str) -> None:
        item = self.table.item(row, col)
        if item is None:
            item = QTableWidgetItem()
            self.table.setItem(row, col, item)
        item.setText(str(value))

    def reset_company_state(self) -> None:
        self.table.setRowCount(0)
        self.company_row_by_cnpj.clear()
        self.company_data_by_cnpj.clear()
        self._cleanup_manual_runtime_file(force=True)
        self.controller.load_companies([])
        self.lbl_total.setText("0")
        self.lbl_sucesso.setText("0")
        self.lbl_sem_comp.setText("0")
        self.lbl_sem_serv.setText("0")
        self.lbl_revisao_manual.setText("0")
        self.lbl_falha.setText("0")
        self.lbl_manual_status.setText("Nenhuma etapa manual aguardando.")
        self.btn_manual_continue.setEnabled(False)
        self.update_execution_count()

    def _refresh_runtime_paths(self) -> None:
        output_text = self.output_edit.text().strip()
        if not output_text:
            return
        self.runtime_paths = build_runtime_paths(PROJECT_DIR, Path(output_text))
        self.controller.runtime_paths = self.runtime_paths

    def update_execution_count(self) -> None:
        em_execucao = 0
        for row in range(self.table.rowCount()):
            item = self.table.item(row, self.COL_RESULTADO)
            if item and item.text() == ResultadoFinal.EM_EXECUCAO.value:
                em_execucao += 1
        self.lbl_execucao.setText(str(em_execucao))

    def mark_running_rows_as_interrupted(self) -> None:
        for row in range(self.table.rowCount()):
            status_item = self.table.item(row, self.COL_RESULTADO)
            if not status_item or status_item.text() != ResultadoFinal.EM_EXECUCAO.value:
                continue

            self.set_table_value(row, self.COL_RESULTADO, ResultadoFinal.INTERROMPIDO.value)
            if not (self.table.item(row, self.COL_FIM) and self.table.item(row, self.COL_FIM).text()):
                self.set_table_value(row, self.COL_FIM, self._now_text())

    @staticmethod
    def _now_text() -> str:
        from datetime import datetime

        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
