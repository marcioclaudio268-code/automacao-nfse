from __future__ import annotations

import os
import re
import sys
from pathlib import Path

from PySide6.QtCore import QProcess, QTimer, Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QHeaderView,
)

from core.config_runtime import RuntimeConfig, mes_atual_str
from core.event_parser import load_report_rows, parse_log_line
from core.paths import build_runtime_paths, choose_existing


PROJECT_DIR = Path(__file__).resolve().parents[1]
WORKER_PATH = PROJECT_DIR / "worker_lote.py"


def only_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


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
        "Código",
        "Razão social",
        "CNPJ",
        "Status",
        "Etapa atual",
        "Última mensagem",
        "Início",
        "Fim",
    ]

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Automação Prefeitura")
        self.resize(1300, 850)

        self.process: QProcess | None = None
        self.current_company_key: str | None = None
        self.company_row_by_cnpj: dict[str, int] = {}
        self.company_data_by_cnpj: dict[str, dict] = {}
        self.runtime_paths = build_runtime_paths(PROJECT_DIR, PROJECT_DIR)

        self._build_ui()
        self._build_timer()

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        layout.addWidget(self._build_group_arquivo())
        layout.addWidget(self._build_group_competencia_saida())
        layout.addWidget(self._build_group_execucao())
        layout.addWidget(self._build_group_resumo())
        layout.addWidget(self._build_group_grade())
        layout.addWidget(self._build_group_log())
        layout.addWidget(self._build_group_acoes())

    def _build_group_arquivo(self) -> QGroupBox:
        group = QGroupBox("Arquivo")
        form = QFormLayout(group)

        row = QHBoxLayout()
        self.planilha_edit = QLineEdit()
        self.planilha_edit.setPlaceholderText("Selecione a planilha (.xlsx ou .csv)")
        self.btn_planilha = QPushButton("Selecionar planilha")
        self.btn_validar = QPushButton("Validar planilha")

        self.btn_planilha.clicked.connect(self.select_planilha)
        self.btn_validar.clicked.connect(self.validate_planilha)

        row.addWidget(self.planilha_edit)
        row.addWidget(self.btn_planilha)
        row.addWidget(self.btn_validar)

        form.addRow("Planilha:", row)
        return group

    def _build_group_competencia_saida(self) -> QGroupBox:
        group = QGroupBox("Competência e saída")
        form = QFormLayout(group)

        self.mes_atual_edit = QLineEdit(mes_atual_str())
        self.mes_atual_edit.setReadOnly(True)

        self.apuracao_edit = QLineEdit(mes_atual_str())
        self.apuracao_edit.setPlaceholderText("MM/AAAA")

        row_saida = QHBoxLayout()
        self.output_edit = QLineEdit(str(PROJECT_DIR))
        self.btn_output = QPushButton("Selecionar pasta")
        self.btn_output_abrir = QPushButton("Abrir pasta de saída")

        self.btn_output.clicked.connect(self.select_output_dir)
        self.btn_output_abrir.clicked.connect(self.open_output_dir)

        row_saida.addWidget(self.output_edit)
        row_saida.addWidget(self.btn_output)
        row_saida.addWidget(self.btn_output_abrir)

        form.addRow("Mês atual:", self.mes_atual_edit)
        form.addRow("Mês de apuração:", self.apuracao_edit)
        form.addRow("Pasta base:", row_saida)
        return group

    def _build_group_execucao(self) -> QGroupBox:
        group = QGroupBox("Execução")
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

    def _build_group_resumo(self) -> QGroupBox:
        group = QGroupBox("Resumo")
        grid = QGridLayout(group)

        self.lbl_total = QLabel("0")
        self.lbl_execucao = QLabel("0")
        self.lbl_sucesso = QLabel("0")
        self.lbl_sem_comp = QLabel("0")
        self.lbl_sem_serv = QLabel("0")
        self.lbl_falha = QLabel("0")

        labels = [
            ("Total de empresas", self.lbl_total),
            ("Em execução", self.lbl_execucao),
            ("SUCESSO", self.lbl_sucesso),
            ("SUCESSO_SEM_COMPETENCIA", self.lbl_sem_comp),
            ("SUCESSO_SEM_SERVICOS", self.lbl_sem_serv),
            ("FALHA", self.lbl_falha),
        ]

        for i, (titulo, valor) in enumerate(labels):
            box = QGroupBox(titulo)
            box_layout = QVBoxLayout(box)
            valor.setAlignment(Qt.AlignCenter)
            box_layout.addWidget(valor)
            grid.addWidget(box, 0, i)

        return group

    def _build_group_grade(self) -> QGroupBox:
        group = QGroupBox("Grade operacional")
        layout = QVBoxLayout(group)

        self.table = QTableWidget(0, len(self.TABLE_HEADERS))
        self.table.setHorizontalHeaderLabels(self.TABLE_HEADERS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)

        layout.addWidget(self.table)
        return group

    def _build_group_log(self) -> QGroupBox:
        group = QGroupBox("Log ao vivo")
        layout = QVBoxLayout(group)

        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)

        layout.addWidget(self.log_edit)
        return group

    def _build_group_acoes(self) -> QGroupBox:
        group = QGroupBox("Ações da empresa selecionada")
        row = QHBoxLayout(group)

        self.btn_empresa_pasta = QPushButton("Abrir pasta")
        self.btn_empresa_log = QPushButton("Abrir log")
        self.btn_empresa_evidencias = QPushButton("Abrir evidências")
        self.btn_empresa_debug = QPushButton("Abrir debug")

        self.btn_empresa_pasta.clicked.connect(lambda: self.open_selected_company_target("pasta"))
        self.btn_empresa_log.clicked.connect(lambda: self.open_selected_company_target("log"))
        self.btn_empresa_evidencias.clicked.connect(lambda: self.open_selected_company_target("evidencias"))
        self.btn_empresa_debug.clicked.connect(lambda: self.open_selected_company_target("debug"))

        row.addWidget(self.btn_empresa_pasta)
        row.addWidget(self.btn_empresa_log)
        row.addWidget(self.btn_empresa_evidencias)
        row.addWidget(self.btn_empresa_debug)
        row.addStretch(1)

        return group

    def _build_timer(self) -> None:
        self.report_timer = QTimer(self)
        self.report_timer.setInterval(1500)
        self.report_timer.timeout.connect(self.sync_report)

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
            "Selecionar pasta de saída",
            self.output_edit.text() or str(PROJECT_DIR),
        )
        if path:
            self.output_edit.setText(path)

    def build_config(self) -> RuntimeConfig:
        cfg = RuntimeConfig(
            empresas_arquivo=Path(self.planilha_edit.text().strip()),
            apuracao_referencia=self.apuracao_edit.text().strip(),
            output_base_dir=Path(self.output_edit.text().strip()),
        )
        cfg.validate()
        self.runtime_paths = build_runtime_paths(PROJECT_DIR, cfg.output_base_dir)
        return cfg

    def validate_planilha(self) -> None:
        try:
            cfg = self.build_config()

            from orquestrador_empresas import carregar_empresas

            empresas = carregar_empresas(str(cfg.empresas_arquivo))
            if not empresas:
                raise ValueError("Nenhuma empresa válida foi encontrada na planilha.")

            self.table.setRowCount(0)
            self.company_row_by_cnpj.clear()
            self.company_data_by_cnpj.clear()

            for empresa in empresas:
                cnpj = only_digits(empresa.get("cnpj", ""))
                row = self.table.rowCount()
                self.table.insertRow(row)

                values = [
                    empresa.get("codigo", ""),
                    empresa.get("razao_social", ""),
                    empresa.get("cnpj", ""),
                    "AGUARDANDO",
                    "",
                    "",
                    "",
                    "",
                ]

                for col, value in enumerate(values):
                    self.table.setItem(row, col, QTableWidgetItem(str(value)))

                self.company_row_by_cnpj[cnpj] = row
                self.company_data_by_cnpj[cnpj] = empresa

            self.lbl_total.setText(str(len(empresas)))
            self.lbl_execucao.setText("0")
            self.lbl_sucesso.setText("0")
            self.lbl_sem_comp.setText("0")
            self.lbl_sem_serv.setText("0")
            self.lbl_falha.setText("0")

            QMessageBox.information(self, "Validação", f"Planilha validada com sucesso.\nEmpresas carregadas: {len(empresas)}")
        except Exception as e:
            QMessageBox.critical(self, "Erro na validação", str(e))

    def start_worker(self) -> None:
        try:
            cfg = self.build_config()
            if self.table.rowCount() == 0:
                self.validate_planilha()

            self.log_edit.clear()
            self.append_log("Iniciando lote...")

            self.process = QProcess(self)
            self.process.setProgram(sys.executable)
            self.process.setArguments([
                str(WORKER_PATH),
                "--planilha", str(cfg.empresas_arquivo),
                "--apuracao", cfg.apuracao_referencia,
                "--output-dir", str(cfg.output_base_dir),
                "--login-wait-seconds", str(cfg.login_wait_seconds),
                "--timeout-processo-main", str(cfg.timeout_processo_main),
                "--continuar", "1" if cfg.continuar_de_onde_parou else "0",
                "--usar-checkpoint", "1" if cfg.usar_checkpoint else "0",
            ])
            self.process.setWorkingDirectory(str(PROJECT_DIR))
            self.process.readyReadStandardOutput.connect(self.read_process_output)
            self.process.finished.connect(self.on_process_finished)
            self.process.start()

            self.btn_iniciar.setEnabled(False)
            self.btn_parar.setEnabled(True)
            self.report_timer.start()
        except Exception as e:
            QMessageBox.critical(self, "Erro ao iniciar", str(e))

    def stop_worker(self) -> None:
        if not self.process:
            return

        self.append_log("Solicitação de parada enviada.")
        self.process.terminate()

        if not self.process.waitForFinished(3000):
            self.process.kill()

    def read_process_output(self) -> None:
        if not self.process:
            return

        data = self.process.readAllStandardOutput().data().decode("utf-8", errors="replace")
        for line in data.splitlines():
            self.append_log(line)
            self.handle_event_line(line)

    def handle_event_line(self, line: str) -> None:
        event = parse_log_line(line)

        if event.kind == "empresa_atual":
            codigo_event = event.codigo.strip()
            for cnpj, empresa in self.company_data_by_cnpj.items():
                if str(empresa.get("codigo", "")).strip() == codigo_event:
                    self.current_company_key = cnpj
                    row = self.company_row_by_cnpj.get(cnpj)
                    if row is not None:
                        self.set_table_value(row, 3, "EM_EXECUCAO")
                        self.set_table_value(row, 4, event.etapa)
                        self.set_table_value(row, 5, event.mensagem)
                        if not self.table.item(row, 6) or not self.table.item(row, 6).text():
                            self.set_table_value(row, 6, self._now_text())
                    break
            return

        if self.current_company_key and event.kind in {"etapa", "erro", "log"}:
            row = self.company_row_by_cnpj.get(self.current_company_key)
            if row is not None:
                if event.etapa:
                    self.set_table_value(row, 4, event.etapa)
                if event.mensagem:
                    self.set_table_value(row, 5, event.mensagem)

    def on_process_finished(self) -> None:
        self.report_timer.stop()
        self.sync_report()

        self.btn_iniciar.setEnabled(True)
        self.btn_parar.setEnabled(False)
        self.append_log("Processo finalizado.")

    def sync_report(self) -> None:
        report_path = choose_existing(self.runtime_paths.report_path, self.runtime_paths.legacy_report_path)
        rows = load_report_rows(report_path)

        sucesso = 0
        sem_comp = 0
        sem_serv = 0
        falha = 0

        for row in rows:
            cnpj = only_digits(row.get("cnpj", ""))
            table_row = self.company_row_by_cnpj.get(cnpj)
            if table_row is None:
                continue

            status = (row.get("status") or "").strip()
            motivo = (row.get("motivo") or "").strip()

            self.set_table_value(table_row, 3, status)
            self.set_table_value(table_row, 5, motivo)
            self.set_table_value(table_row, 6, row.get("timestamp_inicio", ""))
            self.set_table_value(table_row, 7, row.get("timestamp_fim", ""))

            if status == "SUCESSO":
                sucesso += 1
            elif status == "SUCESSO_SEM_COMPETENCIA":
                sem_comp += 1
            elif status == "SUCESSO_SEM_SERVICOS":
                sem_serv += 1
            elif status == "FALHA":
                falha += 1

        em_execucao = 0
        for row in range(self.table.rowCount()):
            status_item = self.table.item(row, 3)
            if status_item and status_item.text() == "EM_EXECUCAO":
                em_execucao += 1

        self.lbl_execucao.setText(str(em_execucao))
        self.lbl_sucesso.setText(str(sucesso))
        self.lbl_sem_comp.setText(str(sem_comp))
        self.lbl_sem_serv.setText(str(sem_serv))
        self.lbl_falha.setText(str(falha))

    def open_output_dir(self) -> None:
        try:
            output_dir = Path(self.output_edit.text().strip())
            output_dir.mkdir(parents=True, exist_ok=True)
            open_in_os(output_dir)
        except Exception as e:
            QMessageBox.critical(self, "Erro ao abrir pasta", str(e))

    def open_selected_company_target(self, target: str) -> None:
        selected = self.table.currentRow()
        if selected < 0:
            QMessageBox.warning(self, "Seleção", "Selecione uma empresa na grade.")
            return

        cnpj = only_digits(self.table.item(selected, 2).text())
        empresa = self.company_data_by_cnpj.get(cnpj)
        if not empresa:
            QMessageBox.warning(self, "Seleção", "Empresa não encontrada na memória da interface.")
            return

        from orquestrador_empresas import montar_pasta_empresa_padrao

        pasta_nome = montar_pasta_empresa_padrao(empresa)
        base_downloads = choose_existing(self.runtime_paths.downloads_dir, self.runtime_paths.legacy_downloads_dir)
        pasta_empresa = base_downloads / pasta_nome

        try:
            if target == "pasta":
                open_in_os(pasta_empresa)
                return

            if target == "log":
                open_in_os(pasta_empresa / "log_downloads_nfse.txt")
                return

            if target == "evidencias":
                open_in_os(pasta_empresa / "_evidencias")
                return

            if target == "debug":
                open_in_os(pasta_empresa / "_debug")
                return
        except Exception as e:
            QMessageBox.critical(self, "Erro ao abrir item", str(e))

    def set_table_value(self, row: int, col: int, value: str) -> None:
        item = self.table.item(row, col)
        if item is None:
            item = QTableWidgetItem()
            self.table.setItem(row, col, item)
        item.setText(str(value))

    @staticmethod
    def _now_text() -> str:
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")