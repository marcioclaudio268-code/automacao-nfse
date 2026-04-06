from __future__ import annotations

from datetime import datetime
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from uuid import uuid4

from PySide6.QtCore import QObject, QProcess, QProcessEnvironment, QTimer, Signal

from core.app_info import APP_ORCHESTRATOR_EXE_NAME, is_frozen_app
from core.company_paths import nome_pasta_empresa_por_dados
from core.config_runtime import RuntimeConfig
from core.event_parser import load_report_rows, parse_log_line
from ui.models import ResultadoFinal


def only_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


class ExecutionController(QObject):
    log_received = Signal(str)
    company_updated = Signal(str, dict)
    summary_updated = Signal(dict)
    process_started = Signal()
    process_finished = Signal(int, int)
    process_failed = Signal(str)
    manual_wait_changed = Signal(dict)

    def __init__(self, orchestrator_path: Path, project_dir: Path, runtime_paths):
        super().__init__()
        self.orchestrator_path = orchestrator_path
        self.project_dir = project_dir
        self.runtime_paths = runtime_paths
        self.process: QProcess | None = None
        self.current_company_key: str | None = None
        self.manual_wait_company_key: str | None = None
        self.company_data_by_cnpj: dict[str, dict] = {}
        self.stop_requested = False
        self._stdout_buffer = ""
        self._stderr_buffer = ""
        self.control_dir: Path | None = None
        self.manual_wait_dir: Path | None = None
        self.manual_wait_token: str | None = None

        self.report_timer = QTimer(self)
        self.report_timer.setInterval(1500)
        self.report_timer.timeout.connect(self.sync_report)

    def load_companies(self, companies: list[dict]) -> None:
        self.company_data_by_cnpj.clear()
        for empresa in companies:
            cnpj = only_digits(empresa.get("cnpj", ""))
            self.company_data_by_cnpj[cnpj] = empresa

    def start(self, cfg: RuntimeConfig) -> None:
        if self.process is not None:
            self.process_failed.emit("Ja existe um lote em execucao.")
            return

        if (not is_frozen_app()) and (not self.orchestrator_path.exists()):
            self.process_failed.emit(f"Orquestrador nao encontrado: {self.orchestrator_path}")
            return

        self.stop_requested = False
        self.current_company_key = None
        self.manual_wait_company_key = None
        self.manual_wait_dir = None
        self.manual_wait_token = None
        self._stdout_buffer = ""
        self._stderr_buffer = ""
        self.control_dir = cfg.output_base_dir / "_control" / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}"
        self.control_dir.mkdir(parents=True, exist_ok=True)
        self.manual_wait_changed.emit({"active": False})

        env = QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONUNBUFFERED", "1")
        env.insert("EMPRESAS_ARQUIVO", str(cfg.empresas_arquivo))
        env.insert("APURACAO_REFERENCIA", cfg.apuracao_referencia)
        env.insert("OUTPUT_BASE_DIR", str(cfg.output_base_dir))
        env.insert("EXECUTION_CONTROL_DIR", str(self.control_dir))
        env.insert("LOGIN_WAIT_SECONDS", str(cfg.login_wait_seconds))
        env.insert("TIMEOUT_PROCESSO_MAIN", str(cfg.timeout_processo_main))
        env.insert("CONTINUAR_DE_ONDE_PAROU", "1" if cfg.continuar_de_onde_parou else "0")
        env.insert("USAR_CHECKPOINT", "1" if cfg.usar_checkpoint else "0")
        env.insert("PERFIL_EXECUCAO_ATIVO", "0")

        if not cfg.execution_profile.is_default():
            for key, value in cfg.execution_profile.as_env().items():
                env.insert(key, value)

        launch = self.resolve_orchestrator_launch()

        process = QProcess(self)
        process.setProcessEnvironment(env)
        process.setProgram(launch["program"])
        process.setArguments(launch["args"])
        process.setWorkingDirectory(str(launch["workdir"]))
        process.readyReadStandardOutput.connect(self.read_stdout)
        process.readyReadStandardError.connect(self.read_stderr)
        process.started.connect(self.on_started)
        process.errorOccurred.connect(self.on_error)
        process.finished.connect(self.on_finished)
        process.start()

        self.process = process
        if not process.waitForStarted(5000):
            message = f"Falha ao iniciar o orquestrador: {process.errorString()}"
            self.log_received.emit(message)
            self.process = None
            process.deleteLater()
            self._cleanup_control_dir()
            self.process_failed.emit(message)
            return

        self.report_timer.start()

    def stop(self) -> None:
        if not self.process:
            return

        self.stop_requested = True
        self.manual_wait_changed.emit({"active": False})
        pid = int(self.process.processId())
        if os.name == "nt" and pid > 0:
            try:
                subprocess.run(
                    ["taskkill", "/PID", str(pid), "/T", "/F"],
                    check=False,
                    capture_output=True,
                    text=True,
                    timeout=10,
                )
            except Exception as exc:
                self.log_received.emit(f"Falha ao encerrar arvore do processo: {exc}")

            if not self.process.waitForFinished(5000):
                self.process.kill()
            return

        self.process.terminate()
        if not self.process.waitForFinished(3000):
            self.process.kill()

    def read_stdout(self) -> None:
        if not self.process:
            return

        data = self.process.readAllStandardOutput().data().decode("utf-8", errors="replace")
        self._stdout_buffer += data
        self._stdout_buffer = self._emit_buffered_lines(self._stdout_buffer, is_stderr=False)

    def read_stderr(self) -> None:
        if not self.process:
            return

        data = self.process.readAllStandardError().data().decode("utf-8", errors="replace")
        self._stderr_buffer += data
        self._stderr_buffer = self._emit_buffered_lines(self._stderr_buffer, is_stderr=True)

    def _emit_buffered_lines(self, buffer: str, is_stderr: bool) -> str:
        lines = buffer.splitlines(keepends=True)
        tail = ""
        if lines and not lines[-1].endswith(("\n", "\r")):
            tail = lines.pop()

        for raw_line in lines:
            line = raw_line.rstrip("\r\n")
            text = f"[stderr] {line}" if is_stderr else line
            self.log_received.emit(text)
            if not is_stderr:
                self.handle_event_line(line)
        return tail

    def handle_event_line(self, line: str) -> None:
        event = parse_log_line(line)

        if event.kind == "empresa_atual":
            if self.manual_wait_company_key:
                self.manual_wait_company_key = None
                self.manual_wait_changed.emit({"active": False})
            codigo_event = event.codigo.strip()
            for cnpj, empresa in self.company_data_by_cnpj.items():
                if str(empresa.get("codigo", "")).strip() != codigo_event:
                    continue

                self.current_company_key = cnpj
                self.company_updated.emit(
                    cnpj,
                    {
                        "resultado": ResultadoFinal.EM_EXECUCAO.value,
                        "etapa": event.etapa or "Preparando empresa",
                        "ultima_mensagem": event.mensagem,
                        "inicio": self._now_text(),
                    },
                )
                return

        if event.kind == "manual_wait" and self.current_company_key:
            empresa = self.company_data_by_cnpj.get(self.current_company_key, {})
            self.manual_wait_company_key = self.current_company_key
            self.manual_wait_changed.emit(
                {
                    "active": True,
                    "cnpj": self.current_company_key,
                    "codigo": str(empresa.get("codigo", "")),
                    "razao_social": str(empresa.get("razao_social", "")),
                    "etapa": event.etapa,
                    "mensagem": event.mensagem,
                }
            )

        if event.kind == "manual_resumed":
            self.manual_wait_company_key = None
            self.manual_wait_changed.emit({"active": False})

        if self.current_company_key and event.kind in {"etapa", "erro", "log", "final", "interrompido", "manual_wait", "manual_resumed"}:
            payload: dict[str, str] = {}
            if event.etapa:
                payload["etapa"] = event.etapa
            if event.mensagem:
                payload["ultima_mensagem"] = event.mensagem
            if event.kind == "interrompido":
                payload["resultado"] = ResultadoFinal.INTERROMPIDO.value
            self.company_updated.emit(self.current_company_key, payload)

    def sync_report(self) -> None:
        self.sync_manual_wait()
        rows = load_report_rows(self.runtime_paths.report_path)
        resumo = {
            "sucesso": 0,
            "sem_comp": 0,
            "sem_serv": 0,
            "revisao_manual": 0,
            "falha": 0,
        }

        for row in rows:
            cnpj = only_digits(row.get("cnpj", ""))
            status = (row.get("status") or "").strip()
            motivo = (row.get("motivo") or "").strip()

            self.company_updated.emit(
                cnpj,
                {
                    "resultado": status,
                    "ultima_mensagem": motivo,
                    "tentativas": row.get("tentativas", "0"),
                    "inicio": row.get("timestamp_inicio", ""),
                    "fim": row.get("timestamp_fim", ""),
                    "acao_recomendada": row.get("acao_recomendada", ""),
                },
            )

            if status == ResultadoFinal.SUCESSO.value:
                resumo["sucesso"] += 1
            elif status == ResultadoFinal.SUCESSO_SEM_COMPETENCIA.value:
                resumo["sem_comp"] += 1
            elif status == ResultadoFinal.SUCESSO_SEM_SERVICOS.value:
                resumo["sem_serv"] += 1
            elif status == ResultadoFinal.REVISAO_MANUAL.value:
                resumo["revisao_manual"] += 1
            elif status == ResultadoFinal.FALHA.value:
                resumo["falha"] += 1

        self.summary_updated.emit(resumo)

    def on_started(self) -> None:
        self.process_started.emit()

    def on_error(self, process_error) -> None:
        if not self.process:
            return

        if process_error == QProcess.ProcessError.FailedToStart:
            message = f"Falha ao iniciar o orquestrador: {self.process.errorString()}"
            self.report_timer.stop()
            self.log_received.emit(message)
            self.process.deleteLater()
            self.process = None
            self._cleanup_control_dir()
            self.process_failed.emit(message)

    def on_finished(self, exit_code: int, exit_status: int) -> None:
        self._flush_remaining_buffer(is_stderr=False)
        self._flush_remaining_buffer(is_stderr=True)

        self.report_timer.stop()
        self.sync_report()
        self.manual_wait_company_key = None
        self.manual_wait_changed.emit({"active": False})

        process = self.process
        self.process = None
        if process is not None:
            process.deleteLater()

        self._cleanup_control_dir()

        self.process_finished.emit(int(exit_code), int(exit_status.value))

    def release_manual_wait(self) -> None:
        if self.manual_wait_dir:
            continue_path = self.manual_wait_dir / "continue.signal"
            continue_path.parent.mkdir(parents=True, exist_ok=True)
            continue_path.write_text(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), encoding="utf-8")
            self.log_received.emit(f"Sinal manual enviado para {self.manual_wait_dir.name}.")
            return

        if not self.control_dir or not self.manual_wait_company_key:
            self.log_received.emit("Nenhuma etapa manual aguardando liberacao.")
            return

        empresa = self.company_data_by_cnpj.get(self.manual_wait_company_key)
        if not empresa:
            self.log_received.emit("Empresa da etapa manual nao encontrada na memoria do controller.")
            return

        pasta_empresa = nome_pasta_empresa_por_dados(empresa)
        continue_path = self.control_dir / pasta_empresa / "continue.signal"
        continue_path.parent.mkdir(parents=True, exist_ok=True)
        continue_path.write_text(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), encoding="utf-8")
        self.log_received.emit(f"Sinal manual enviado para {empresa.get('codigo', '')} - {empresa.get('razao_social', '')}.")

    def resolve_orchestrator_launch(self) -> dict[str, object]:
        if not is_frozen_app():
            return {
                "program": sys.executable,
                "args": ["-u", str(self.orchestrator_path)],
                "workdir": self.project_dir,
            }

        app_dir = Path(sys.executable).resolve().parent
        candidates = [
            app_dir / APP_ORCHESTRATOR_EXE_NAME,
            app_dir / "backend" / APP_ORCHESTRATOR_EXE_NAME,
            app_dir / "bin" / APP_ORCHESTRATOR_EXE_NAME,
        ]
        for candidate in candidates:
            if candidate.exists():
                return {
                    "program": str(candidate),
                    "args": [],
                    "workdir": candidate.parent,
                }

        raise FileNotFoundError(
            "Executavel do orquestrador nao encontrado no modo instalado. "
            f"Esperado: {APP_ORCHESTRATOR_EXE_NAME}"
        )

    def _cleanup_control_dir(self) -> None:
        self.manual_wait_company_key = None
        self.manual_wait_dir = None
        self.manual_wait_token = None
        self.manual_wait_changed.emit({"active": False})
        if self.control_dir and self.control_dir.exists():
            shutil.rmtree(self.control_dir, ignore_errors=True)
        self.control_dir = None

    @staticmethod
    def _now_text() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _flush_remaining_buffer(self, is_stderr: bool) -> None:
        buffer = self._stderr_buffer if is_stderr else self._stdout_buffer
        if not buffer:
            return

        text = f"[stderr] {buffer}" if is_stderr else buffer
        self.log_received.emit(text)
        if not is_stderr:
            self.handle_event_line(buffer)

        if is_stderr:
            self._stderr_buffer = ""
        else:
            self._stdout_buffer = ""

    def sync_manual_wait(self) -> None:
        if not self.control_dir or not self.control_dir.exists():
            self._clear_manual_wait_state()
            return

        wait_files = sorted(
            self.control_dir.rglob("manual_wait.json"),
            key=lambda path: path.stat().st_mtime,
            reverse=True,
        )
        if not wait_files:
            self._clear_manual_wait_state()
            return

        wait_path = wait_files[0]
        try:
            payload = json.loads(wait_path.read_text(encoding="utf-8"))
        except Exception as exc:
            self.log_received.emit(f"Falha ao ler estado de etapa manual: {exc}")
            return

        empresa_pasta = str(payload.get("empresa_pasta", "")).strip() or wait_path.parent.name
        empresa = self._resolve_company_by_folder(empresa_pasta)
        tag = str(payload.get("tag", "")).strip()
        contexto = str(payload.get("contexto", "")).strip()
        timestamp = str(payload.get("timestamp", "")).strip()
        etapa = "Aguardando captcha" if tag == "MANUAL-LOGIN" else "Livros / Manual"
        token = "|".join([str(wait_path), tag, contexto, timestamp])

        self.manual_wait_dir = wait_path.parent
        self.manual_wait_token = token
        if empresa:
            self.manual_wait_company_key = only_digits(empresa.get("cnpj", ""))

        self.manual_wait_changed.emit(
            {
                "active": True,
                "cnpj": self.manual_wait_company_key or "",
                "codigo": str((empresa or {}).get("codigo", "")),
                "razao_social": str((empresa or {}).get("razao_social", "")) or empresa_pasta,
                "etapa": etapa,
                "mensagem": contexto,
                "tag": tag,
            }
        )

    def _clear_manual_wait_state(self) -> None:
        if not self.manual_wait_dir and not self.manual_wait_company_key and not self.manual_wait_token:
            return

        self.manual_wait_dir = None
        self.manual_wait_company_key = None
        self.manual_wait_token = None
        self.manual_wait_changed.emit({"active": False})

    def _resolve_company_by_folder(self, empresa_pasta: str) -> dict | None:
        empresa_pasta = (empresa_pasta or "").strip()
        if not empresa_pasta:
            return None

        for empresa in self.company_data_by_cnpj.values():
            if nome_pasta_empresa_por_dados(empresa) == empresa_pasta:
                return empresa
        return None
