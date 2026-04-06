from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import csv
import re


@dataclass(slots=True)
class ParsedEvent:
    kind: str
    raw: str
    codigo: str = ""
    razao: str = ""
    etapa: str = ""
    mensagem: str = ""


_RE_EMPRESA = re.compile(r"^Empresa\s+(?P<codigo>.+?)\s+-\s+(?P<razao>.+)$", re.IGNORECASE)


def parse_log_line(line: str) -> ParsedEvent:
    text = (line or "").strip()
    if not text:
        return ParsedEvent(kind="empty", raw=line)

    m = _RE_EMPRESA.match(text)
    if m:
        return ParsedEvent(
            kind="empresa_atual",
            raw=line,
            codigo=m.group("codigo").strip(),
            razao=m.group("razao").strip(),
            etapa="Preparando empresa",
            mensagem=text,
        )

    lower = text.lower()

    if text.startswith("[stderr]"):
        return ParsedEvent(kind="stderr", raw=line, etapa="Erro", mensagem=text)

    if text.startswith("[MANUAL-LOGIN]"):
        return ParsedEvent(kind="etapa", raw=line, etapa="Aguardando captcha", mensagem=text)

    if text.startswith("[MANUAL-FINAL]") or text.startswith("[MANUAL]"):
        if "CONTINUANDO_EXECUCAO_MANUAL" in text:
            return ParsedEvent(kind="manual_resumed", raw=line, etapa="Livros / Manual", mensagem=text)
        return ParsedEvent(kind="manual_wait", raw=line, etapa="Livros / Manual", mensagem=text)

    if "worker_interrompido" in lower or "interrompido pelo usuario" in lower:
        return ParsedEvent(kind="interrompido", raw=line, etapa="Interrompido", mensagem=text)

    if "credencial invalida" in lower or "credencial inválida" in lower:
        return ParsedEvent(kind="erro", raw=line, etapa="Login", mensagem=text)

    if "timeout" in lower:
        return ParsedEvent(kind="erro", raw=line, etapa="Erro", mensagem=text)

    if "falha tentativa" in lower or "[tomados] erro" in lower or "[prestados] erro" in lower:
        return ParsedEvent(kind="erro", raw=line, etapa="Erro", mensagem=text)

    if lower.startswith("erro") or "erro inesperado" in lower:
        return ParsedEvent(kind="erro", raw=line, etapa="Erro", mensagem=text)

    if "processamento finalizado" in lower or text == "Processo finalizado.":
        return ParsedEvent(kind="final", raw=line, etapa="Concluido", mensagem=text)

    if "etapa 1:" in lower or "login automatico" in lower:
        return ParsedEvent(kind="etapa", raw=line, etapa="Login", mensagem=text)

    if "captcha" in lower:
        return ParsedEvent(kind="etapa", raw=line, etapa="Aguardando captcha", mensagem=text)

    if "tomados" in lower:
        return ParsedEvent(kind="etapa", raw=line, etapa="Tomados", mensagem=text)

    if "prestados" in lower:
        return ParsedEvent(kind="etapa", raw=line, etapa="Prestados", mensagem=text)

    if "livro" in lower or "guia" in lower or "manual" in lower:
        return ParsedEvent(kind="etapa", raw=line, etapa="Livros / Manual", mensagem=text)

    return ParsedEvent(kind="log", raw=line, mensagem=text)


def load_report_rows(report_path: Path) -> list[dict]:
    if not report_path.exists():
        return []

    rows: list[dict] = []
    with report_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            rows.append(row)
    return rows
