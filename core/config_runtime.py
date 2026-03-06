from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re


_MM_AAAA_RE = re.compile(r"^(0[1-9]|1[0-2])/\d{4}$")


@dataclass(slots=True)
class RuntimeConfig:
    empresas_arquivo: Path
    apuracao_referencia: str
    output_base_dir: Path
    continuar_de_onde_parou: bool = True
    usar_checkpoint: bool = True
    login_wait_seconds: int = 120
    timeout_processo_main: int = 1800

    def validate(self) -> None:
        if not self.empresas_arquivo.exists():
            raise FileNotFoundError(f"Planilha não encontrada: {self.empresas_arquivo}")

        if self.empresas_arquivo.suffix.lower() not in {".xlsx", ".csv"}:
            raise ValueError("A planilha deve ser .xlsx ou .csv")

        if not is_mm_aaaa(self.apuracao_referencia):
            raise ValueError("Mês de apuração deve estar no formato MM/AAAA")

        self.output_base_dir.mkdir(parents=True, exist_ok=True)


def is_mm_aaaa(value: str) -> bool:
    return bool(_MM_AAAA_RE.fullmatch((value or "").strip()))


def mes_atual_str() -> str:
    return datetime.now().strftime("%m/%Y")