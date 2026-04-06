from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
import re


_MM_AAAA_RE = re.compile(r"^(0[1-9]|1[0-2])/\d{4}$")


@dataclass(slots=True)
class ExecutionProfile:
    executar_prestados: bool = True
    executar_tomados: bool = True
    executar_xml: bool = True
    executar_livros: bool = True
    executar_iss: bool = True
    pausa_manual_final: bool = True
    modo_automatico: bool = False
    apurar_completo: bool = False

    def normalize(self) -> None:
        if self.apurar_completo:
            self.executar_xml = True
            self.executar_livros = False
            self.executar_iss = False
            self.pausa_manual_final = False

        if self.modo_automatico:
            self.pausa_manual_final = False

    def validate(self) -> None:
        self.normalize()

        if not (self.executar_prestados or self.executar_tomados):
            raise ValueError("Selecione pelo menos um modulo: Prestados ou Tomados.")

        if not (self.executar_xml or self.executar_livros):
            raise ValueError("Selecione pelo menos uma saida: XML ou Livros.")

        if self.executar_iss and not self.executar_livros:
            raise ValueError("ISS depende de Livros.")

        if self.pausa_manual_final and not self.executar_livros:
            raise ValueError("Pausa manual no final depende de Livros.")

    def is_default(self) -> bool:
        normalized = ExecutionProfile(
            executar_prestados=self.executar_prestados,
            executar_tomados=self.executar_tomados,
            executar_xml=self.executar_xml,
            executar_livros=self.executar_livros,
            executar_iss=self.executar_iss,
            pausa_manual_final=self.pausa_manual_final,
            modo_automatico=self.modo_automatico,
            apurar_completo=self.apurar_completo,
        )
        normalized.normalize()
        return normalized == ExecutionProfile()

    def as_env(self) -> dict[str, str]:
        self.validate()
        return {
            "PERFIL_EXECUCAO_ATIVO": "1",
            "EXECUTAR_PRESTADOS": "1" if self.executar_prestados else "0",
            "EXECUTAR_TOMADOS": "1" if self.executar_tomados else "0",
            "EXECUTAR_XML": "1" if self.executar_xml else "0",
            "EXECUTAR_LIVROS": "1" if self.executar_livros else "0",
            "EXECUTAR_ISS": "1" if self.executar_iss else "0",
            "PAUSA_MANUAL_FINAL": "1" if self.pausa_manual_final else "0",
            "MODO_AUTOMATICO": "1" if self.modo_automatico else "0",
            "APURAR_COMPLETO": "1" if self.apurar_completo else "0",
        }


@dataclass(slots=True)
class RuntimeConfig:
    empresas_arquivo: Path
    apuracao_referencia: str
    output_base_dir: Path
    execution_profile: ExecutionProfile = field(default_factory=ExecutionProfile)
    continuar_de_onde_parou: bool = True
    usar_checkpoint: bool = True
    login_wait_seconds: int = 120
    timeout_processo_main: int = 1800

    def validate(self) -> None:
        if not self.empresas_arquivo.exists():
            raise FileNotFoundError(f"Planilha nao encontrada: {self.empresas_arquivo}")

        if self.empresas_arquivo.suffix.lower() not in {".xlsx", ".csv"}:
            raise ValueError("A planilha deve ser .xlsx ou .csv")

        if not is_mm_aaaa(self.apuracao_referencia):
            raise ValueError("Mes de apuracao deve estar no formato MM/AAAA")

        validate_apuracao_referencia(self.apuracao_referencia)

        if not str(self.output_base_dir).strip():
            raise ValueError("Pasta base de saida nao informada.")

        if self.login_wait_seconds <= 0:
            raise ValueError("login_wait_seconds deve ser maior que zero")

        if self.timeout_processo_main <= 0:
            raise ValueError("timeout_processo_main deve ser maior que zero")

        self.execution_profile.validate()
        self.output_base_dir.mkdir(parents=True, exist_ok=True)


def is_mm_aaaa(value: str) -> bool:
    return bool(_MM_AAAA_RE.fullmatch((value or "").strip()))


def parse_mm_aaaa(value: str) -> tuple[int, int]:
    text = (value or "").strip()
    if not is_mm_aaaa(text):
        raise ValueError("Mes de apuracao deve estar no formato MM/AAAA")
    mes, ano = text.split("/")
    return int(mes), int(ano)


def competencia_alvo(apuracao_referencia: str) -> tuple[int, int]:
    mes, ano = parse_mm_aaaa(apuracao_referencia)
    if mes == 1:
        return ano - 1, 12
    return ano, mes - 1


def validate_apuracao_referencia(apuracao_referencia: str, now: datetime | None = None) -> None:
    mes_apuracao, ano_apuracao = parse_mm_aaaa(apuracao_referencia)
    now = now or datetime.now()
    if (ano_apuracao, mes_apuracao) > (now.year, now.month):
        raise ValueError(
            "Mes de apuracao nao pode ser posterior ao mes atual. "
            f"Use no maximo {now.strftime('%m/%Y')}."
        )


def competencias_alvo(apuracao_referencia: str, apurar_completo: bool = False) -> list[tuple[int, int]]:
    validate_apuracao_referencia(apuracao_referencia)
    ano_alvo, mes_alvo = competencia_alvo(apuracao_referencia)
    if not apurar_completo:
        return [(ano_alvo, mes_alvo)]
    return [(ano_alvo, mes) for mes in range(1, mes_alvo + 1)]


def competencia_alvo_dir_name(apuracao_referencia: str) -> str:
    ano, mes = competencia_alvo(apuracao_referencia)
    return f"{mes:02d}.{ano}"


def mes_atual_str() -> str:
    return datetime.now().strftime("%m/%Y")
