from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class EtapaExecucao(str, Enum):
    AGUARDANDO = "AGUARDANDO"
    PREPARANDO = "PREPARANDO"
    LOGIN = "LOGIN"
    AGUARDANDO_CAPTCHA = "AGUARDANDO_CAPTCHA"
    PRESTADOS = "PRESTADOS"
    TOMADOS = "TOMADOS"
    LIVROS_MANUAL = "LIVROS_MANUAL"
    CONCLUIDO = "CONCLUIDO"
    ERRO = "ERRO"


class ResultadoFinal(str, Enum):
    AGUARDANDO = "AGUARDANDO"
    EM_EXECUCAO = "EM_EXECUCAO"
    SUCESSO = "SUCESSO"
    SUCESSO_SEM_COMPETENCIA = "SUCESSO_SEM_COMPETENCIA"
    SUCESSO_SEM_SERVICOS = "SUCESSO_SEM_SERVICOS"
    REVISAO_MANUAL = "REVISAO_MANUAL"
    FALHA = "FALHA"
    INTERROMPIDO = "INTERROMPIDO"


class AcaoRecomendada(str, Enum):
    NENHUMA = ""
    CONFERIR_SENHA = "CONFERIR_SENHA"
    REPROCESSAR = "REPROCESSAR"
    ABRIR_DEBUG = "ABRIR_DEBUG"


@dataclass
class CompanyRowState:
    codigo: str
    razao_social: str
    cnpj: str
    resultado: str = ResultadoFinal.AGUARDANDO
    etapa: str = EtapaExecucao.AGUARDANDO
    ultima_mensagem: str = ""
    tentativas: str = "0"
    inicio: str = ""
    fim: str = ""
    acao_recomendada: str = AcaoRecomendada.NENHUMA


@dataclass
class ValidationIssue:
    severity: str
    code: str
    message: str
    row_number: int | None = None
    cnpj: str = ""


@dataclass
class ValidationResult:
    companies: list[dict] = field(default_factory=list)
    issues: list[ValidationIssue] = field(default_factory=list)

    @property
    def total_validas(self) -> int:
        return len(self.companies)
