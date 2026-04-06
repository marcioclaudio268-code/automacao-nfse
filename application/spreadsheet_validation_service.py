from __future__ import annotations

from collections import Counter
import re

from core.config_runtime import RuntimeConfig
from ui.models import ValidationIssue, ValidationResult


def only_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


class SpreadsheetValidationService:
    def validate(self, cfg: RuntimeConfig) -> ValidationResult:
        from orquestrador_empresas import carregar_empresas

        empresas = carregar_empresas(str(cfg.empresas_arquivo))
        result = ValidationResult(companies=empresas)

        if not empresas:
            result.issues.append(
                ValidationIssue("ERROR", "SEM_EMPRESAS", "Nenhuma empresa valida foi encontrada na planilha.")
            )
            return result

        cnpjs: list[str] = []
        for i, empresa in enumerate(empresas, start=1):
            cnpj = only_digits(empresa.get("cnpj", ""))
            senha = (empresa.get("senha_prefeitura") or "").strip()

            if not senha:
                result.issues.append(
                    ValidationIssue(
                        "WARNING",
                        "SEM_SENHA",
                        "Empresa sem senha da prefeitura.",
                        row_number=i,
                        cnpj=cnpj,
                    )
                )

            if cnpj:
                cnpjs.append(cnpj)

        duplicados = [cnpj for cnpj, qtd in Counter(cnpjs).items() if qtd > 1]
        for cnpj in duplicados:
            result.issues.append(
                ValidationIssue("ERROR", "CNPJ_DUPLICADO", f"CNPJ duplicado encontrado: {cnpj}", cnpj=cnpj)
            )

        return result
