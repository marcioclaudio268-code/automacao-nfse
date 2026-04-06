from __future__ import annotations

import csv
import re
import tempfile
from pathlib import Path
from uuid import uuid4


MANUAL_COMPANY_CODE = "MANUAL"
MANUAL_COMPANY_SEGMENT = "MANUAL"
MANUAL_COMPANY_HEADER = [
    "Codigo",
    "Razao Social",
    "CNPJ",
    "Segmento",
    "Senha Prefeitura",
]


def only_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def has_manual_credentials(cnpj: str, senha: str) -> bool:
    return bool(only_digits(cnpj) and (senha or "").strip())


def has_partial_manual_credentials(cnpj: str, senha: str) -> bool:
    cnpj_preenchido = bool(only_digits(cnpj))
    senha_preenchida = bool((senha or "").strip())
    return cnpj_preenchido != senha_preenchida


def validate_manual_credentials(cnpj: str, senha: str) -> tuple[str, str]:
    cnpj_digits = only_digits(cnpj)
    senha_limpa = (senha or "").strip()

    if has_partial_manual_credentials(cnpj, senha):
        raise ValueError(
            "Para usar o modo manual, preencha CNPJ e Senha. "
            "Se quiser usar a planilha, deixe ambos vazios."
        )

    if not cnpj_digits or not senha_limpa:
        raise ValueError("CNPJ e Senha manuais nao informados.")

    if len(cnpj_digits) != 14:
        raise ValueError("CNPJ manual deve conter 14 digitos.")

    return cnpj_digits, senha_limpa


def build_manual_company(cnpj: str, senha: str) -> dict[str, str]:
    cnpj_digits, senha_limpa = validate_manual_credentials(cnpj, senha)
    return {
        "codigo": MANUAL_COMPANY_CODE,
        "razao_social": f"CNPJ_{cnpj_digits}",
        "cnpj": cnpj_digits,
        "segmento": MANUAL_COMPANY_SEGMENT,
        "senha_prefeitura": senha_limpa,
    }


def write_manual_company_csv(empresa: dict[str, str], base_dir: Path | None = None) -> Path:
    runtime_dir = base_dir if base_dir is not None else Path(tempfile.gettempdir()) / "AutomacaoNFSe" / "runtime_inputs"
    runtime_dir.mkdir(parents=True, exist_ok=True)

    csv_path = runtime_dir / f"empresa_manual_{uuid4().hex}.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, delimiter=";")
        writer.writerow(MANUAL_COMPANY_HEADER)
        writer.writerow(
            [
                empresa.get("codigo", ""),
                empresa.get("razao_social", ""),
                empresa.get("cnpj", ""),
                empresa.get("segmento", ""),
                empresa.get("senha_prefeitura", ""),
            ]
        )

    return csv_path
