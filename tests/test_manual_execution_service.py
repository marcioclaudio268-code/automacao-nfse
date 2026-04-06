import pytest

import orquestrador_empresas as oq
from application.manual_execution_service import (
    build_manual_company,
    has_manual_credentials,
    has_partial_manual_credentials,
    write_manual_company_csv,
)


def test_build_manual_company_normalizes_cnpj_and_defaults():
    empresa = build_manual_company("12.345.678/0001-90", "  segredo  ")

    assert empresa == {
        "codigo": "MANUAL",
        "razao_social": "CNPJ_12345678000190",
        "cnpj": "12345678000190",
        "segmento": "MANUAL",
        "senha_prefeitura": "segredo",
    }


def test_build_manual_company_requires_both_fields():
    with pytest.raises(ValueError, match="preencha CNPJ e Senha"):
        build_manual_company("12.345.678/0001-90", "")


def test_build_manual_company_requires_valid_cnpj_length():
    with pytest.raises(ValueError, match="14 digitos"):
        build_manual_company("123", "segredo")


def test_manual_credential_flags():
    assert has_manual_credentials("12.345.678/0001-90", "segredo") is True
    assert has_manual_credentials("", "segredo") is False
    assert has_partial_manual_credentials("12.345.678/0001-90", "") is True
    assert has_partial_manual_credentials("", "") is False


def test_write_manual_company_csv_is_compatible_with_orquestrador_loader(tmp_path):
    empresa = build_manual_company("12.345.678/0001-90", "segredo")
    csv_path = write_manual_company_csv(empresa, base_dir=tmp_path)

    carregadas = oq.carregar_empresas(str(csv_path))

    assert carregadas == [{**empresa, "indice_lista": 1, "linha_planilha": 2}]
