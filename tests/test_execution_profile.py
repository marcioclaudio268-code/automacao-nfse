import pytest

from datetime import datetime

from core.config_runtime import ExecutionProfile, competencias_alvo, validate_apuracao_referencia


def test_execution_profile_default_matches_current_behavior():
    profile = ExecutionProfile()

    profile.validate()

    assert profile.is_default() is True
    assert profile.pausa_manual_final is True


def test_execution_profile_auto_mode_disables_final_manual_pause():
    profile = ExecutionProfile(modo_automatico=True)

    profile.validate()

    assert profile.pausa_manual_final is False
    assert profile.as_env()["MODO_AUTOMATICO"] == "1"
    assert profile.as_env()["PAUSA_MANUAL_FINAL"] == "0"


def test_execution_profile_rejects_iss_without_books():
    profile = ExecutionProfile(executar_livros=False, executar_iss=True)

    with pytest.raises(ValueError, match="ISS depende de Livros"):
        profile.validate()


def test_execution_profile_requires_at_least_one_module():
    profile = ExecutionProfile(executar_prestados=False, executar_tomados=False)

    with pytest.raises(ValueError, match="Selecione pelo menos um modulo"):
        profile.validate()


def test_execution_profile_requires_xml_or_books():
    profile = ExecutionProfile(executar_xml=False, executar_livros=False, executar_iss=False, pausa_manual_final=False)

    with pytest.raises(ValueError, match="Selecione pelo menos uma saida"):
        profile.validate()


def test_execution_profile_apurar_completo_forca_xml_e_desliga_fechamento():
    profile = ExecutionProfile(
        executar_xml=False,
        executar_livros=True,
        executar_iss=True,
        pausa_manual_final=True,
        apurar_completo=True,
    )

    profile.validate()

    assert profile.executar_xml is True
    assert profile.executar_livros is False
    assert profile.executar_iss is False
    assert profile.pausa_manual_final is False
    assert profile.as_env()["APURAR_COMPLETO"] == "1"


def test_validate_apuracao_referencia_rejeita_mes_posterior_ao_atual():
    with pytest.raises(ValueError, match="Use no maximo 03/2026"):
        validate_apuracao_referencia("04/2026", now=datetime(2026, 3, 17))


def test_competencias_alvo_apurar_completo_usa_ano_da_competencia_alvo():
    assert competencias_alvo("01/2026", apurar_completo=True) == [
        (2025, 1),
        (2025, 2),
        (2025, 3),
        (2025, 4),
        (2025, 5),
        (2025, 6),
        (2025, 7),
        (2025, 8),
        (2025, 9),
        (2025, 10),
        (2025, 11),
        (2025, 12),
    ]
