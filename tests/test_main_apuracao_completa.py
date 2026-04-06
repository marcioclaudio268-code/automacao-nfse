import pytest

import main


def test_executar_fluxo_apuracao_completa_reabre_lista_e_processa_competencias(monkeypatch):
    logs = []
    prestados = []
    tomados = []

    monkeypatch.setattr(main, "runtime_competencias_alvo", lambda *_args, **_kwargs: [(2025, 1), (2025, 2), (2025, 3)])
    monkeypatch.setattr(main, "EXECUTAR_PRESTADOS", True)
    monkeypatch.setattr(main, "EXECUTAR_TOMADOS", True)
    monkeypatch.setattr(main, "PRESTADOS_SEM_MODULO", False)
    monkeypatch.setattr(main, "PERFIL_EXECUCAO_ATIVO", True)
    monkeypatch.setattr(main, "descrever_perfil_execucao", lambda: "Perfil de teste")
    monkeypatch.setattr(main, "registrar_apuracao_completa", lambda ref, etapa, status, mensagem="": logs.append((ref, etapa, status, mensagem)))
    monkeypatch.setattr(
        main,
        "executar_etapa_servicos_prestados_apuracao_completa",
        lambda ano: prestados.append(ano) or {(2025, 1), (2025, 2), (2025, 3)},
    )
    monkeypatch.setattr(main, "executar_etapa_servicos_tomados", lambda ano, mes: tomados.append((ano, mes)) or "OK")
    monkeypatch.setattr(main, "sleep", lambda *_: None)

    main.executar_fluxo_apuracao_completa("01/2026")

    assert prestados == [2025]
    assert tomados == [(2025, 3), (2025, 2), (2025, 1)]
    assert ("01/2025", "PRESTADOS", "OK", "") in logs
    assert ("03/2025", "TOMADOS", "OK", "") in logs


def test_executar_fluxo_apuracao_completa_eleva_revisao_manual_quando_alguma_competencia_falha(monkeypatch):
    logs = []
    tomados = []

    monkeypatch.setattr(main, "runtime_competencias_alvo", lambda *_args, **_kwargs: [(2025, 1), (2025, 2)])
    monkeypatch.setattr(main, "EXECUTAR_PRESTADOS", True)
    monkeypatch.setattr(main, "EXECUTAR_TOMADOS", True)
    monkeypatch.setattr(main, "PRESTADOS_SEM_MODULO", False)
    monkeypatch.setattr(main, "PERFIL_EXECUCAO_ATIVO", True)
    monkeypatch.setattr(main, "descrever_perfil_execucao", lambda: "Perfil de teste")
    monkeypatch.setattr(main, "registrar_apuracao_completa", lambda ref, etapa, status, mensagem="": logs.append((ref, etapa, status, mensagem)))
    monkeypatch.setattr(main, "executar_etapa_servicos_prestados_apuracao_completa", lambda *_args: {(2025, 1), (2025, 2)})

    def fake_tomados(ano, mes):
        tomados.append((ano, mes))
        if mes == 2:
            raise RuntimeError("timeout portal")
        return "OK"

    monkeypatch.setattr(main, "executar_etapa_servicos_tomados", fake_tomados)
    monkeypatch.setattr(main, "sleep", lambda *_: None)

    with pytest.raises(RuntimeError, match=main.MSG_APURACAO_COMPLETA_REVISAO_MANUAL):
        main.executar_fluxo_apuracao_completa("01/2026")

    assert tomados == [(2025, 2), (2025, 1)]
    assert any(ref == "02/2025" and etapa == "TOMADOS" and status == "ERRO" for ref, etapa, status, _ in logs)
