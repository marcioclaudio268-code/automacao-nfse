from types import SimpleNamespace

import orquestrador_empresas as oq


def _empresa_dummy():
    return {
        "codigo": "1",
        "razao_social": "Empresa Teste",
        "cnpj": "12.345.678/0001-90",
        "segmento": "SERVICOS",
        "senha_prefeitura": "segredo",
    }


def test_executar_empresa_rc0_com_alertas(monkeypatch):
    monkeypatch.setattr(
        oq.subprocess,
        "run",
        lambda *args, **kwargs: SimpleNamespace(returncode=0, stdout="execucao iniciada\nexecucao finalizada", stderr=""),
    )

    res = oq.executar_empresa(_empresa_dummy())

    assert res["status"] == "SUCESSO"
    assert res["motivo"] == "OK"


def test_executar_empresa_rc_empresa_multipla(monkeypatch):
    monkeypatch.setattr(
        oq.subprocess,
        "run",
        lambda *args, **kwargs: SimpleNamespace(returncode=oq.EXIT_CODE_EMPRESA_MULTIPLA, stdout="", stderr=""),
    )

    res = oq.executar_empresa(_empresa_dummy())

    assert res["status"] == "REVISAO_MANUAL"
    assert "EMPRESA_MULTIPLA" in res["motivo"]


def test_executar_empresa_rc_multi_cadastro_revisao_manual(monkeypatch):
    monkeypatch.setattr(
        oq.subprocess,
        "run",
        lambda *args, **kwargs: SimpleNamespace(returncode=oq.EXIT_CODE_MULTI_CADASTRO_REVISAO_MANUAL, stdout="", stderr=""),
    )

    res = oq.executar_empresa(_empresa_dummy())

    assert res["status"] == "REVISAO_MANUAL"
    assert "MULTIPLOS_CADASTROS" in res["motivo"]


def test_main_registra_revisao_manual_multiplas_no_report(monkeypatch, tmp_path):
    csv_path = tmp_path / "empresas.xlsx"
    csv_path.write_text("stub", encoding="utf-8")

    monkeypatch.setattr(oq, "CSV_EMPRESAS", str(csv_path))
    monkeypatch.setattr(oq, "CONTINUAR_DE_ONDE_PAROU", False)
    monkeypatch.setattr(oq, "USAR_CHECKPOINT", False)

    monkeypatch.setattr(oq, "carregar_empresas", lambda _: [_empresa_dummy()])
    monkeypatch.setattr(
        oq,
        "executar_empresa",
        lambda _: {
            "status": "REVISAO_MANUAL",
            "motivo": "EMPRESA_MULTIPLA",
            "tentativas": 1,
            "inicio": oq.datetime.now(),
            "fim": oq.datetime.now(),
        },
    )

    rows = []
    monkeypatch.setattr(oq, "append_report_row", lambda *args, **kwargs: rows.append((args, kwargs)))
    monkeypatch.setattr(oq, "resetar_report", lambda *_: None)
    monkeypatch.setattr(oq, "garantir_report_com_header", lambda *_: None)

    oq.main()

    assert len(rows) == 1
    assert rows[0][0][1]["resultado"]["status"] == "REVISAO_MANUAL"


def test_main_fluxo_pulado_nao_usa_variavel_res_inexistente(monkeypatch, tmp_path):
    csv_path = tmp_path / "empresas.xlsx"
    csv_path.write_text("stub", encoding="utf-8")

    emp = _empresa_dummy()
    cnpj_limpo = "12345678000190"

    monkeypatch.setattr(oq, "CSV_EMPRESAS", str(csv_path))
    monkeypatch.setattr(oq, "CONTINUAR_DE_ONDE_PAROU", True)
    monkeypatch.setattr(oq, "USAR_CHECKPOINT", False)

    monkeypatch.setattr(oq, "carregar_empresas", lambda _: [emp])
    monkeypatch.setattr(oq, "carregar_report_existente", lambda *_: {cnpj_limpo})
    monkeypatch.setattr(oq, "carregar_rows_report_existente", lambda *_: [])

    rows = []
    monkeypatch.setattr(oq, "append_report_row", lambda *args, **kwargs: rows.append((args, kwargs)))
    monkeypatch.setattr(oq, "garantir_report_com_header", lambda *_: None)
    monkeypatch.setattr(oq, "resetar_report", lambda *_: None)

    oq.main()

    assert len(rows) == 1
