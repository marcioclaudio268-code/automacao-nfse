import pytest

import main


@pytest.fixture(autouse=True)
def _resetar_estado_cadastro():
    main.resetar_contexto_cadastro()
    yield
    main.resetar_contexto_cadastro()


def test_tela_cadastros_relacionados_aberta_detecta_grid_e_botao(monkeypatch):
    class _Driver:
        page_source = "Selecao do Contribuinte Cadastros Relacionados"

        def find_elements(self, by, value):
            if value == "_imagebutton1":
                return [object()]
            if value == "_grid1Selected":
                return [object()]
            return []

    monkeypatch.setattr(main, "driver", _Driver())

    assert main.tela_cadastros_relacionados_aberta() is True


def test_coletar_cadastros_relacionados_percorre_todas_as_paginas(monkeypatch):
    state = {"page": 1}
    pages = {1: ["a", "b"], 2: ["c"]}

    monkeypatch.setattr(main, "tela_cadastros_relacionados_aberta", lambda: True)
    monkeypatch.setattr(main, "pagina_atual_cadastros_relacionados", lambda: state["page"])
    monkeypatch.setattr(main, "_linhas_cadastros_relacionados_page", lambda: pages[state["page"]])
    monkeypatch.setattr(
        main,
        "_extrair_cadastro_relacionado_da_linha",
        lambda tr, page, ordem, row_index: {
            "ordem": ordem,
            "page": page,
            "row_index": row_index,
            "row_key": f"{page}-{row_index}",
            "valor": str(ordem),
            "ccm": "619824",
            "identificacao": f"cadastro-{tr}",
        },
    )

    def _next():
        if state["page"] == 1:
            state["page"] = 2
            return True
        return False

    monkeypatch.setattr(main, "ir_para_proxima_pagina_cadastros_relacionados", _next)

    cadastros = main.coletar_cadastros_relacionados()

    assert [item["ordem"] for item in cadastros] == [1, 2, 3]
    assert [item["page"] for item in cadastros] == [1, 1, 2]


def test_agregar_resultados_multicadastro_mistura_sucesso_e_sem_competencia_retorna_sucesso():
    status, motivo = main.agregar_resultados_multicadastro(
        [
            {"status": "SUCESSO", "motivo": "OK", "cadastro": {"ordem": 1, "ccm": "1"}},
            {"status": "SUCESSO_SEM_COMPETENCIA", "motivo": main.MSG_SEM_COMPETENCIA, "cadastro": {"ordem": 2, "ccm": "2"}},
        ]
    )

    assert status == "SUCESSO"
    assert motivo == "OK"


def test_executar_fluxo_multiplos_cadastros_levanta_revisao_manual_quando_um_falha(monkeypatch):
    cadastros = [
        {"ordem": 1, "page": 1, "row_index": 1, "row_key": "1", "ccm": "619824", "identificacao": "cad-1"},
        {"ordem": 2, "page": 1, "row_index": 2, "row_key": "2", "ccm": "619824", "identificacao": "cad-2"},
    ]

    monkeypatch.setattr(main, "tela_cadastros_relacionados_aberta", lambda: True)
    monkeypatch.setattr(main, "coletar_cadastros_relacionados", lambda: cadastros)
    monkeypatch.setattr(main, "retornar_para_cadastros_relacionados", lambda timeout=0: None)
    monkeypatch.setattr(main, "selecionar_cadastro_relacionado", lambda cadastro, timeout=0: cadastro)
    monkeypatch.setattr(
        main,
        "executar_fluxo_cadastro",
        lambda cadastro, *_: {
            "cadastro": cadastro,
            "status": "FALHA" if cadastro["ordem"] == 2 else "SUCESSO",
            "motivo": "Erro no cadastro 2" if cadastro["ordem"] == 2 else "OK",
        },
    )

    try:
        main.executar_fluxo_multiplos_cadastros("03/2026", 2026, 2)
    except RuntimeError as exc:
        assert main.MSG_MULTI_CADASTRO_REVISAO_MANUAL in str(exc)
        assert "CCM_619824__CAD_02" in str(exc)
    else:
        raise AssertionError("Era esperado RuntimeError de revisao manual.")


def test_executar_fluxo_multiplos_cadastros_todos_sem_competencia_repassa_status(monkeypatch):
    cadastros = [
        {"ordem": 1, "page": 1, "row_index": 1, "row_key": "1", "ccm": "619824", "identificacao": "cad-1"},
        {"ordem": 2, "page": 1, "row_index": 2, "row_key": "2", "ccm": "619824", "identificacao": "cad-2"},
    ]

    monkeypatch.setattr(main, "tela_cadastros_relacionados_aberta", lambda: True)
    monkeypatch.setattr(main, "coletar_cadastros_relacionados", lambda: cadastros)
    monkeypatch.setattr(main, "retornar_para_cadastros_relacionados", lambda timeout=0: None)
    monkeypatch.setattr(main, "selecionar_cadastro_relacionado", lambda cadastro, timeout=0: cadastro)
    monkeypatch.setattr(
        main,
        "executar_fluxo_cadastro",
        lambda cadastro, *_: {
            "cadastro": cadastro,
            "status": "SUCESSO_SEM_COMPETENCIA",
            "motivo": main.MSG_SEM_COMPETENCIA,
        },
    )

    try:
        main.executar_fluxo_multiplos_cadastros("03/2026", 2026, 2)
    except RuntimeError as exc:
        assert str(exc) == main.MSG_SEM_COMPETENCIA
    else:
        raise AssertionError("Era esperado RuntimeError de sem competencia.")


def test_clicar_inicio_para_dashboard_retoma_cadastro_atual_quando_inicio_volta_para_cadastros(monkeypatch):
    class _FakeInicioLink:
        text = "Inicio"

        def get_attribute(self, name):
            if name == "title":
                return "Inicio"
            return ""

    class _Driver:
        def __init__(self):
            self.state = "nf"
            self.link = _FakeInicioLink()

        def find_elements(self, by, value):
            if value in {"imgdeclaracaofiscal", "divtxtdeclaracaofiscal"}:
                return [object()] if self.state == "dashboard" else []
            if value == "a.historic-item":
                return [self.link]
            if value == "a.historic-item[title='Início']":
                return [self.link]
            return []

    class _DummyWait:
        def __init__(self, driver, timeout, *args, **kwargs):
            self.driver = driver

        def until(self, condition):
            for _ in range(5):
                result = condition(self.driver)
                if result:
                    return result
            raise AssertionError("timeout")

    driver = _Driver()
    cadastro = {"ordem": 1, "page": 1, "row_index": 1, "row_key": "1", "ccm": "619824", "identificacao": "cad-1"}
    retomadas = []

    monkeypatch.setattr(main, "driver", driver)
    monkeypatch.setattr(main, "WebDriverWait", _DummyWait)
    monkeypatch.setattr(main, "MULTI_CADASTRO_ATIVO", True)
    monkeypatch.setattr(main, "contexto_cadastro_atual", lambda: dict(cadastro))
    monkeypatch.setattr(main, "tela_cadastros_relacionados_aberta", lambda: driver.state == "cadastros")

    def _click(_el):
        driver.state = "cadastros"

    def _selecionar(cadastro_atual, timeout=0):
        retomadas.append((cadastro_atual["ordem"], timeout))
        driver.state = "dashboard"
        return dict(cadastro_atual)

    monkeypatch.setattr(main, "click_robusto", _click)
    monkeypatch.setattr(main, "selecionar_cadastro_relacionado", _selecionar)

    main.clicar_inicio_para_dashboard(timeout=5)

    assert retomadas == [(1, 5)]
    assert driver.state == "dashboard"


def test_recuperar_contexto_cadastro_relacionado_para_declaracao_reseleciona_e_reabre_destino(monkeypatch):
    class _DummyWait:
        def __init__(self, driver, timeout, *args, **kwargs):
            self.driver = driver

        def until(self, condition):
            for _ in range(5):
                result = condition(self.driver)
                if result:
                    return result
            raise AssertionError("timeout")

    state = {"screen": "cadastros"}
    cadastro = {"ordem": 1, "page": 1, "row_index": 1, "row_key": "1", "ccm": "619824", "identificacao": "cad-1"}
    retomadas = []
    aberturas = []
    evidencias = []

    monkeypatch.setattr(main, "driver", object())
    monkeypatch.setattr(main, "WebDriverWait", _DummyWait)
    monkeypatch.setattr(main, "MULTI_CADASTRO_ATIVO", True)
    monkeypatch.setattr(main, "contexto_cadastro_atual", lambda: dict(cadastro))
    monkeypatch.setattr(main, "tela_cadastros_relacionados_aberta", lambda: state["screen"] == "cadastros")
    monkeypatch.setattr(main, "dashboard_aberto", lambda: state["screen"] == "dashboard")
    monkeypatch.setattr(main, "tela_declaracao_fiscal_eletronica_aberta", lambda: state["screen"] == "declaracao")
    monkeypatch.setattr(main, "append_txt", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        main,
        "salvar_evidencia_html",
        lambda contexto, motivo, **kwargs: evidencias.append((contexto, motivo)) or {},
    )

    def _selecionar(cadastro_atual, timeout=0):
        retomadas.append((cadastro_atual["ordem"], timeout))
        state["screen"] = "dashboard"
        return dict(cadastro_atual)

    def _abrir_destino(timeout=0):
        aberturas.append(timeout)
        state["screen"] = "declaracao"

    monkeypatch.setattr(main, "selecionar_cadastro_relacionado", _selecionar)

    resultado = main.recuperar_contexto_cadastro_relacionado(
        "declaracao_fiscal",
        timeout=5,
        abrir_destino=_abrir_destino,
        log_path="dummy.log",
        contexto_evidencia="DECLARACAO_FISCAL_PRESTADOS",
        ano_alvo=2026,
        mes_alvo=2,
        ref_alvo="02/2026",
    )

    assert resultado is True
    assert retomadas == [(1, 5)]
    assert aberturas == [5]
    assert ("DECLARACAO_FISCAL_PRESTADOS", "CADASTROS_RELACIONADOS_REAPARECEU") in evidencias
    assert ("DECLARACAO_FISCAL_PRESTADOS", "SERVICOS_PRESTADOS_NAO_IDENTIFICADO") not in evidencias


def test_garantir_declaracao_fiscal_contexto_bootstrap_tardio_ativa_primeiro_cadastro(monkeypatch):
    class _DummyWait:
        def __init__(self, driver, timeout, *args, **kwargs):
            self.driver = driver

        def until(self, condition):
            for _ in range(5):
                result = condition(self.driver)
                if result:
                    return result
            raise AssertionError("timeout")

    state = {"screen": "cadastros"}
    cadastros = [
        {"ordem": 1, "page": 1, "row_index": 1, "row_key": "1", "valor": "1", "ccm": "619824", "identificacao": "igual"},
        {"ordem": 2, "page": 1, "row_index": 2, "row_key": "2", "valor": "1", "ccm": "619824", "identificacao": "igual"},
    ]
    evidencias = []

    main.resetar_multicadastro_tardio()
    monkeypatch.setattr(main, "driver", object())
    monkeypatch.setattr(main, "WebDriverWait", _DummyWait)
    monkeypatch.setattr(main, "CONTEXTO_CADASTRO_ATUAL", {})
    monkeypatch.setattr(main, "MULTI_CADASTRO_ATIVO", False)
    monkeypatch.setattr(main, "tela_declaracao_fiscal_eletronica_aberta", lambda: state["screen"] == "declaracao")
    monkeypatch.setattr(main, "tela_cadastros_relacionados_aberta", lambda: state["screen"] == "cadastros")
    monkeypatch.setattr(main, "dashboard_aberto", lambda: state["screen"] == "dashboard")
    monkeypatch.setattr(main, "coletar_cadastros_relacionados", lambda: [dict(item) for item in cadastros])
    monkeypatch.setattr(main, "append_txt", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        main,
        "salvar_evidencia_html",
        lambda contexto, motivo, **kwargs: evidencias.append((contexto, motivo, kwargs.get("extra") or {})) or {},
    )

    def _selecionar(cadastro_atual, timeout=0):
        state["screen"] = "dashboard"
        return dict(cadastro_atual)

    def _abrir_declaracao(timeout=0):
        state["screen"] = "declaracao"

    monkeypatch.setattr(main, "selecionar_cadastro_relacionado", _selecionar)
    monkeypatch.setattr(main, "abrir_declaracao_fiscal", _abrir_declaracao)

    resultado = main.garantir_declaracao_fiscal_contexto(
        timeout=5,
        log_path="dummy.log",
        contexto_evidencia="DECLARACAO_FISCAL_TOMADOS",
        ano_alvo=2026,
        mes_alvo=2,
        ref_alvo="02/2026",
    )

    assert resultado is True
    assert main.multicadastro_tardio_ativo() is True
    assert main.contexto_cadastro_atual()["ordem"] == 1
    assert any(
        contexto == "DECLARACAO_FISCAL_TOMADOS"
        and motivo == "CADASTROS_RELACIONADOS_DESCOBERTO_TARDIO"
        and extra.get("cadastro", {}).get("ordem") == 1
        and extra.get("multi_cadastro_ativo") is True
        for contexto, motivo, extra in evidencias
    )
    assert any(
        contexto == "DECLARACAO_FISCAL_TOMADOS"
        and motivo == "CADASTROS_RELACIONADOS_REAPARECEU"
        and extra.get("cadastro", {}).get("ordem") == 1
        for contexto, motivo, extra in evidencias
    )
    main.resetar_contexto_cadastro()


def test_recuperar_contexto_cadastro_relacionado_falha_sem_cadastro_ativo_para_dashboard(monkeypatch):
    evidencias = []

    main.resetar_multicadastro_tardio()
    monkeypatch.setattr(main, "tela_cadastros_relacionados_aberta", lambda: True)
    monkeypatch.setattr(main, "tela_declaracao_fiscal_eletronica_aberta", lambda: False)
    monkeypatch.setattr(main, "dashboard_aberto", lambda: False)
    monkeypatch.setattr(main, "MULTI_CADASTRO_ATIVO", False)
    monkeypatch.setattr(main, "contexto_cadastro_atual", lambda: {})
    monkeypatch.setattr(main, "append_txt", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        main,
        "salvar_evidencia_html",
        lambda contexto, motivo, **kwargs: evidencias.append((contexto, motivo)) or {},
    )

    with pytest.raises(RuntimeError, match="sem cadastro ativo"):
        main.recuperar_contexto_cadastro_relacionado(
            "dashboard",
            timeout=5,
            log_path="dummy.log",
            contexto_evidencia="DECLARACAO_FISCAL_PRESTADOS",
            ano_alvo=2026,
            mes_alvo=2,
            ref_alvo="02/2026",
        )

    assert ("DECLARACAO_FISCAL_PRESTADOS", "CADASTROS_RELACIONADOS_REAPARECEU") in evidencias
    assert ("DECLARACAO_FISCAL_PRESTADOS", "RECUPERACAO_CADASTRO_FALHOU") in evidencias


def test_executar_fluxo_fiscal_multicadastro_tardio_percorre_ordem_e_continua_apos_falha(monkeypatch):
    cadastros = [
        {"ordem": 1, "page": 1, "row_index": 1, "row_key": "1", "valor": "1", "ccm": "619824", "identificacao": "igual"},
        {"ordem": 2, "page": 1, "row_index": 2, "row_key": "2", "valor": "1", "ccm": "619824", "identificacao": "igual"},
        {"ordem": 3, "page": 1, "row_index": 3, "row_key": "3", "valor": "1", "ccm": "619824", "identificacao": "igual"},
    ]
    calls = []

    main.inicializar_multicadastro_tardio(cadastros, indice_atual=0)
    monkeypatch.setattr(main, "append_txt", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "clicar_inicio_para_dashboard", lambda timeout=0: calls.append(("inicio", timeout)))
    monkeypatch.setattr(main, "abrir_declaracao_fiscal", lambda timeout=0: calls.append(("declaracao", timeout)))
    monkeypatch.setattr(main, "garantir_declaracao_fiscal_contexto", lambda **kwargs: True)
    monkeypatch.setattr(main, "retornar_para_cadastros_relacionados", lambda timeout=0: calls.append(("voltar", timeout)))
    monkeypatch.setattr(
        main,
        "selecionar_cadastro_relacionado",
        lambda cadastro, timeout=0: calls.append(("selecionar", cadastro["ordem"])) or dict(cadastro),
    )
    monkeypatch.setattr(
        main,
        "executar_etapa_servicos_tomados",
        lambda ano, mes: calls.append(("tomados_xml", main.contexto_cadastro_atual()["ordem"])) or ("ERRO" if main.contexto_cadastro_atual()["ordem"] == 1 else "OK"),
    )
    monkeypatch.setattr(
        main,
        "executar_cadeado_tomados",
        lambda ano, mes: calls.append(("cadeado_tomados", main.contexto_cadastro_atual()["ordem"])) or "OK",
    )
    monkeypatch.setattr(
        main,
        "executar_cadeado_prestados",
        lambda ano, mes: calls.append(("cadeado_prestados", main.contexto_cadastro_atual()["ordem"])) or "OK",
    )
    monkeypatch.setattr(main, "pausa_final_livros_manual", lambda *args, **kwargs: calls.append(("pausa_final", "")))

    resultado = main.executar_fluxo_fiscal_multicadastro_tardio_se_necessario(
        2026,
        2,
        executar_tomados_xml=True,
        executar_cadeado_tomados_flag=True,
        executar_cadeado_prestados_flag=True,
        executar_pausa_final=False,
    )

    assert resultado is not None
    assert resultado["status"] == "REVISAO_MANUAL"
    assert [item for item in calls if item[0] == "selecionar"] == [("selecionar", 2), ("selecionar", 3)]
    assert [item for item in calls if item[0] == "tomados_xml"] == [("tomados_xml", 1), ("tomados_xml", 2), ("tomados_xml", 3)]
    assert [item for item in calls if item[0] == "cadeado_tomados"] == [("cadeado_tomados", 1), ("cadeado_tomados", 2), ("cadeado_tomados", 3)]
    assert [item for item in calls if item[0] == "cadeado_prestados"] == [("cadeado_prestados", 1), ("cadeado_prestados", 2), ("cadeado_prestados", 3)]
    assert main.estado_multicadastro_tardio()["falhos"] == [1]
    main.resetar_contexto_cadastro()


def test_abrir_servicos_nao_permitem_cadastros_relacionados_aberto(monkeypatch):
    monkeypatch.setattr(main, "tela_cadastros_relacionados_aberta", lambda: True)

    with pytest.raises(RuntimeError, match="Cadastros Relacionados"):
        main.abrir_servicos_tomados(timeout=1)

    with pytest.raises(RuntimeError, match="Cadastros Relacionados"):
        main.abrir_servicos_prestados(timeout=1)


def test_executar_cadeado_prestados_nao_avanca_quando_recuperacao_da_declaracao_falha(monkeypatch):
    chamadas = []

    monkeypatch.setattr(main, "PRESTADOS_SEM_MODULO", False)
    monkeypatch.setattr(main, "PAUSAR_ENTRE_MODULOS_CADEADO", False)
    monkeypatch.setattr(main, "append_txt", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "clicar_inicio_para_dashboard", lambda timeout=0: None)
    monkeypatch.setattr(main, "abrir_declaracao_fiscal", lambda timeout=0: None)
    monkeypatch.setattr(
        main,
        "garantir_declaracao_fiscal_contexto",
        lambda **kwargs: (_ for _ in ()).throw(RuntimeError("Cadastros Relacionados reapareceu")),
    )
    monkeypatch.setattr(main, "abrir_servicos_prestados", lambda timeout=0: chamadas.append("prestados"))

    resultado = main.executar_cadeado_prestados(2026, 2)

    assert resultado == "ERRO"
    assert chamadas == []


def test_executar_fluxo_empresa_atual_usa_fluxo_tardio_sem_repetir_prestados_xml(monkeypatch):
    calls = []

    monkeypatch.setattr(main, "PRESTADOS_SEM_MODULO", False)
    monkeypatch.setattr(main, "MODO_APENAS_CADEADO", False)
    monkeypatch.setattr(main, "PERFIL_EXECUCAO_ATIVO", False)
    monkeypatch.setattr(main, "APURAR_COMPLETO", False)
    monkeypatch.setattr(main, "ACRESCENTAR_CADEADO_E_PAUSA_FINAL", True)
    monkeypatch.setattr(main, "PAUSA_MANUAL_FINAL", False)
    monkeypatch.setattr(main, "precisa_lista_nota_fiscal_inicial", lambda: False)
    monkeypatch.setattr(main, "executar_etapa_servicos_prestados", lambda ano, mes: calls.append("prestados_xml") or "OK")
    monkeypatch.setattr(
        main,
        "executar_fluxo_fiscal_multicadastro_tardio_se_necessario",
        lambda *args, **kwargs: calls.append("tardio") or {"status": "SUCESSO", "motivo": "OK", "resultados": []},
    )
    monkeypatch.setattr(main, "executar_etapa_servicos_tomados", lambda ano, mes: calls.append("tomados_xml") or "OK")
    monkeypatch.setattr(main, "executar_fechamento_fiscal", lambda *args, **kwargs: calls.append("fechamento"))
    monkeypatch.setattr(main, "sleep", lambda *_: None)

    resultado = main.executar_fluxo_empresa_atual("03/2026", 2026, 2)

    assert resultado == "SUCESSO"
    assert calls == ["prestados_xml", "tardio"]
