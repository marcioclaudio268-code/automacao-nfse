import main
import re
from types import SimpleNamespace


class _FakeTextElement:
    def __init__(self, text):
        self.text = text


class _FakeRefCell:
    def __init__(self):
        self.clicked = 0

    def is_selected(self):
        return False


class _FakeRow:
    def __init__(self):
        self.ref_cell = _FakeRefCell()

    def find_element(self, by, value):
        if ",7_grid" in value:
            return self.ref_cell
        raise LookupError(value)

    def find_elements(self, by, value):
        return []

    def get_attribute(self, name):
        return ""


class _FakeGridHeader:
    def __init__(self, text, columnorder):
        self.text = text
        self._columnorder = str(columnorder)

    def get_attribute(self, name):
        if name == "columnorder":
            return self._columnorder
        return ""


class _FakeGridRow:
    def __init__(self, cells):
        self.cells = dict(cells)

    def find_element(self, by, value):
        match = re.search(r",(-?\d+)_gridLista", value)
        if match:
            col = int(match.group(1))
            if col in self.cells:
                return _FakeTextElement(self.cells[col])
        raise LookupError(value)

    def find_elements(self, by, value):
        return []

    def get_attribute(self, name):
        return ""


class _FakeGridCheckbox:
    def __init__(self, row, value="linha-1"):
        self.row = row
        self.value = value

    def get_attribute(self, name):
        if name == "value":
            return self.value
        return ""

    def find_element(self, by, value):
        if "ancestor::tr" in value:
            return self.row
        raise LookupError(value)


class _FakeGridDriver:
    def __init__(self, headers):
        self.headers = list(headers)

    def execute_script(self, script, *args):
        return None

    def find_elements(self, by, value):
        if value == "#_gridListaTHeadLinhas th[columnorder]":
            return self.headers
        return []


class _FakeDriver:
    def __init__(self, state):
        self.state = state

    def execute_script(self, script, element):
        return None

    def find_element(self, by, value):
        if value == "_label30" and self.state["topo"]:
            return _FakeTextElement(self.state["topo"])
        raise LookupError(value)


class _DummyWait:
    def __init__(self, *args, **kwargs):
        pass

    def until(self, condition):
        return True


class _FakeFrameEl:
    def __init__(self, frame_id, name=None, displayed=True):
        self._id = frame_id
        self._name = name if name is not None else frame_id
        self._displayed = displayed

    def get_attribute(self, name):
        if name == "id":
            return self._id
        if name == "name":
            return self._name
        return ""

    def is_displayed(self):
        return self._displayed


class _FakeSwitchTo:
    def __init__(self, driver):
        self.driver = driver

    def default_content(self):
        self.driver.current_outer = None
        self.driver.current_inner = None

    def frame(self, frame):
        if self.driver.current_outer is None:
            self.driver.current_outer = frame
            self.driver.current_inner = None
            return
        self.driver.current_inner = frame


class _FakeContextDriver:
    def __init__(self):
        self.outer_frames = [
            _FakeFrameEl("_iFilho0"),
            _FakeFrameEl("_iFilho1"),
        ]
        self.inner_frames = {
            "_iFilho0": [_FakeFrameEl("inferior", "inferior")],
            "_iFilho1": [_FakeFrameEl("inferior", "inferior")],
        }
        self.current_outer = None
        self.current_inner = None
        self.switch_to = _FakeSwitchTo(self)

    def find_elements(self, by, value):
        if value == "iframe":
            return self.outer_frames
        if value == "frame" and self.current_outer is not None:
            return self.inner_frames.get(self.current_outer.get_attribute("id"), [])
        return []


class _EvalWait:
    def __init__(self, driver, timeout, *args, **kwargs):
        self.driver = driver

    def until(self, condition):
        result = condition(self.driver)
        return True if result is None else result


class _FakeClickable:
    def __init__(self, title=""):
        self._title = title

    def is_displayed(self):
        return True

    def is_enabled(self):
        return True

    def get_attribute(self, name):
        if name == "title":
            return self._title
        return ""


class _FakeTomadosRow:
    def __init__(self, icon):
        self.icon = icon

    def find_element(self, by, value):
        if "ancestor::tr" in value:
            raise LookupError(value)
        if "Visualizar NFSe" in value:
            return self.icon
        raise LookupError(value)


class _FakeTomadosCheckbox:
    def __init__(self, row):
        self.row = row

    def is_selected(self):
        return True

    def find_element(self, by, value):
        if "ancestor::tr" in value:
            return self.row
        raise LookupError(value)


class _FakeTomadosDriver:
    def __init__(self, checkbox, btn_pdf):
        self.checkbox = checkbox
        self.btn_pdf = btn_pdf
        self.window_handles = ["base"]
        self.switch_to = SimpleNamespace(default_content=lambda: None)

    def find_elements(self, by, value):
        if value == "gridListaCheck":
            return [self.checkbox]
        return []

    def find_element(self, by, value):
        if value == "btnPDF":
            return self.btn_pdf
        raise LookupError(value)

    def execute_script(self, script, *args):
        return None


class _FakeLivroSwitchTo:
    def default_content(self):
        return None

    def frame(self, frame):
        return None


class _FakeLivroDriver:
    def __init__(self):
        self.window_handles = ["base"]
        self.switch_to = _FakeLivroSwitchTo()

    def find_elements(self, by, value):
        if value == "iframe":
            return [object()]
        if value == "frame":
            return [object(), object()]
        return []

    def execute_script(self, script, *args):
        return None


class _LivroWait:
    def __init__(self, driver, timeout):
        self.driver = driver

    def until(self, condition):
        result = condition(self.driver)
        return True if result is None else result


class _FakeCheckboxDriver:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = 0

    def find_elements(self, by, value):
        if value != "gridListaCheck":
            return []
        idx = min(self.calls, len(self.responses) - 1)
        self.calls += 1
        return self.responses[idx]


class _FakePaginationSpan:
    def __init__(self, text="", onclick="", on_click=None):
        self.text = text
        self._onclick = onclick
        self._on_click = on_click

    def get_attribute(self, name):
        if name == "onclick":
            return self._onclick
        return ""

    def click(self):
        if self._on_click:
            self._on_click()


class _FakePaginationDriver:
    def __init__(self, current_page=1):
        self.current_page = current_page
        self.signature = f"pagina-{current_page}"

    def execute_script(self, script, *args):
        return ""

    def find_elements(self, by, value):
        if "contains(@class,'active')" in value:
            return [_FakePaginationSpan(text=str(self.current_page))]

        def go(page):
            def _advance():
                self.current_page = page
                self.signature = f"pagina-{page}"
            return _advance

        if "contains(@class,'next')" in value:
            return [
                _FakePaginationSpan(
                    text="»",
                    onclick=f"fastSubmit(this,'navegacao','mudarPagina,gridLista,{self.current_page + 1}','x',gridListaRefresh)",
                    on_click=go(self.current_page + 1),
                )
            ]

        if "contains(@onclick,'mudarPagina,gridLista,')" in value:
            return [
                _FakePaginationSpan(
                    text=str(self.current_page + 1),
                    onclick=f"fastSubmit(this,'navegacao','mudarPagina,gridLista,{self.current_page + 1}','x',gridListaRefresh)",
                    on_click=go(self.current_page + 1),
                ),
                _FakePaginationSpan(
                    text=str(self.current_page + 2),
                    onclick=f"fastSubmit(this,'navegacao','mudarPagina,gridLista,{self.current_page + 2}','x',gridListaRefresh)",
                    on_click=go(self.current_page + 2),
                ),
            ]

        return []


def test_garantir_contexto_referencia_grid_clica_ate_topo_bater(monkeypatch):
    state = {"topo": "03/2026"}
    row = _FakeRow()

    monkeypatch.setattr(main, "driver", _FakeDriver(state))
    monkeypatch.setattr(main, "sleep", lambda *_: None)
    monkeypatch.setattr(main, "encontrar_tr_por_referencia_grid", lambda ref, col: row)

    def fake_click(target):
        if target in (row.ref_cell, row):
            row.ref_cell.clicked += 1
            state["topo"] = "02/2026"

    monkeypatch.setattr(main, "click_robusto", fake_click)

    retorno = main.garantir_contexto_referencia_grid("02/2026", "7", timeout=1)

    assert retorno is row
    assert row.ref_cell.clicked >= 1
    assert state["topo"] == "02/2026"


def test_executar_cadeado_tomados_ativa_contexto_antes_do_livro(monkeypatch):
    calls = []

    monkeypatch.setattr(main, "driver", object())
    monkeypatch.setattr(main, "PAUSAR_ENTRE_MODULOS_CADEADO", False)
    monkeypatch.setattr(main, "pasta_empresa", lambda: "C:/tmp")
    monkeypatch.setattr(main, "append_txt", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "clicar_inicio_para_dashboard", lambda timeout=0: None)
    monkeypatch.setattr(main, "abrir_declaracao_fiscal", lambda timeout=0: None)
    monkeypatch.setattr(main, "garantir_declaracao_fiscal_contexto", lambda **kwargs: True)
    monkeypatch.setattr(main, "declaracao_fiscal_tem_servicos_prestados", lambda: True)
    monkeypatch.setattr(main, "abrir_servicos_tomados", lambda timeout=0: None)
    monkeypatch.setattr(main, "fechar_alerta_mensagens_prefeitura_se_aberto", lambda **kwargs: False)
    monkeypatch.setattr(main, "WebDriverWait", _DummyWait)
    monkeypatch.setattr(main, "tomados_grid_sem_registros", lambda timeout=0: False)
    monkeypatch.setattr(main, "esperar_grid_declaracao_fiscal", lambda timeout=0: "7")
    monkeypatch.setattr(main, "encontrar_tr_por_referencia_grid", lambda ref, col: object())
    monkeypatch.setattr(
        main,
        "garantir_contexto_referencia_grid",
        lambda ref, col, **kwargs: calls.append("context") or object(),
    )
    monkeypatch.setattr(main, "clicar_cadeado_na_linha", lambda ref, col: calls.append("lock") or False)
    monkeypatch.setattr(main, "abrir_modal_livro_e_visualizar", lambda *args: calls.append("livro") or "OK")
    monkeypatch.setattr(main, "baixar_guia_iss_se_existir", lambda *args: calls.append("guia") or "NAO")

    resultado = main.executar_cadeado_tomados(2026, 2)

    assert resultado == "OK"
    assert calls[:3] == ["context", "lock", "livro"]


def test_entrar_contexto_livro_varre_todos_ifilhos(monkeypatch):
    driver = _FakeContextDriver()

    monkeypatch.setattr(main, "driver", driver)
    monkeypatch.setattr(main, "sleep", lambda *_: None)

    def fake_esperar_controles(timeout=0):
        outer = driver.current_outer.get_attribute("id") if driver.current_outer else ""
        inner = driver.current_inner.get_attribute("id") if driver.current_inner else ""
        if outer == "_iFilho1" and inner == "inferior":
            return {"exercicio": object(), "mes": object(), "visualizar": object()}
        raise Exception("sem controles")

    monkeypatch.setattr(main, "esperar_controles_livro", fake_esperar_controles)

    contexto = main.entrar_contexto_livro(timeout=1)

    assert contexto["iframe"]["id"] == "_iFilho1"
    assert contexto["inner_frame"]["id"] == "inferior"


def test_obter_checkboxes_lista_retorna_vazio_quando_lista_sem_checkbox(monkeypatch):
    monkeypatch.setattr(main, "driver", _FakeCheckboxDriver([[]]))
    monkeypatch.setattr(main, "esperar_lista_ou_sem_checkbox", lambda timeout=0: "sem_checkbox_com_data")

    assert main.obter_checkboxes_lista(timeout=5) == []


def test_obter_checkboxes_lista_retorna_itens_apos_fallback(monkeypatch):
    checkboxes = [object(), object()]
    driver = _FakeCheckboxDriver([[], checkboxes])

    monkeypatch.setattr(main, "driver", driver)
    monkeypatch.setattr(main, "esperar_lista_ou_sem_checkbox", lambda timeout=0: "checkboxes")

    assert main.obter_checkboxes_lista(timeout=5) == checkboxes


def test_pagina_atual_lida_com_botao_ativo_visual(monkeypatch):
    driver = _FakePaginationDriver(current_page=3)

    monkeypatch.setattr(main, "driver", driver)

    assert main.pagina_atual() == 3


def test_ir_para_proxima_pagina_usa_onclick_real_da_paginacao(monkeypatch):
    driver = _FakePaginationDriver(current_page=1)

    monkeypatch.setattr(main, "driver", driver)
    monkeypatch.setattr(main, "WebDriverWait", _EvalWait)
    monkeypatch.setattr(main, "click_robusto", lambda el: el.click())
    monkeypatch.setattr(main, "assinatura_lista", lambda: driver.signature)
    monkeypatch.setattr(main, "ULTIMO_CHECKBOX_MARCADO", object())

    assert main.ir_para_proxima_pagina() is True
    assert driver.current_page == 2
    assert main.ULTIMO_CHECKBOX_MARCADO is None


def test_mapa_colunas_grid_lista_e_extrair_info_linha_com_cabecalhos_acentuados(monkeypatch):
    headers = [
        _FakeGridHeader("NF", 15),
        _FakeGridHeader("Data Emissão", 19),
        _FakeGridHeader("Data Emissão RPS", 20),
        _FakeGridHeader("Nº RPS", 21),
        _FakeGridHeader("Situação", 22),
        _FakeGridHeader("Chave de Validação/Acesso", 23),
    ]
    row = _FakeGridRow(
        {
            14: "23207",
            18: "11/03/2026",
            20: "49999",
            21: "Normal",
            22: "ABC123",
        }
    )
    checkbox = _FakeGridCheckbox(row)

    monkeypatch.setattr(main, "driver", _FakeGridDriver(headers))
    monkeypatch.setattr(main, "GRID_HEADER_CACHE", {"ts": 0.0, "map": {}})

    mapa = main.mapa_colunas_grid_lista(force=True)
    assert mapa["nf"] == 14
    assert mapa["data_emissao"] == 18
    assert mapa["rps"] == 20
    assert mapa["situacao"] == 21
    assert mapa["chave"] == 22

    info = main.extrair_info_linha(checkbox)
    assert info["nf"] == "23207"
    assert info["data_emissao"] == "11/03/2026"
    assert info["rps"] == "49999"
    assert info["situacao"] == "Normal"
    assert info["chave"] == "ABC123"


def test_baixar_pdf_tomados_por_indice_prioriza_popup(monkeypatch):
    icon = _FakeClickable("Visualizar NFSe")
    btn_pdf = _FakeClickable("PDF")
    row = _FakeTomadosRow(icon)
    checkbox = _FakeTomadosCheckbox(row)
    driver = _FakeTomadosDriver(checkbox, btn_pdf)
    calls = []

    monkeypatch.setattr(main, "driver", driver)
    monkeypatch.setattr(main, "WebDriverWait", _EvalWait)
    monkeypatch.setattr(main, "marcar_somente_checkbox", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "extrair_info_linha_tomados", lambda *_: {"nf": "885", "data_emissao": "24/02/2026"})
    monkeypatch.setattr(main, "click_robusto", lambda *_: None)
    monkeypatch.setattr(
        main,
        "baixar_pdf_tomados_via_popup",
        lambda *args, **kwargs: (calls.append("popup") or True) and ("C:/tmp/popup.pdf", {"url": "https://tributario.bauru.sp.gov.br/resultados/teste.pdf", "title": "popup"}),
    )
    monkeypatch.setattr(
        main,
        "baixar_arquivo_via_submit_form",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("submit_form nao deveria ser chamado")),
    )
    monkeypatch.setattr(
        main,
        "baixar_pdf_tomados_via_post_controle",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("post_controle nao deveria ser chamado")),
    )
    monkeypatch.setattr(main, "salvar_pdf_tomados_nota", lambda *args, **kwargs: "C:/dest/NFT_885_24-02-2026.pdf")
    monkeypatch.setattr(main, "limpar_status_tomados_pdf", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "salvar_log_tomados", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "fechar_janelas_extras", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "fechar_modal_tomados_visualizacao", lambda *args, **kwargs: None)

    destino = main.baixar_pdf_nota_tomados_por_indice(0, 2026, 2, "02/2026")

    assert destino == "C:/dest/NFT_885_24-02-2026.pdf"
    assert calls == ["popup"]


def test_baixar_pdf_tomados_via_popup_recupera_fluxo_de_logs(monkeypatch):
    fluxo = {
        "controle_302": {
            "url": "https://tributario.bauru.sp.gov.br/servlet/controle",
            "location": "https://tributario.bauru.sp.gov.br/resultados/teste.pdf",
        },
        "pdf_response": {},
    }

    monkeypatch.setattr(main, "_limpar_logs_performance", lambda: None)
    monkeypatch.setattr(main, "click_robusto", lambda *_: None)
    monkeypatch.setattr(
        main,
        "aguardar_popup_pdf_tomados",
        lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("popup travada em about:blank")),
    )
    monkeypatch.setattr(main, "capturar_fluxo_pdf_via_logs", lambda timeout=0: fluxo)
    monkeypatch.setattr(
        main,
        "materializar_pdf_de_fluxo_rede",
        lambda *args, **kwargs: ("C:/tmp/NFT_885.pdf", "logs_rede_url"),
    )

    pdf_tmp, info = main.baixar_pdf_tomados_via_popup(
        object(),
        ["base"],
        timeout=120,
        nome_fallback="NFT_885.pdf",
    )

    assert pdf_tmp == "C:/tmp/NFT_885.pdf"
    assert info["source"] == "logs_rede_url"
    assert info["url"] == "https://tributario.bauru.sp.gov.br/resultados/teste.pdf"
    assert "about:blank" in info["popup_error"]


def test_prefs_nao_forcam_download_externo_de_pdf():
    assert main.prefs["plugins.always_open_pdf_externally"] is False


def test_abrir_modal_livro_prioriza_popup_real(monkeypatch):
    driver = _FakeLivroDriver()
    btn_livro = _FakeClickable("Livro")
    btn_visualizar = _FakeClickable("Visualizar")
    modal = {"visualizar": btn_visualizar}
    logs = []

    monkeypatch.setattr(main, "driver", driver)
    monkeypatch.setattr(main, "pasta_empresa", lambda: "C:/tmp")
    monkeypatch.setattr(main, "append_txt", lambda _path, linha: logs.append(linha))
    monkeypatch.setattr(main, "WebDriverWait", _LivroWait)
    monkeypatch.setattr(main.EC, "element_to_be_clickable", lambda locator: (lambda d: btn_livro))
    monkeypatch.setattr(main, "click_robusto", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        main,
        "entrar_contexto_livro",
        lambda timeout=0: {
            "iframe": {"idx": 0, "id": "_iFilho1", "name": "_iFilho1"},
            "inner_frame": {"idx": 1, "id": "inferior", "name": "inferior"},
        },
    )
    monkeypatch.setattr(main, "esperar_controles_livro", lambda timeout=0: modal)
    monkeypatch.setattr(main, "selecionar_competencia_modal_livro", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "limpar_tmp_livros_pdf", lambda: None)
    monkeypatch.setattr(
        main,
        "baixar_arquivo_via_submit_form",
        lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("submit falhou")),
    )
    monkeypatch.setattr(
        main,
        "baixar_pdf_via_popup_ou_logs",
        lambda *args, **kwargs: (
            "C:/tmp/livro.pdf",
            {
                "url": "https://tributario.bauru.sp.gov.br/resultados/livro.pdf",
                "source": "popup_url",
            },
        ),
    )
    monkeypatch.setattr(
        main,
        "aguardar_pdf_novo",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("nao deveria aguardar arquivo _tmp")),
    )
    monkeypatch.setattr(
        main,
        "capturar_fluxo_pdf_via_logs",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("nao deveria capturar logs depois do popup")),
    )
    monkeypatch.setattr(
        main,
        "coletar_urls_pdf_pos_visualizar",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("nao deveria varrer URLs depois do popup")),
    )
    monkeypatch.setattr(main, "salvar_pdf_livro", lambda *args, **kwargs: "C:/dest/LIVRO_SERVICOS_TOMADOS_02-2026.pdf")
    monkeypatch.setattr(main, "fechar_janelas_extras", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "fechar_modal_livro_se_aberto", lambda timeout=0: False)

    resultado = main.abrir_modal_livro_e_visualizar("TOMADOS", "tomados", 2026, 2, "02/2026")

    assert resultado == "OK"
    assert any("PDF_POPUP_URL=OK" in linha for linha in logs)
    assert any("PDF_FETCH=OK | via popup_url" in linha for linha in logs)
