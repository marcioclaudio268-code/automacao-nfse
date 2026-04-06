import main


def test_baixar_xml_prestados_via_menu_prioriza_submit_direto(monkeypatch):
    action = object()
    calls = {}

    monkeypatch.setattr(main, "_obter_acao_exportar_xml_nacional_prestados", lambda: action)
    monkeypatch.setattr(
        main,
        "baixar_arquivo_via_submit_form",
        lambda button, timeout=0, nome_fallback="": calls.update(
            {"button": button, "timeout": timeout, "nome_fallback": nome_fallback}
        ) or "C:/tmp/exportado_menu.xml",
    )
    monkeypatch.setattr(main, "click_robusto", lambda *_: (_ for _ in ()).throw(AssertionError("nao deveria clicar no fallback")))

    path, origem = main._baixar_xml_prestados_via_menu_exportar({"nf": "413"}, 0)

    assert path == "C:/tmp/exportado_menu.xml"
    assert origem == "menu_submit"
    assert calls["button"] is action
    assert calls["nome_fallback"] == "413.xml"


def test_baixar_xml_prestados_via_menu_faz_fallback_para_clique(monkeypatch):
    menu = object()
    action = object()
    clicks = []
    waits = {}

    monkeypatch.setattr(main, "_obter_menu_mais_acoes_prestados", lambda: menu)
    monkeypatch.setattr(main, "_obter_acao_exportar_xml_nacional_prestados", lambda: action)
    monkeypatch.setattr(main, "baixar_arquivo_via_submit_form", lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("submit falhou")))
    monkeypatch.setattr(main, "click_robusto", lambda el: clicks.append(el))
    monkeypatch.setattr(main, "fechar_aviso_se_existir", lambda timeout=0: False)
    monkeypatch.setattr(main.os, "listdir", lambda *_: [])

    def fake_aguardar_xml_novo(*args, **kwargs):
        waits["timeout"] = kwargs.get("timeout")
        return "C:/tmp/exportado_click.xml"

    monkeypatch.setattr(main, "aguardar_xml_novo", fake_aguardar_xml_novo)

    path, origem = main._baixar_xml_prestados_via_menu_exportar({"nf": "413"}, 0)

    assert path == "C:/tmp/exportado_click.xml"
    assert origem == "menu_click"
    assert clicks == [menu, action]
    assert waits["timeout"] == 60
