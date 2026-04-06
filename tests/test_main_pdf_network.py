import main


def test_materializar_pdf_de_fluxo_rede_prioriza_body_pdf(monkeypatch):
    calls = {}

    def fake_save(payload, nome):
        calls["payload"] = payload
        calls["nome"] = nome
        return "C:/tmp/arquivo_portal.pdf"

    monkeypatch.setattr(main, "_salvar_bytes_tmp", fake_save)

    fluxo = {
        "pdf_response": {
            "body": b"%PDF-1.4\n01234567890123456789",
            "contentDisposition": 'attachment; filename="arquivo_portal.pdf"',
            "url": "https://portal.exemplo/resultados/arquivo_portal.pdf",
        }
    }

    path, origem = main.materializar_pdf_de_fluxo_rede(fluxo, nome_fallback="fallback.pdf", timeout=30)

    assert path == "C:/tmp/arquivo_portal.pdf"
    assert origem == "logs_rede_body"
    assert calls["payload"].startswith(b"%PDF")
    assert calls["nome"] == "arquivo_portal.pdf"


def test_materializar_pdf_de_fluxo_rede_baixa_via_location(monkeypatch):
    calls = {}

    def fake_download(url, timeout=0, referer=""):
        calls["url"] = url
        calls["timeout"] = timeout
        calls["referer"] = referer
        return "C:/tmp/location.pdf"

    monkeypatch.setattr(main, "baixar_pdf_por_url_com_cookies", fake_download)
    monkeypatch.setattr(main, "_base_url_portal", lambda: "https://portal.exemplo/")

    fluxo = {
        "controle_302": {
            "url": "https://portal.exemplo/servlet/controle",
            "location": "http://portal.exemplo/resultados/location.pdf",
        }
    }

    path, origem = main.materializar_pdf_de_fluxo_rede(fluxo, nome_fallback="fallback.pdf", timeout=33)

    assert path == "C:/tmp/location.pdf"
    assert origem == "logs_rede_url"
    assert calls["url"] == "https://portal.exemplo/resultados/location.pdf"
    assert calls["timeout"] == 33
    assert calls["referer"] == "https://portal.exemplo/servlet/controle"


def test_materializar_pdf_de_fluxo_rede_extrai_url_do_html(monkeypatch):
    calls = {}

    def fake_download(url, timeout=0, referer=""):
        calls["url"] = url
        calls["timeout"] = timeout
        calls["referer"] = referer
        return "C:/tmp/html.pdf"

    monkeypatch.setattr(main, "baixar_pdf_por_url_com_cookies", fake_download)
    monkeypatch.setattr(main, "_base_url_portal", lambda: "https://portal.exemplo/")

    fluxo = {
        "pdf_response": {
            "body": b'<html><body><a href="/resultados/html.pdf">PDF</a></body></html>',
            "url": "https://portal.exemplo/servlet/controle",
        }
    }

    path, origem = main.materializar_pdf_de_fluxo_rede(fluxo, nome_fallback="fallback.pdf", timeout=18)

    assert path == "C:/tmp/html.pdf"
    assert origem == "logs_rede_url"
    assert calls["url"] == "https://portal.exemplo/resultados/html.pdf"
    assert calls["timeout"] == 18
    assert calls["referer"] == "https://portal.exemplo/"


def test_baixar_pdf_via_popup_ou_logs_usa_timeouts_curtos_de_fallback(monkeypatch):
    capturas = {}

    monkeypatch.setattr(main, "TEMP_DOWNLOAD_DIR", "C:/tmp")
    monkeypatch.setattr(main, "ENABLE_CHROME_PERFORMANCE_LOGS", True)
    monkeypatch.setattr(main, "PDF_POPUP_WAIT_TIMEOUT", 7)
    monkeypatch.setattr(main, "PDF_FILE_FALLBACK_TIMEOUT", 5)
    monkeypatch.setattr(main, "PDF_NETWORK_CAPTURE_TIMEOUT", 6)
    monkeypatch.setattr(main.os, "listdir", lambda *_: [])
    monkeypatch.setattr(main, "_limpar_logs_performance", lambda: None)
    monkeypatch.setattr(main, "click_robusto", lambda *_: None)
    monkeypatch.setattr(
        main,
        "aguardar_popup_pdf_tomados",
        lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError(f"popup timeout={kwargs.get('timeout')}")),
    )

    def fake_aguardar_pdf_novo(*args, **kwargs):
        capturas["file_timeout"] = kwargs.get("timeout")
        raise RuntimeError("arquivo ausente")

    monkeypatch.setattr(main, "aguardar_pdf_novo", fake_aguardar_pdf_novo)

    def fake_logs(timeout=0):
        capturas["network_timeout"] = timeout
        return {
            "controle_302": {
                "url": "https://tributario.bauru.sp.gov.br/servlet/controle",
                "location": "https://tributario.bauru.sp.gov.br/resultados/teste.pdf",
            },
            "pdf_response": {},
        }

    monkeypatch.setattr(main, "capturar_fluxo_pdf_via_logs", fake_logs)
    monkeypatch.setattr(
        main,
        "materializar_pdf_de_fluxo_rede",
        lambda *args, **kwargs: ("C:/tmp/fallback.pdf", "logs_rede_url"),
    )

    pdf_tmp, info = main.baixar_pdf_via_popup_ou_logs(
        object(),
        ["base"],
        timeout=120,
        nome_fallback="fallback.pdf",
    )

    assert pdf_tmp == "C:/tmp/fallback.pdf"
    assert info["source"] == "logs_rede_url"
    assert "timeout=7" in info["popup_error"]
    assert capturas["file_timeout"] == 5
    assert capturas["network_timeout"] == 6
