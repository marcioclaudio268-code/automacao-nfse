from pathlib import Path

import main


def _configurar_contexto_saida_geral(monkeypatch, tmp_path):
    monkeypatch.setattr(main, "OUTPUT_BASE_DIR", str(tmp_path))
    monkeypatch.setattr(main, "DOWNLOAD_DIR", str(tmp_path / "downloads"))
    monkeypatch.setattr(main, "EMPRESA_PASTA", "ABS_REPRESENTACOES")
    monkeypatch.setattr(
        main,
        "contexto_cadastro_atual",
        lambda: {
            "ccm": "00833",
            "identificacao": "ABS REPRESENTACOES",
            "ordem": 1,
        },
    )


def test_salvar_pdf_guia_espelha_para_saida_geral(monkeypatch, tmp_path):
    _configurar_contexto_saida_geral(monkeypatch, tmp_path)

    origem = tmp_path / "tmp" / "guia.pdf"
    origem.parent.mkdir(parents=True, exist_ok=True)
    origem.write_text("guia", encoding="utf-8")

    destino = Path(main.salvar_pdf_guia(str(origem), "PRESTADOS", 2026, 2))
    espelho = tmp_path / "SAIDAS_GERAIS" / "ISS" / "02.2026" / "00833_ABS_REPRESENTACOES_ISS_02-2026.pdf"

    assert destino == tmp_path / "downloads" / "ABS_REPRESENTACOES" / "02.2026" / "PRESTADOS" / "CCM_00833__CAD_01__GUIA_ISS_PRESTADOS_02-2026.pdf"
    assert destino.exists()
    assert espelho.exists()
    assert destino.read_text(encoding="utf-8") == espelho.read_text(encoding="utf-8")
    assert not origem.exists()


def test_salvar_xml_tomados_espelha_para_saida_geral(monkeypatch, tmp_path):
    _configurar_contexto_saida_geral(monkeypatch, tmp_path)

    origem = tmp_path / "tmp" / "tomados.xml"
    origem.parent.mkdir(parents=True, exist_ok=True)
    origem.write_text("<xml />", encoding="utf-8")

    destino = Path(main.salvar_xml_tomados(str(origem), 2026, 2))
    espelho = tmp_path / "SAIDAS_GERAIS" / "XML_TOMADOS" / "02.2026" / "00833_ABS_REPRESENTACOES_SERVICOS_TOMADOS_02-2026.xml"

    assert destino == tmp_path / "downloads" / "ABS_REPRESENTACOES" / "02.2026" / "TOMADOS" / "CCM_00833__CAD_01__SERVICOS_TOMADOS_02-2026.xml"
    assert destino.exists()
    assert espelho.exists()
    assert destino.read_text(encoding="utf-8") == espelho.read_text(encoding="utf-8")
    assert not origem.exists()


def test_falha_no_espelhamento_nao_quebra_salvamento_principal(monkeypatch, tmp_path):
    _configurar_contexto_saida_geral(monkeypatch, tmp_path)

    origem = tmp_path / "tmp" / "guia.pdf"
    origem.parent.mkdir(parents=True, exist_ok=True)
    origem.write_text("guia", encoding="utf-8")

    monkeypatch.setattr(main.shutil, "copy2", lambda *args, **kwargs: (_ for _ in ()).throw(OSError("falha espelho")))

    destino = Path(main.salvar_pdf_guia(str(origem), "PRESTADOS", 2026, 2))
    espelho = tmp_path / "SAIDAS_GERAIS" / "ISS" / "02.2026" / "00833_ABS_REPRESENTACOES_ISS_02-2026.pdf"

    assert destino.exists()
    assert not espelho.exists()

