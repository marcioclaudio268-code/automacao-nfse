from pathlib import Path

import main


def test_organizar_xml_por_pasta_salva_em_prestados_por_competencia(monkeypatch, tmp_path):
    origem = tmp_path / "nota.xml"
    origem.write_text("<xml />", encoding="utf-8")

    monkeypatch.setattr(main, "DOWNLOAD_DIR", str(tmp_path))
    monkeypatch.setattr(main, "EMPRESA_PASTA", "833_EMPRESA")
    monkeypatch.setattr(main, "parse_data_emissao_site", lambda *_: ("2026", "02", "14", "2026-02-14", "02.2026"))

    destino = Path(main.organizar_xml_por_pasta(str(origem), {"nf": "413", "data_emissao": "14/02/2026"}))

    assert destino == tmp_path / "833_EMPRESA" / "02.2026" / "PRESTADOS" / "NFS_413_14-02-2026.xml"
    assert destino.exists()


def test_salvar_xml_tomados_salva_em_tomados_por_competencia(monkeypatch, tmp_path):
    origem = tmp_path / "tomados.xml"
    origem.write_text("<xml />", encoding="utf-8")

    monkeypatch.setattr(main, "DOWNLOAD_DIR", str(tmp_path))
    monkeypatch.setattr(main, "EMPRESA_PASTA", "833_EMPRESA")

    destino = Path(main.salvar_xml_tomados(str(origem), 2026, 2))

    assert destino == tmp_path / "833_EMPRESA" / "02.2026" / "TOMADOS" / "SERVICOS_TOMADOS_02-2026.xml"
    assert destino.exists()


def test_novo_prefixo_evidencia_separa_por_servico_e_geral(monkeypatch, tmp_path):
    monkeypatch.setattr(main, "DOWNLOAD_DIR", str(tmp_path))
    monkeypatch.setattr(main, "EMPRESA_PASTA", "833_EMPRESA")
    monkeypatch.setattr(main, "APURACAO_REFERENCIA", "03/2026")
    monkeypatch.setattr(main, "APURAR_COMPLETO", False)

    prefixo_tomados = Path(
        main._novo_prefixo_evidencia(
            "SERVICOS_TOMADOS",
            "PDF_NFSE_FALHA",
            ano_alvo=2026,
            mes_alvo=2,
            ref_alvo="02/2026",
        )
    )
    prefixo_geral = Path(main._novo_prefixo_evidencia("LOGIN_CAPTCHA", "MANUAL_LOGIN_REQUIRED"))

    assert prefixo_tomados.parent == tmp_path / "833_EMPRESA" / "02.2026" / "TOMADOS" / "_evidencias"
    assert prefixo_geral.parent == tmp_path / "833_EMPRESA" / "02.2026" / "_GERAL" / "_evidencias"


def test_helpers_de_geral_resolvem_logs_debug_e_status_por_competencia(monkeypatch, tmp_path):
    monkeypatch.setattr(main, "DOWNLOAD_DIR", str(tmp_path))
    monkeypatch.setattr(main, "EMPRESA_PASTA", "833_EMPRESA")
    monkeypatch.setattr(main, "APURACAO_REFERENCIA", "03/2026")
    monkeypatch.setattr(main, "APURAR_COMPLETO", False)

    log_nfse = Path(main.log_path_empresa({"data_emissao": "14/02/2026"}))
    log_tomados = Path(main.tomados_log_path_empresa("02/2026"))
    log_manual = Path(main.caminho_log_manual("02/2026"))
    debug_dir = Path(main.pasta_debug())
    status_tomados = Path(main.caminho_status_tomados_pdf(referencia="02/2026"))

    assert log_nfse == tmp_path / "833_EMPRESA" / "02.2026" / "_GERAL" / "log_downloads_nfse.txt"
    assert log_tomados == tmp_path / "833_EMPRESA" / "02.2026" / "_GERAL" / "log_tomados.txt"
    assert log_manual == tmp_path / "833_EMPRESA" / "02.2026" / "_GERAL" / "log_fechamento_manual.txt"
    assert debug_dir == tmp_path / "833_EMPRESA" / "02.2026" / "_GERAL" / "_debug"
    assert status_tomados == tmp_path / "833_EMPRESA" / "02.2026" / "_GERAL" / "_tomados_pdf_status.json"


def test_log_nfse_normal_fica_no_mes_alvo_mesmo_quando_nota_for_de_outro_mes(monkeypatch, tmp_path):
    monkeypatch.setattr(main, "DOWNLOAD_DIR", str(tmp_path))
    monkeypatch.setattr(main, "EMPRESA_PASTA", "833_EMPRESA")
    monkeypatch.setattr(main, "APURACAO_REFERENCIA", "03/2026")
    monkeypatch.setattr(main, "APURAR_COMPLETO", False)

    log_nfse = Path(main.log_path_empresa({"data_emissao": "09/03/2026"}))

    assert log_nfse == tmp_path / "833_EMPRESA" / "02.2026" / "_GERAL" / "log_downloads_nfse.txt"


def test_helpers_gerais_de_apuracao_completa_usam_mes_mais_recente(monkeypatch, tmp_path):
    monkeypatch.setattr(main, "DOWNLOAD_DIR", str(tmp_path))
    monkeypatch.setattr(main, "EMPRESA_PASTA", "833_EMPRESA")
    monkeypatch.setattr(main, "APURACAO_REFERENCIA", "01/2026")
    monkeypatch.setattr(main, "APURAR_COMPLETO", True)

    prefixo_geral = Path(main._novo_prefixo_evidencia("LOGIN_CAPTCHA", "MANUAL_LOGIN_REQUIRED"))
    log_apuracao = Path(main.caminho_log_apuracao_completa())
    log_nfse = Path(main.log_path_empresa({"data_emissao": "14/03/2025"}))

    assert prefixo_geral.parent == tmp_path / "833_EMPRESA" / "12.2025" / "_GERAL" / "_evidencias"
    assert log_apuracao == tmp_path / "833_EMPRESA" / "12.2025" / "_GERAL" / "log_apuracao_completa.txt"
    assert log_nfse == tmp_path / "833_EMPRESA" / "03.2025" / "_GERAL" / "log_downloads_nfse.txt"


def test_registrar_apuracao_completa_salva_no_mes_da_referencia(monkeypatch, tmp_path):
    monkeypatch.setattr(main, "DOWNLOAD_DIR", str(tmp_path))
    monkeypatch.setattr(main, "EMPRESA_PASTA", "833_EMPRESA")
    monkeypatch.setattr(main, "APURACAO_REFERENCIA", "01/2026")
    monkeypatch.setattr(main, "APURAR_COMPLETO", True)

    main.registrar_apuracao_completa("03/2025", "TOMADOS", "OK", "processado")

    path = tmp_path / "833_EMPRESA" / "03.2025" / "_GERAL" / "log_apuracao_completa.txt"
    assert path.exists()
    assert "REFERENCIA=03/2025" in path.read_text(encoding="utf-8")


def test_multi_cadastro_prefixa_artefatos_no_mes_sem_subpastas(monkeypatch, tmp_path):
    origem_nfse = tmp_path / "nota.xml"
    origem_nfse.write_text("<xml />", encoding="utf-8")
    origem_tomados = tmp_path / "tomados.xml"
    origem_tomados.write_text("<xml />", encoding="utf-8")

    monkeypatch.setattr(main, "DOWNLOAD_DIR", str(tmp_path))
    monkeypatch.setattr(main, "EMPRESA_PASTA", "833_EMPRESA")
    monkeypatch.setattr(main, "APURACAO_REFERENCIA", "03/2026")
    monkeypatch.setattr(main, "APURAR_COMPLETO", False)
    monkeypatch.setattr(main, "parse_data_emissao_site", lambda *_: ("2026", "02", "14", "2026-02-14", "02.2026"))

    main.ativar_contexto_cadastro(
        {
            "ordem": 2,
            "ccm": "619824",
            "identificacao": "Nome: EMPRESA TESTE CNPJ: 12.345.678/0001-90 CCM: 619824",
        }
    )
    try:
        destino_nfse = Path(main.organizar_xml_por_pasta(str(origem_nfse), {"nf": "413", "data_emissao": "14/02/2026"}))
        destino_tomados = Path(main.salvar_xml_tomados(str(origem_tomados), 2026, 2))
        status_path = Path(main.caminho_status_tomados_pdf(referencia="02/2026"))
    finally:
        main.resetar_contexto_cadastro()

    assert destino_nfse.parent == tmp_path / "833_EMPRESA" / "02.2026" / "PRESTADOS"
    assert destino_nfse.name.startswith("CCM_619824__CAD_02__")
    assert destino_tomados.parent == tmp_path / "833_EMPRESA" / "02.2026" / "TOMADOS"
    assert destino_tomados.name.startswith("CCM_619824__CAD_02__")
    assert status_path == tmp_path / "833_EMPRESA" / "02.2026" / "_GERAL" / "CCM_619824__CAD_02___tomados_pdf_status.json"
