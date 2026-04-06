from application.artifact_locator_service import ArtifactLocatorService
from core.config_runtime import competencia_alvo_dir_name
from core.paths import build_runtime_paths


def test_competencia_alvo_dir_name_uses_previous_month():
    assert competencia_alvo_dir_name("02/2026") == "01.2026"
    assert competencia_alvo_dir_name("01/2026") == "12.2025"


def test_artifact_locator_prefers_output_base_dir_and_competencia_folder(tmp_path):
    project_dir = tmp_path / "project"
    output_dir = tmp_path / "output"
    runtime_paths = build_runtime_paths(project_dir, output_dir)
    service = ArtifactLocatorService()
    empresa = {
        "codigo": "833",
        "razao_social": "A. B. S. REPRESENTACOES",
        "cnpj": "41.570.055.0001-14",
    }

    pasta_empresa = runtime_paths.downloads_dir / "833_B_REPRESENTACOES"
    competencia_dir = pasta_empresa / "01.2026"
    geral_dir = competencia_dir / "_GERAL"
    evidencias_dir = geral_dir / "_evidencias"
    evidencias_dir.mkdir(parents=True)
    (geral_dir / "log_downloads_nfse.txt").write_text("nfse", encoding="utf-8")
    (geral_dir / "log_tomados.txt").write_text("tomados", encoding="utf-8")
    (geral_dir / "log_fechamento_manual.txt").write_text("manual", encoding="utf-8")

    artifacts = service.get_company_artifacts(runtime_paths, empresa, competencia_dir_name="01.2026")

    assert artifacts.pasta_empresa == pasta_empresa
    assert artifacts.competencia_dir == competencia_dir
    assert artifacts.geral_dir == geral_dir
    assert artifacts.evidencias_dir == evidencias_dir
    assert artifacts.log_downloads == geral_dir / "log_downloads_nfse.txt"
    assert artifacts.log_tomados == geral_dir / "log_tomados.txt"
    assert artifacts.log_manual == geral_dir / "log_fechamento_manual.txt"


def test_artifact_locator_aceita_evidencias_de_servico_na_estrutura_nova(tmp_path):
    project_dir = tmp_path / "project"
    output_dir = tmp_path / "output"
    runtime_paths = build_runtime_paths(project_dir, output_dir)
    service = ArtifactLocatorService()
    empresa = {
        "codigo": "833",
        "razao_social": "A. B. S. REPRESENTACOES",
        "cnpj": "41.570.055.0001-14",
    }

    pasta_empresa = runtime_paths.downloads_dir / "833_B_REPRESENTACOES"
    competencia_dir = pasta_empresa / "01.2026"
    evidencias_dir = competencia_dir / "PRESTADOS" / "_evidencias"
    evidencias_dir.mkdir(parents=True)

    artifacts = service.get_company_artifacts(runtime_paths, empresa, competencia_dir_name="01.2026")

    assert artifacts.competencia_dir == competencia_dir
    assert artifacts.geral_dir == competencia_dir / "_GERAL"
    assert artifacts.evidencias_dir == evidencias_dir


def test_artifact_locator_falls_back_to_legacy_downloads_for_historical_runs(tmp_path):
    project_dir = tmp_path / "project"
    output_dir = tmp_path / "output"
    runtime_paths = build_runtime_paths(project_dir, output_dir)
    service = ArtifactLocatorService()
    empresa = {
        "codigo": "833",
        "razao_social": "A. B. S. REPRESENTACOES",
        "cnpj": "41.570.055.0001-14",
    }

    legacy_company_dir = runtime_paths.legacy_downloads_dir / "833_B_REPRESENTACOES"
    (legacy_company_dir / "12.2025" / "_evidencias").mkdir(parents=True)

    artifacts = service.get_company_artifacts(runtime_paths, empresa, competencia_dir_name="01.2026")

    assert artifacts.pasta_empresa == legacy_company_dir
    assert artifacts.competencia_dir == legacy_company_dir / "12.2025"
    assert artifacts.geral_dir == legacy_company_dir / "12.2025" / "_GERAL"
    assert artifacts.evidencias_dir == legacy_company_dir / "12.2025" / "_evidencias"


def test_artifact_locator_faz_fallback_para_logs_e_debug_na_raiz_legada(tmp_path):
    project_dir = tmp_path / "project"
    output_dir = tmp_path / "output"
    runtime_paths = build_runtime_paths(project_dir, output_dir)
    service = ArtifactLocatorService()
    empresa = {
        "codigo": "833",
        "razao_social": "A. B. S. REPRESENTACOES",
        "cnpj": "41.570.055.0001-14",
    }

    legacy_company_dir = runtime_paths.legacy_downloads_dir / "833_B_REPRESENTACOES"
    (legacy_company_dir / "12.2025" / "_evidencias").mkdir(parents=True)
    (legacy_company_dir / "log_downloads_nfse.txt").write_text("nfse", encoding="utf-8")
    (legacy_company_dir / "log_tomados.txt").write_text("tomados", encoding="utf-8")
    (legacy_company_dir / "log_fechamento_manual.txt").write_text("manual", encoding="utf-8")
    (legacy_company_dir / "_debug").mkdir(parents=True)

    artifacts = service.get_company_artifacts(runtime_paths, empresa, competencia_dir_name="01.2026")

    assert artifacts.log_downloads == legacy_company_dir / "log_downloads_nfse.txt"
    assert artifacts.log_tomados == legacy_company_dir / "log_tomados.txt"
    assert artifacts.log_manual == legacy_company_dir / "log_fechamento_manual.txt"
    assert artifacts.debug_dir == legacy_company_dir / "_debug"
