import os
from pathlib import Path
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QScrollArea, QTableWidgetItem

import ui.main_window as main_window_module
from ui.main_window import MainWindow
from core.paths import RESUMO_USA_XLSX


def test_main_window_usa_scroll_area_e_alturas_minimas():
    app = QApplication.instance() or QApplication([])
    window = MainWindow()

    try:
        scroll = window.findChild(QScrollArea)
        assert scroll is not None
        assert window.table.minimumHeight() >= 220
        assert window.log_edit.minimumHeight() >= 140
    finally:
        window.close()


def test_main_window_apurar_completo_forca_xml_e_desliga_fechamento():
    app = QApplication.instance() or QApplication([])
    window = MainWindow()

    try:
        window.chk_xml.setChecked(False)
        window.chk_livros.setChecked(True)
        window.chk_iss.setChecked(True)
        window.chk_pausa_manual_final.setChecked(True)

        window.chk_apurar_completo.setChecked(True)
        app.processEvents()

        assert window.chk_xml.isChecked() is True
        assert window.chk_livros.isChecked() is False
        assert window.chk_iss.isChecked() is False
        assert window.chk_pausa_manual_final.isChecked() is False
        assert window.chk_livros.isEnabled() is False
        assert window.chk_iss.isEnabled() is False
        assert "Apuracao completa ativa" in window.lbl_execution_profile_hint.text()
    finally:
        window.close()


def test_main_window_atalhos_abrem_competencia_e_geral(monkeypatch, tmp_path):
    app = QApplication.instance() or QApplication([])
    window = MainWindow()

    try:
        empresa = {
            "codigo": "833",
            "razao_social": "A. B. S. REPRESENTACOES",
            "cnpj": "41.570.055.0001-14",
        }
        competencia_dir = tmp_path / "833_B_REPRESENTACOES" / "02.2026"
        geral_dir = competencia_dir / "_GERAL"
        artifacts = SimpleNamespace(
            pasta_empresa=tmp_path / "833_B_REPRESENTACOES",
            competencia_dir=competencia_dir,
            geral_dir=geral_dir,
            evidencias_dir=geral_dir / "_evidencias",
            log_downloads=geral_dir / "log_downloads_nfse.txt",
            log_tomados=geral_dir / "log_tomados.txt",
            log_manual=geral_dir / "log_fechamento_manual.txt",
            debug_dir=geral_dir / "_debug",
        )
        abertos = []

        monkeypatch.setattr(window, "_refresh_runtime_paths", lambda: None)
        monkeypatch.setattr(
            window.artifact_service,
            "get_company_artifacts",
            lambda *_args, **_kwargs: artifacts,
        )
        monkeypatch.setattr(main_window_module, "open_in_os", lambda path: abertos.append(Path(path)))

        window.table.setRowCount(1)
        window.table.setItem(0, window.COL_CNPJ, QTableWidgetItem(empresa["cnpj"]))
        window.table.setCurrentCell(0, window.COL_CNPJ)
        window.company_data_by_cnpj["41570055000114"] = empresa
        window.apuracao_edit.setText("03/2026")
        app.processEvents()

        window.open_selected_company_target("pasta")
        window.open_selected_company_target("evidencias")
        window.open_selected_company_target("debug")
        window.open_selected_company_target("log_nfse")
        window.open_selected_company_target("log_tomados")
        window.open_selected_company_target("log_manual")

        assert abertos == [
            competencia_dir,
            geral_dir / "_evidencias",
            geral_dir / "_debug",
            geral_dir / "log_downloads_nfse.txt",
            geral_dir / "log_tomados.txt",
            geral_dir / "log_fechamento_manual.txt",
        ]
    finally:
        window.close()


def test_main_window_exibe_filtros_e_monta_paths_de_lote(tmp_path):
    app = QApplication.instance() or QApplication([])
    window = MainWindow()

    try:
        csv_path = tmp_path / "empresas.csv"
        csv_path.write_text(
            "Codigo;Razao Social;CNPJ;Segmento;Senha Prefeitura\n"
            "1;Empresa A;12.345.678/0001-90;SERVICOS;segredo\n",
            encoding="utf-8",
        )
        output_dir = tmp_path / "saida"

        assert window.empresa_inicio_edit.placeholderText() == "Inicio"
        assert window.empresa_fim_edit.placeholderText() == "Fim"
        assert window.empresas_edit.placeholderText() == "101,115,833"
        assert window.filtrar_erro_combo.count() >= 8

        window.planilha_edit.setText(str(csv_path))
        window.output_edit.setText(str(output_dir))
        window.empresa_inicio_edit.setText("1")
        window.empresa_fim_edit.setText("100")
        window.empresas_edit.setText("101,115,833")
        window.filtrar_erro_combo.setCurrentText("LOGIN_INVALIDO")

        cfg = window.build_config()

        assert cfg.empresa_inicio == "1"
        assert cfg.empresa_fim == "100"
        assert cfg.empresas == "101,115,833"
        assert cfg.filtrar_erro_tipo == "LOGIN_INVALIDO"
        assert window.runtime_paths.report_path.name == "report_execucao_empresas__lote_001_100.csv"
        assert window.runtime_paths.checkpoint_path.name == "checkpoint_execucao_empresas__lote_001_100.json"
        ext_resumo = ".xlsx" if RESUMO_USA_XLSX else ".csv"
        assert window.runtime_paths.summary_path.name == f"resumo_execucao_empresas__lote_001_100{ext_resumo}"
        window._set_execution_inputs_enabled(False)
        assert not window.empresa_inicio_edit.isEnabled()
        assert not window.empresa_fim_edit.isEnabled()
        assert not window.empresas_edit.isEnabled()
        assert not window.filtrar_erro_combo.isEnabled()
        assert output_dir.exists()
    finally:
        window.close()
