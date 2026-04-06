import os
from pathlib import Path
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QScrollArea, QTableWidgetItem

import ui.main_window as main_window_module
from ui.main_window import MainWindow


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
