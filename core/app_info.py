from __future__ import annotations

import sys
from pathlib import Path


APP_DISPLAY_NAME = "Automacao NFSe"
APP_GUI_EXE_NAME = "AutomacaoNFSe.exe"
APP_MAIN_BUNDLE_NAME = "AutomacaoNFSe-Main"
APP_MAIN_EXE_NAME = f"{APP_MAIN_BUNDLE_NAME}.exe"
APP_ORCHESTRATOR_EXE_NAME = "AutomacaoNFSe-Orquestrador.exe"
APP_VERSION = "1.0.0-interno"
APP_ICON_RELATIVE_PATH = Path("assets") / "app.ico"


def is_frozen_app() -> bool:
    return bool(getattr(sys, "frozen", False))


def project_dir() -> Path:
    if is_frozen_app():
        return installed_app_dir()
    return Path(__file__).resolve().parents[1]


def installed_app_dir() -> Path:
    return Path(sys.executable).resolve().parent


def default_output_base_dir() -> Path:
    documents_dir = Path.home() / "Documents"
    base_dir = documents_dir if documents_dir.exists() else Path.home()
    return base_dir / APP_DISPLAY_NAME


def icon_path() -> Path | None:
    candidate = project_dir() / APP_ICON_RELATIVE_PATH
    if candidate.exists():
        return candidate

    if is_frozen_app():
        frozen_candidate = installed_app_dir() / APP_ICON_RELATIVE_PATH.name
        if frozen_candidate.exists():
            return frozen_candidate

    return None
