from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class RuntimePaths:
    output_base_dir: Path
    report_path: Path
    checkpoint_path: Path
    downloads_dir: Path

    legacy_report_path: Path
    legacy_checkpoint_path: Path
    legacy_downloads_dir: Path


def build_runtime_paths(project_dir: Path, output_base_dir: Path) -> RuntimePaths:
    return RuntimePaths(
        output_base_dir=output_base_dir,
        report_path=output_base_dir / "report_execucao_empresas.csv",
        checkpoint_path=output_base_dir / "checkpoint_execucao_empresas.json",
        downloads_dir=output_base_dir / "downloads",
        legacy_report_path=project_dir / "report_execucao_empresas.csv",
        legacy_checkpoint_path=project_dir / "checkpoint_execucao_empresas.json",
        legacy_downloads_dir=project_dir / "downloads",
    )


def choose_existing(preferred: Path, legacy: Path) -> Path:
    if preferred.exists():
        return preferred
    return legacy