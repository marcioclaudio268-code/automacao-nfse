from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

try:
    from openpyxl import Workbook  # noqa: F401

    RESUMO_USA_XLSX = True
except Exception:  # pragma: no cover - fallback environment without openpyxl
    RESUMO_USA_XLSX = False


@dataclass(slots=True)
class RuntimePaths:
    output_base_dir: Path
    report_path: Path
    checkpoint_path: Path
    summary_path: Path
    downloads_dir: Path

    legacy_report_path: Path
    legacy_checkpoint_path: Path
    legacy_downloads_dir: Path


def sufixo_lote_execucao(empresa_inicio: str | int | None = None, empresa_fim: str | int | None = None) -> str:
    raw_inicio = "" if empresa_inicio is None else str(empresa_inicio).strip()
    raw_fim = "" if empresa_fim is None else str(empresa_fim).strip()

    if not raw_inicio and not raw_fim:
        return ""

    if not raw_inicio or not raw_fim:
        return ""

    try:
        inicio = int(raw_inicio)
        fim = int(raw_fim)
    except Exception:
        return ""

    if inicio <= 0 or fim <= 0 or inicio > fim:
        return ""

    largura = max(3, len(str(max(inicio, fim))))
    return f"__lote_{inicio:0{largura}d}_{fim:0{largura}d}"


def build_runtime_paths(
    project_dir: Path,
    output_base_dir: Path,
    empresa_inicio: str | int | None = None,
    empresa_fim: str | int | None = None,
) -> RuntimePaths:
    sufixo = sufixo_lote_execucao(empresa_inicio, empresa_fim)
    resumo_ext = ".xlsx" if RESUMO_USA_XLSX else ".csv"
    return RuntimePaths(
        output_base_dir=output_base_dir,
        report_path=output_base_dir / f"report_execucao_empresas{sufixo}.csv",
        checkpoint_path=output_base_dir / f"checkpoint_execucao_empresas{sufixo}.json",
        summary_path=output_base_dir / f"resumo_execucao_empresas{sufixo}{resumo_ext}",
        downloads_dir=output_base_dir / "downloads",
        legacy_report_path=project_dir / "report_execucao_empresas.csv",
        legacy_checkpoint_path=project_dir / "checkpoint_execucao_empresas.json",
        legacy_downloads_dir=project_dir / "downloads",
    )


def choose_existing(preferred: Path, legacy: Path) -> Path:
    if preferred.exists():
        return preferred
    return legacy
