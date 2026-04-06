from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from core.company_paths import nome_pasta_empresa_legada_por_dados, nome_pasta_empresa_por_dados


@dataclass
class CompanyArtifacts:
    pasta_empresa: Path
    competencia_dir: Path
    geral_dir: Path
    evidencias_dir: Path
    log_downloads: Path
    log_tomados: Path
    log_manual: Path
    debug_dir: Path


class ArtifactLocatorService:
    _COMPETENCIA_DIR_RE = re.compile(r"^\d{2}\.\d{4}$")
    _EVIDENCIAS_CANDIDATOS = (
        Path("_GERAL") / "_evidencias",
        Path("PRESTADOS") / "_evidencias",
        Path("TOMADOS") / "_evidencias",
        Path("_evidencias"),
    )

    def get_company_folder_name(self, empresa: dict) -> str:
        return nome_pasta_empresa_por_dados(empresa)

    def get_company_folder_candidates(self, empresa: dict) -> list[str]:
        nomes = []
        for nome in (
            nome_pasta_empresa_por_dados(empresa),
            nome_pasta_empresa_legada_por_dados(empresa),
        ):
            if nome and nome not in nomes:
                nomes.append(nome)
        return nomes

    def get_company_artifacts(
        self,
        runtime_paths,
        empresa: dict,
        competencia_dir_name: str = "",
    ) -> CompanyArtifacts:
        pasta_empresa = None
        candidatos = []
        for pasta_nome in self.get_company_folder_candidates(empresa):
            candidatos.append(runtime_paths.downloads_dir / pasta_nome)
        for pasta_nome in self.get_company_folder_candidates(empresa):
            candidatos.append(runtime_paths.legacy_downloads_dir / pasta_nome)

        for candidato in candidatos:
            if candidato.exists():
                pasta_empresa = candidato
                break

        if pasta_empresa is None:
            pasta_empresa = runtime_paths.downloads_dir / self.get_company_folder_name(empresa)

        competencia_dir = self._resolve_competencia_dir(pasta_empresa, competencia_dir_name)
        geral_dir = self._resolve_geral_dir(competencia_dir, pasta_empresa)
        evidencias_dir = self._resolve_evidencias_dir(competencia_dir, pasta_empresa, geral_dir)

        return CompanyArtifacts(
            pasta_empresa=pasta_empresa,
            competencia_dir=competencia_dir,
            geral_dir=geral_dir,
            evidencias_dir=evidencias_dir,
            log_downloads=self._resolve_artifact_file(geral_dir, pasta_empresa, "log_downloads_nfse.txt"),
            log_tomados=self._resolve_artifact_file(geral_dir, pasta_empresa, "log_tomados.txt"),
            log_manual=self._resolve_artifact_file(geral_dir, pasta_empresa, "log_fechamento_manual.txt"),
            debug_dir=self._resolve_debug_dir(geral_dir, pasta_empresa),
        )

    def _resolve_competencia_dir(self, pasta_empresa: Path, competencia_dir_name: str) -> Path:
        competencia_dir_name = (competencia_dir_name or "").strip()
        if competencia_dir_name:
            candidato = pasta_empresa / competencia_dir_name
            if candidato.exists():
                return candidato

        competencias_existentes = []
        if pasta_empresa.exists():
            for child in pasta_empresa.iterdir():
                if child.is_dir() and self._COMPETENCIA_DIR_RE.fullmatch(child.name):
                    competencias_existentes.append(child)

        if competencias_existentes:
            competencias_existentes.sort(
                key=lambda path: (int(path.name.split(".")[1]), int(path.name.split(".")[0])),
                reverse=True,
            )
            return competencias_existentes[0]

        if competencia_dir_name:
            return pasta_empresa / competencia_dir_name
        return pasta_empresa

    def _resolve_geral_dir(self, competencia_dir: Path, pasta_empresa: Path) -> Path:
        candidatos = (
            competencia_dir / "_GERAL",
            pasta_empresa / "_GERAL",
        )
        for candidato in candidatos:
            if candidato.exists():
                return candidato
        return competencia_dir / "_GERAL"

    def _resolve_artifact_file(self, geral_dir: Path, pasta_empresa: Path, filename: str) -> Path:
        candidatos = (
            geral_dir / filename,
            pasta_empresa / filename,
        )
        for candidato in candidatos:
            if candidato.exists():
                return candidato
        return geral_dir / filename

    def _resolve_debug_dir(self, geral_dir: Path, pasta_empresa: Path) -> Path:
        candidatos = (
            geral_dir / "_debug",
            pasta_empresa / "_debug",
        )
        for candidato in candidatos:
            if candidato.exists():
                return candidato
        return geral_dir / "_debug"

    def _resolve_evidencias_dir(self, competencia_dir: Path, pasta_empresa: Path, geral_dir: Path) -> Path:
        candidatos_absolutos = (
            geral_dir / "_evidencias",
            pasta_empresa / "_GERAL" / "_evidencias",
        )
        for candidato in candidatos_absolutos:
            if candidato.exists():
                return candidato
        for rel_path in self._EVIDENCIAS_CANDIDATOS:
            candidato = competencia_dir / rel_path
            if candidato.exists():
                return candidato
        return geral_dir / "_evidencias"
