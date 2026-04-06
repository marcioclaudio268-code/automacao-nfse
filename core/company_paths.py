from __future__ import annotations

import re


SUFIXOS_LEGAIS = {
    "LTDA",
    "LTDA.",
    "ME",
    "EPP",
    "EIRELI",
    "S/A",
    "SA",
    "S.A",
    "S.A.",
}
_SIGLA_ALFANUM_PATTERN = re.compile(r"^(?=.*[A-Z])(?=.*\d)[A-Z0-9]{2,10}$")


def normalizar_nome_empresa(nome: str) -> str:
    nome = (nome or "").strip().upper()
    nome = re.sub(r"[\/\.\-]", " ", nome)
    nome = re.sub(r"[^A-Z0-9\s]", " ", nome)
    tokens = [token for token in re.split(r"\s+", nome) if token]
    tokens = [token for token in tokens if token not in SUFIXOS_LEGAIS and token not in {"S", "A"}]

    siglas = [token for token in tokens if _SIGLA_ALFANUM_PATTERN.match(token)]
    resto = [token for token in tokens if token not in siglas]

    nome_final = siglas + resto
    return "_".join(nome_final) if nome_final else "EMPRESA_DESCONHECIDA"


def nome_pasta_empresa_por_razao_social(razao_social: str) -> str:
    return normalizar_nome_empresa(razao_social)


def normalizar_codigo_empresa(codigo: str) -> str:
    codigo_limpo = re.sub(r"\s+", "", str(codigo or ""))
    if re.fullmatch(r"\d+\.0", codigo_limpo):
        codigo_limpo = codigo_limpo.split(".")[0]
    codigo_limpo = re.sub(r"[^A-Z0-9]+", "_", codigo_limpo.upper()).strip("_")
    return codigo_limpo


def nome_pasta_empresa_legada_por_dados(empresa: dict) -> str:
    return nome_pasta_empresa_por_razao_social((empresa or {}).get("razao_social", ""))


def nome_pasta_empresa_por_dados(empresa: dict) -> str:
    codigo = normalizar_codigo_empresa((empresa or {}).get("codigo", ""))
    nome = nome_pasta_empresa_legada_por_dados(empresa)

    if codigo and nome:
        return f"{codigo}_{nome}"
    return codigo or nome or "EMPRESA_DESCONHECIDA"
