from pathlib import Path

from core.config_runtime import ExecutionProfile, RuntimeConfig, mes_atual_str
from core.paths import build_runtime_paths
from ui.controller import ExecutionController


def _build_cfg(tmp_path: Path, **overrides) -> RuntimeConfig:
    empresas_arquivo = tmp_path / "empresas.csv"
    empresas_arquivo.write_text(
        "Codigo;Razao Social;CNPJ;Segmento;Senha Prefeitura\n"
        "1;Empresa A;12.345.678/0001-90;SERVICOS;segredo\n",
        encoding="utf-8",
    )
    output_dir = tmp_path / "saida"
    output_dir.mkdir()
    cfg = RuntimeConfig(
        empresas_arquivo=empresas_arquivo,
        apuracao_referencia=mes_atual_str(),
        output_base_dir=output_dir,
        execution_profile=ExecutionProfile(),
        **overrides,
    )
    cfg.validate()
    return cfg


def test_controller_repasse_filtros_para_o_processo(tmp_path):
    runtime_paths = build_runtime_paths(tmp_path / "project", tmp_path / "saida", "1", "100")
    controller = ExecutionController(tmp_path / "orq.py", tmp_path / "project", runtime_paths)
    controller.control_dir = tmp_path / "saida" / "_control" / "run_0001"

    cfg = _build_cfg(
        tmp_path,
        empresa_inicio="1",
        empresa_fim="100",
        empresas="101,115,833",
        filtrar_erro_tipo="LOGIN_INVALIDO",
    )

    env_vars = controller.build_process_env_vars(cfg)

    assert env_vars["EMPRESA_INICIO"] == "1"
    assert env_vars["EMPRESA_FIM"] == "100"
    assert env_vars["EMPRESAS"] == "101,115,833"
    assert env_vars["FILTRAR_ERRO_TIPO"] == "LOGIN_INVALIDO"
    assert env_vars["EXECUTION_CONTROL_DIR"] == str(controller.control_dir)
    assert env_vars["OUTPUT_BASE_DIR"] == str(cfg.output_base_dir)
    assert env_vars["PERFIL_EXECUCAO_ATIVO"] == "0"


def test_controller_env_vars_ficam_vazias_sem_filtros(tmp_path):
    runtime_paths = build_runtime_paths(tmp_path / "project", tmp_path / "saida")
    controller = ExecutionController(tmp_path / "orq.py", tmp_path / "project", runtime_paths)
    controller.control_dir = tmp_path / "saida" / "_control" / "run_0002"

    cfg = _build_cfg(tmp_path)

    env_vars = controller.build_process_env_vars(cfg)

    assert env_vars["EMPRESA_INICIO"] == ""
    assert env_vars["EMPRESA_FIM"] == ""
    assert env_vars["EMPRESAS"] == ""
    assert env_vars["FILTRAR_ERRO_TIPO"] == ""
