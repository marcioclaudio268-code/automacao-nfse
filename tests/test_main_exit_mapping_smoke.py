import main


def test_map_runtime_error_to_exit_code_captcha_timeout():
    code = main.map_runtime_error_to_exit_code(main.MSG_CAPTCHA_TIMEOUT)
    assert code == main.EXIT_CODE_CAPTCHA_TIMEOUT


def test_map_runtime_error_to_exit_code_captcha_incorreto():
    code = main.map_runtime_error_to_exit_code(main.MSG_CAPTCHA_INCORRETO)
    assert code == main.EXIT_CODE_CAPTCHA_TIMEOUT


def test_map_runtime_error_to_exit_code_sem_competencia():
    code = main.map_runtime_error_to_exit_code(main.MSG_SEM_COMPETENCIA)
    assert code == main.EXIT_CODE_SEM_COMPETENCIA


def test_map_runtime_error_to_exit_code_sem_servicos():
    code = main.map_runtime_error_to_exit_code(main.MSG_SEM_SERVICOS)
    assert code == main.EXIT_CODE_SEM_SERVICOS


def test_map_runtime_error_to_exit_code_credencial():
    code = main.map_runtime_error_to_exit_code(main.MSG_CREDENCIAL_INVALIDA)
    assert code == main.EXIT_CODE_CREDENCIAL_INVALIDA


def test_map_runtime_error_to_exit_code_empresa_multipla():
    code = main.map_runtime_error_to_exit_code(main.MSG_EMPRESA_MULTIPLA)
    assert code == main.EXIT_CODE_EMPRESA_MULTIPLA


def test_map_runtime_error_to_exit_code_multi_cadastro_revisao_manual():
    code = main.map_runtime_error_to_exit_code(main.MSG_MULTI_CADASTRO_REVISAO_MANUAL)
    assert code == main.EXIT_CODE_MULTI_CADASTRO_REVISAO_MANUAL


def test_map_runtime_error_to_exit_code_tomados():
    code = main.map_runtime_error_to_exit_code(main.MSG_TOMADOS_FALHA)
    assert code == main.EXIT_CODE_TOMADOS_FALHA


def test_map_runtime_error_to_exit_code_chrome_init():
    code = main.map_runtime_error_to_exit_code(main.MSG_CHROME_INIT_FALHA)
    assert code == main.EXIT_CODE_CHROME_INIT_FALHA


def test_map_runtime_error_to_exit_code_apuracao_completa_revisao_manual():
    code = main.map_runtime_error_to_exit_code(main.MSG_APURACAO_COMPLETA_REVISAO_MANUAL)
    assert code == main.EXIT_CODE_APURACAO_COMPLETA_REVISAO_MANUAL


def test_map_runtime_error_to_exit_code_desconhecido():
    assert main.map_runtime_error_to_exit_code("ERRO_QUALQUER") is None
