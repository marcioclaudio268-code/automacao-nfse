from core.event_parser import parse_log_line


def test_parse_manual_login_event():
    event = parse_log_line("[MANUAL-LOGIN] LOGIN_CAPTCHA | Resolva o captcha e clique em 'Entrar'.")

    assert event.kind == "etapa"
    assert event.etapa == "Aguardando captcha"


def test_parse_manual_login_auto_monitoring_event():
    event = parse_log_line("[MANUAL-LOGIN] MONITORAMENTO_AUTOMATICO | Apos clicar em 'Entrar', a automacao segue sozinha.")

    assert event.kind == "etapa"
    assert event.etapa == "Aguardando captcha"
