import main
from selenium.common.exceptions import StaleElementReferenceException


def test_executar_guia_iss_com_retry_stale_tenta_novamente(monkeypatch):
    logs = []
    preparos = []
    sleeps = []
    chamadas = {"total": 0}

    monkeypatch.setattr(main, "append_txt", lambda _path, msg: logs.append(msg))
    monkeypatch.setattr(main, "_preparar_tentativa_guia_iss", lambda handles: preparos.append(list(handles or [])))
    monkeypatch.setattr(main, "sleep", lambda delay: sleeps.append(delay))

    def operacao():
        chamadas["total"] += 1
        if chamadas["total"] == 1:
            raise StaleElementReferenceException("dom redesenhado")
        return "OK"

    resultado = main.executar_guia_iss_com_retry_stale(
        operacao,
        "TOMADOS",
        "04/2026",
        "log-manual.txt",
        handles_base=["janela-principal"],
    )

    assert resultado == "OK"
    assert chamadas["total"] == 2
    assert preparos == [["janela-principal"], ["janela-principal"]]
    assert sleeps == [0.8]
    assert any("GUIA_ISS=TENTATIVA | tentativa 1/3" in msg for msg in logs)
    assert any("GUIA_ISS=STALE | tentativa 1/3" in msg for msg in logs)
    assert any("GUIA_ISS=TENTATIVA | tentativa 2/3" in msg for msg in logs)
