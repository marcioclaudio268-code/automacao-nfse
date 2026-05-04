from pathlib import Path

import main


def test_capturar_print_passo_robo_flag_desligada_nao_chama_screenshot(monkeypatch):
    chamadas = []

    class DriverFake:
        def save_screenshot(self, destino):
            chamadas.append(destino)
            return True

    monkeypatch.setattr(main, "CAPTURAR_PRINTS_PASSOS_ROBO", False)
    monkeypatch.setattr(main, "driver", DriverFake())

    resultado = main.capturar_print_passo_robo("PRESTADOS", "abriu_lista")

    assert resultado == ""
    assert chamadas == []


def test_capturar_print_passo_robo_salva_na_pasta_do_modulo(monkeypatch, tmp_path):
    destinos = []
    logs = []

    class DriverFake:
        def save_screenshot(self, destino):
            destinos.append(destino)
            Path(destino).write_bytes(b"png")
            return True

    monkeypatch.setattr(main, "CAPTURAR_PRINTS_PASSOS_ROBO", True)
    monkeypatch.setattr(main, "_PRINT_PASSO_ROBO_SEQ", 0)
    monkeypatch.setattr(main, "driver", DriverFake())
    monkeypatch.setattr(main, "pasta_servico", lambda modulo, **_kwargs: str(tmp_path / modulo))
    monkeypatch.setattr(main, "append_txt", lambda _path, msg: logs.append(msg))

    resultado = main.capturar_print_passo_robo(
        "TOMADOS",
        "abriu modulo/lista",
        ano_alvo=2026,
        mes_alvo=5,
        ref_alvo="05/2026",
        log_path=str(tmp_path / "log.txt"),
    )

    caminho = Path(resultado)
    assert destinos == [resultado]
    assert caminho.parent == tmp_path / "TOMADOS" / "_prints_robo"
    assert caminho.name.startswith("0001_TOMADOS_abriu_modulo_lista_")
    assert any("PRINT_PASSO=OK | MODULO=TOMADOS | ETAPA=abriu modulo/lista" in msg for msg in logs)


def test_capturar_print_passo_robo_erro_screenshot_nao_quebra(monkeypatch, tmp_path):
    logs = []

    class DriverFake:
        def save_screenshot(self, _destino):
            raise RuntimeError("falha simulada")

    monkeypatch.setattr(main, "CAPTURAR_PRINTS_PASSOS_ROBO", True)
    monkeypatch.setattr(main, "driver", DriverFake())
    monkeypatch.setattr(main, "pasta_servico", lambda modulo, **_kwargs: str(tmp_path / modulo))
    monkeypatch.setattr(main, "append_txt", lambda _path, msg: logs.append(msg))

    resultado = main.capturar_print_passo_robo(
        "PRESTADOS",
        "erro tecnico",
        ano_alvo=2026,
        mes_alvo=5,
        log_path=str(tmp_path / "log.txt"),
    )

    assert resultado == ""
    assert any("PRINT_PASSO=ERRO | MODULO=PRESTADOS | ETAPA=erro tecnico" in msg for msg in logs)
