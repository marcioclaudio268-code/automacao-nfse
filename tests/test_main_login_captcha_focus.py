from __future__ import annotations

from types import SimpleNamespace

from selenium.webdriver.common.by import By

import main


class FakeElement:
    def __init__(self, *, element_id: str = "", name: str = "", element_type: str = "text") -> None:
        self.element_id = element_id
        self.name = name
        self.element_type = element_type
        self.clicked = False
        self.sent_keys: list[tuple] = []
        self.value = ""
        self._displayed = True
        self._enabled = True

    def clear(self) -> None:
        self.value = ""

    def send_keys(self, *keys) -> None:
        self.sent_keys.append(keys)

    def click(self) -> None:
        self.clicked = True
        if hasattr(self, "_driver") and self._driver is not None:
            self._driver._active_element = self

    def is_displayed(self) -> bool:
        return self._displayed

    def is_enabled(self) -> bool:
        return self._enabled

    def get_attribute(self, attr: str) -> str:
        if attr == "id":
            return self.element_id
        if attr == "name":
            return self.name
        if attr == "type":
            return self.element_type
        return ""


class FakeSwitchTo:
    def __init__(self, driver: "FakeDriver") -> None:
        self.driver = driver

    @property
    def active_element(self) -> FakeElement | None:
        return self.driver._active_element


class FakeDriver:
    def __init__(self, elementos: dict[tuple[str, str], FakeElement], dashboard_visivel: bool = True) -> None:
        self.elementos = elementos
        self.dashboard_visivel = dashboard_visivel
        self._active_element: FakeElement | None = None
        self.switch_to = FakeSwitchTo(self)
        for element in self.elementos.values():
            element._driver = self

    def get(self, url: str) -> None:
        self.last_url = url

    def execute_script(self, script: str, *args) -> None:
        if "focus" in script and args:
            self._active_element = args[0]
        if "value=''" in script and args:
            args[0].value = ""

    def find_element(self, by: str, value: str) -> FakeElement:
        element = self.elementos.get((by, value))
        if element is None:
            raise LookupError(f"Elemento nao encontrado: {by}={value}")
        return element

    def find_elements(self, by: str, value: str) -> list[FakeElement]:
        if by == By.ID and value == main.LOGIN_CARD_DASHBOARD and self.dashboard_visivel:
            return [FakeElement(element_id=main.LOGIN_CARD_DASHBOARD)]
        element = self.elementos.get((by, value))
        return [element] if element is not None else []

    @property
    def page_source(self) -> str:
        return ""


class FakeWait:
    def __init__(self, driver: FakeDriver, timeout: int = 30) -> None:
        self.driver = driver
        self.timeout = timeout

    def until(self, condition):
        return condition(self.driver)


class RaisingWait:
    def __init__(self, driver: FakeDriver, timeout: int = 30) -> None:
        self.driver = driver
        self.timeout = timeout

    def until(self, condition):
        raise RuntimeError("timeout")


def test_posicionar_campo_captcha_manual_foca_o_input_certo(monkeypatch):
    captcha = FakeElement(name=main.LOGIN_CAMPO_CAPTCHA)
    driver = FakeDriver({(By.NAME, main.LOGIN_CAMPO_CAPTCHA): captcha})

    monkeypatch.setattr(main, "driver", driver)
    monkeypatch.setattr(main, "WebDriverWait", FakeWait)

    assert main.posicionar_campo_captcha_manual(timeout=1) is True
    assert captcha.clicked is True
    assert driver.switch_to.active_element is captcha


def test_posicionar_campo_captcha_manual_fallback_nao_quebra(monkeypatch, capsys):
    driver = FakeDriver({})

    monkeypatch.setattr(main, "driver", driver)
    monkeypatch.setattr(main, "WebDriverWait", RaisingWait)

    assert main.posicionar_campo_captcha_manual(timeout=1) is False
    saida = capsys.readouterr().out
    assert "mantendo fluxo atual" in saida


def test_preencher_login_prefeitura_chama_foco_antes_da_pausa(monkeypatch, capsys):
    usuario = FakeElement(element_id=main.LOGIN_CAMPO_USUARIO, name=main.LOGIN_CAMPO_USUARIO)
    senha = FakeElement(element_id=main.LOGIN_CAMPO_SENHA, name=main.LOGIN_CAMPO_SENHA, element_type="password")
    botao = FakeElement(element_id=main.LOGIN_BOTAO_ENTRAR, name=main.LOGIN_BOTAO_ENTRAR, element_type="button")
    captcha = FakeElement(name=main.LOGIN_CAMPO_CAPTCHA)
    driver = FakeDriver(
        {
            (By.ID, main.LOGIN_CAMPO_USUARIO): usuario,
            (By.ID, main.LOGIN_CAMPO_SENHA): senha,
            (By.ID, main.LOGIN_BOTAO_ENTRAR): botao,
            (By.NAME, main.LOGIN_CAMPO_CAPTCHA): captcha,
        },
        dashboard_visivel=True,
    )

    order: list[tuple[str, str | None]] = []

    class _Wait(FakeWait):
        pass

    monkeypatch.setattr(main, "driver", driver)
    monkeypatch.setattr(main, "wait", _Wait(driver))
    monkeypatch.setattr(main, "AUTO_LOGIN_PREFEITURA", True)
    monkeypatch.setattr(main, "posicionar_campo_captcha_manual", lambda timeout=5: order.append(("captcha", str(timeout))) or True)
    monkeypatch.setattr(main, "pausa_login_captcha_manual", lambda contexto: order.append(("pause", contexto)))
    monkeypatch.setattr(main, "salvar_print_evidencia", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "append_txt", lambda *args, **kwargs: None)
    monkeypatch.setenv("EMPRESA_CNPJ", "12.345.678/0001-90")
    monkeypatch.setenv("EMPRESA_SENHA", "segredo")

    resultado = main.preencher_login_prefeitura_se_habilitado()

    assert resultado == "dashboard"
    assert order[0][0] == "captcha"
    assert order[1][0] == "pause"
    assert "Aguardando digitacao do captcha" in order[1][1]
    saida = capsys.readouterr().out
    assert "Aguardando digitacao do captcha" in saida
