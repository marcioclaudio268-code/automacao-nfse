from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    WebDriverException,
)
from time import sleep
import os, time, re, csv, tempfile
import sys
import json
import base64
import html
import math
import ssl
import unicodedata
from datetime import datetime
import xml.etree.ElementTree as ET
import shutil
from urllib import request as urllib_request
from urllib import parse as urllib_parse
from urllib import error as urllib_error
from pathlib import Path

from core.company_paths import normalizar_nome_empresa, normalizar_codigo_empresa
from core.config_runtime import competencias_alvo as runtime_competencias_alvo, validate_apuracao_referencia

# =====================
# CONFIGURACAO
# =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_BASE_DIR = os.environ.get("OUTPUT_BASE_DIR", BASE_DIR)
os.makedirs(OUTPUT_BASE_DIR, exist_ok=True)

DOWNLOAD_DIR = os.path.join(OUTPUT_BASE_DIR, "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

TEMP_DOWNLOAD_DIR = os.path.join(DOWNLOAD_DIR, "_tmp")
os.makedirs(TEMP_DOWNLOAD_DIR, exist_ok=True)

EXECUTION_CONTROL_DIR = os.environ.get("EXECUTION_CONTROL_DIR", "").strip()
MANUAL_SIGNAL_POLL_SECONDS = float(os.environ.get("MANUAL_SIGNAL_POLL_SECONDS", "0.5").strip() or "0.5")


def _env_flag(name: str, default: bool) -> bool:
    return os.getenv(name, "1" if default else "0").strip() == "1"


# =====================
# MODO DE DEPURACAO / ETAPAS (CADEADO + PAUSA MANUAL)
# =====================
MODO_APENAS_CADEADO = os.getenv("MODO_APENAS_CADEADO", "0").strip() == "1"
# Quanto tempo deixar o Chrome aberto para você baixar manualmente os livros (segundos).
# - Se 0, o script espera você apertar ENTER no terminal para continuar.
PAUSA_MANUAL_SEGUNDOS = int(os.getenv("PAUSA_MANUAL_SEGUNDOS", "0").strip() or "0")
# Se 1, pausa também entre Tomados e Prestados no modo de cadeado.
# Padrão 0: executa Tomados -> Prestados direto, pausando somente no final.
PAUSAR_ENTRE_MODULOS_CADEADO = os.getenv("PAUSAR_ENTRE_MODULOS_CADEADO", "0").strip() == "1"
# Se 1, mantém o fluxo atual (Prestados XML + Tomados XML) e, ao final, clica no cadeado vermelho
# em Serviços Tomados e Serviços Prestados (na competência alvo), pausando para você baixar os livros manualmente.
ACRESCENTAR_CADEADO_E_PAUSA_FINAL = os.getenv("ACRESCENTAR_CADEADO_E_PAUSA_FINAL", "1").strip() == "1"
PERFIL_EXECUCAO_ATIVO = os.getenv("PERFIL_EXECUCAO_ATIVO", "0").strip() == "1"
EXECUTAR_PRESTADOS = _env_flag("EXECUTAR_PRESTADOS", True)
EXECUTAR_TOMADOS = _env_flag("EXECUTAR_TOMADOS", True)
EXECUTAR_XML = _env_flag("EXECUTAR_XML", not MODO_APENAS_CADEADO)
EXECUTAR_LIVROS = _env_flag("EXECUTAR_LIVROS", MODO_APENAS_CADEADO or ACRESCENTAR_CADEADO_E_PAUSA_FINAL)
EXECUTAR_ISS = _env_flag("EXECUTAR_ISS", MODO_APENAS_CADEADO or ACRESCENTAR_CADEADO_E_PAUSA_FINAL)
PAUSA_MANUAL_FINAL = _env_flag("PAUSA_MANUAL_FINAL", MODO_APENAS_CADEADO or ACRESCENTAR_CADEADO_E_PAUSA_FINAL)
MODO_AUTOMATICO = _env_flag("MODO_AUTOMATICO", False)
APURAR_COMPLETO = _env_flag("APURAR_COMPLETO", False)

if APURAR_COMPLETO:
    EXECUTAR_XML = True
    EXECUTAR_LIVROS = False
    EXECUTAR_ISS = False
    PAUSA_MANUAL_FINAL = False

if MODO_AUTOMATICO:
    PAUSA_MANUAL_FINAL = False

if EXECUTAR_ISS and not EXECUTAR_LIVROS:
    print("EXECUTAR_ISS=1 ignorado porque EXECUTAR_LIVROS=0.")
    EXECUTAR_ISS = False

if PAUSA_MANUAL_FINAL and not EXECUTAR_LIVROS:
    print("PAUSA_MANUAL_FINAL=1 ignorado porque EXECUTAR_LIVROS=0.")
    PAUSA_MANUAL_FINAL = False

ENABLE_CHROME_PERFORMANCE_LOGS = _env_flag(
    "ENABLE_CHROME_PERFORMANCE_LOGS",
    EXECUTAR_TOMADOS or EXECUTAR_LIVROS or EXECUTAR_ISS,
)
PERF_TIMING_LOGS = _env_flag("PERF_TIMING_LOGS", False)



def append_txt(path: str, line: str):
    """Append a human-readable line to a .txt log (one event per line)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    safe = (line or "").replace("\n", " ").replace("\r", " ").strip()
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {safe}\n")


def perf_log(etapa: str, inicio: float, detalhes: str = ""):
    if not PERF_TIMING_LOGS:
        return
    decorrido = max(0.0, time.perf_counter() - inicio)
    sufixo = f" | {detalhes}" if detalhes else ""
    print(f"[PERF] {etapa} | {decorrido:.2f}s{sufixo}")


def bounded_timeout(timeout: int | float, cap: int, floor: int = 1) -> int:
    try:
        total = int(math.ceil(float(timeout)))
    except Exception:
        total = cap
    return max(floor, min(total, cap))


def aguardar_enter_robusto(prompt: str, tag: str = "MANUAL") -> bool:
    """Aguarda ENTER com fallback no Windows quando stdin nao esta em TTY."""
    try:
        if sys.stdin is not None and hasattr(sys.stdin, "isatty") and sys.stdin.isatty():
            input(prompt)
            return True
    except Exception:
        pass

    if os.name == "nt":
        try:
            import msvcrt
            print(f"[{tag}] stdin nao interativo detectado. Pressione ENTER para continuar...")
            while True:
                ch = msvcrt.getwch()
                if ch in ("\r", "\n"):
                    return True
        except Exception as e:
            print(f"[{tag}] Falha no fallback de teclado ({type(e).__name__}).")

    print(f"[{tag}] Entrada interativa indisponivel neste modo de execucao.")
    print(f"[{tag}] Rode pelo terminal: python main.py")
    return False


def _pasta_controle_manual() -> str:
    if not EXECUTION_CONTROL_DIR:
        return ""

    empresa_pasta = globals().get("EMPRESA_PASTA") or os.environ.get("EMPRESA_PASTA_FORCADA") or "SEM_EMPRESA"
    control_dir = os.path.join(EXECUTION_CONTROL_DIR, empresa_pasta)
    os.makedirs(control_dir, exist_ok=True)
    return control_dir


def aguardar_liberacao_manual_app(contexto: str, tag: str = "MANUAL") -> bool:
    control_dir = _pasta_controle_manual()
    if not control_dir:
        return False

    continue_path = os.path.join(control_dir, "continue.signal")
    state_path = os.path.join(control_dir, "manual_wait.json")
    try:
        if os.path.exists(continue_path):
            os.remove(continue_path)
    except Exception:
        pass

    payload = {
        "tag": tag,
        "contexto": contexto,
        "empresa_pasta": globals().get("EMPRESA_PASTA") or "",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    try:
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[{tag}] Falha ao registrar estado manual no controle do app ({type(e).__name__}).")
        return False

    print(f"[{tag}] AGUARDANDO_LIBERACAO_APP | {contexto}")
    try:
        while True:
            if os.path.exists(continue_path):
                try:
                    os.remove(continue_path)
                except Exception:
                    pass
                print(f"[{tag}] CONTINUANDO_EXECUCAO_MANUAL | {contexto}")
                return True
            sleep(max(0.2, MANUAL_SIGNAL_POLL_SECONDS))
    finally:
        try:
            if os.path.exists(state_path):
                os.remove(state_path)
        except Exception:
            pass
MAX_RETRIES_POR_NOTA = 3
WAIT_LISTA_TIMEOUT = int(os.environ.get("WAIT_LISTA_TIMEOUT", "20").strip() or "20")
GRID_POLL_INTERVAL_SECONDS = float(os.environ.get("GRID_POLL_INTERVAL_SECONDS", "0.2").strip() or "0.2")
DOWNLOAD_POLL_INTERVAL_SECONDS = float(os.environ.get("DOWNLOAD_POLL_INTERVAL_SECONDS", "0.2").strip() or "0.2")
NETWORK_LOG_POLL_INTERVAL_SECONDS = float(os.environ.get("NETWORK_LOG_POLL_INTERVAL_SECONDS", "0.1").strip() or "0.1")
GRID_SCROLL_SETTLE_SECONDS = float(os.environ.get("GRID_SCROLL_SETTLE_SECONDS", "0.03").strip() or "0.03")
SELECT_OPTION_SETTLE_TIMEOUT = float(os.environ.get("SELECT_OPTION_SETTLE_TIMEOUT", "1.5").strip() or "1.5")
SUBMIT_FORM_SETTLE_MS = int(os.environ.get("SUBMIT_FORM_SETTLE_MS", "20").strip() or "20")
PDF_POPUP_WAIT_TIMEOUT = int(os.environ.get("PDF_POPUP_WAIT_TIMEOUT", "10").strip() or "10")
PDF_FILE_FALLBACK_TIMEOUT = int(os.environ.get("PDF_FILE_FALLBACK_TIMEOUT", "6").strip() or "6")
PDF_NETWORK_CAPTURE_TIMEOUT = int(os.environ.get("PDF_NETWORK_CAPTURE_TIMEOUT", "8").strip() or "8")
LOG_FILENAME = "log_downloads_nfse.txt"
TOMADOS_LOG_FILENAME = "log_tomados.txt"
TOMADOS_XML_BASENAME = "SERVICOS_TOMADOS_{mm_aaaa}.xml"  # ex: SERVICOS_TOMADOS_01-2026.xml
TOMADOS_NFSE_PDF_BASENAME = "NFT_{nf}_{data}.pdf"
LIVRO_PRESTADOS_PDF_BASENAME = "LIVRO_SERVICOS_PRESTADOS_{mm_aaaa}.pdf"
LIVRO_TOMADOS_PDF_BASENAME = "LIVRO_SERVICOS_TOMADOS_{mm_aaaa}.pdf"
GUIA_ISS_PRESTADOS_PDF_BASENAME = "GUIA_ISS_PRESTADOS_{mm_aaaa}.pdf"
GUIA_ISS_TOMADOS_PDF_BASENAME = "GUIA_ISS_TOMADOS_{mm_aaaa}.pdf"
LIVRO_PDF_FALLBACK_CHROME = os.getenv("LIVRO_PDF_FALLBACK_CHROME", "1").strip() == "1"

# Se quiser travar a empresa SEM depender do texto do site, descomente:
# EMPRESA_PASTA_FORCADA = "H2_IMOBILIARIA"
EMPRESA_PASTA_FORCADA = os.environ.get("EMPRESA_PASTA_FORCADA", "")

# Competência alvo: por padrão, mês anterior ao mês atual (apuração).
# Pode sobrescrever com APURACAO_REFERENCIA=MM/AAAA (ex.: 03/2026 -> alvo 02/2026).
APURACAO_REFERENCIA = os.environ.get("APURACAO_REFERENCIA", "").strip()

PARAR_PROCESSAMENTO = False
ENCONTROU_MES_ALVO = False
CONT_FORA_APOS_ALVO = 0
CONT_FORA_ANTES_ALVO = 0
SEM_COMPETENCIA_NA_EMPRESA = False
PRESTADOS_SEM_MODULO = False
chaves_ok = set()  # cache das CHAVE_UNICA com STATUS=OK (carregado do log)
CONTEXTO_CADASTRO_ATUAL = {}
MULTI_CADASTRO_ATIVO = False
MULTI_CADASTRO_TARDIO_ESTADO = {
    "ativo": False,
    "cadastros": [],
    "indice_atual": -1,
    "resultados": [],
    "concluidos": [],
    "falhos": [],
}
LIMITE_HEURISTICA_FORA_ALVO = int(os.environ.get("LIMITE_HEURISTICA_FORA_ALVO", "2"))
STRICT_LISTA_INICIAL = os.environ.get("STRICT_LISTA_INICIAL", "0").strip() == "1"
MSG_CAPTCHA_TIMEOUT = "CAPTCHA_NAO_RESOLVIDO_NO_TEMPO"
MSG_CAPTCHA_INCORRETO = "CAPTCHA_INCORRETO"
MSG_SEM_COMPETENCIA = "SUCESSO_SEM_COMPETENCIA"
MSG_SEM_SERVICOS = "SUCESSO_SEM_SERVICOS"
MSG_CREDENCIAL_INVALIDA = "CREDENCIAL_INVALIDA"
MSG_EMPRESA_MULTIPLA = "EMPRESA_MULTIPLA"
EXIT_CODE_CAPTCHA_TIMEOUT = 30
EXIT_CODE_SEM_COMPETENCIA = 40
EXIT_CODE_SEM_SERVICOS = 41
EXIT_CODE_CREDENCIAL_INVALIDA = 50
EXIT_CODE_EMPRESA_MULTIPLA = 52
MSG_TOMADOS_PDF_REVISAO_MANUAL = "REVISAO_MANUAL_PDF_TOMADOS"
EXIT_CODE_TOMADOS_PDF_REVISAO_MANUAL = 51
MSG_APURACAO_COMPLETA_REVISAO_MANUAL = "REVISAO_MANUAL_APURACAO_COMPLETA"
EXIT_CODE_APURACAO_COMPLETA_REVISAO_MANUAL = 53
MSG_MULTI_CADASTRO_REVISAO_MANUAL = "REVISAO_MANUAL_MULTIPLOS_CADASTROS"
EXIT_CODE_MULTI_CADASTRO_REVISAO_MANUAL = 54
CODIGO_ERRO_TOMADOS_PDF_NAO_GERADO = "PDF_TOMADOS_NAO_GERADO"
CODIGO_ERRO_TOMADOS_PDF_SEM_LOCATION = "PDF_TOMADOS_SEM_LOCATION"
CODIGO_ERRO_TOMADOS_PDF_RESPOSTA_INVALIDA = "PDF_TOMADOS_RESPOSTA_INVALIDA"
ARQUIVO_STATUS_TOMADOS_PDF = "_tomados_pdf_status.json"
MSG_TOMADOS_FALHA = "FALHA_TOMADOS"
EXIT_CODE_TOMADOS_FALHA = 42
TOMADOS_OBRIGATORIO = os.environ.get("TOMADOS_OBRIGATORIO", "0").strip() == "1"

AUTO_LOGIN_PREFEITURA = os.environ.get("AUTO_LOGIN_PREFEITURA", "0").strip() == "1"
LOGIN_URL_PREFEITURA = os.environ.get(
    "LOGIN_URL_PREFEITURA",
    "https://tributario.bauru.sp.gov.br/loginCNPJContribuinte.jsp?execobj=ContribuintesWebRelacionados",
).strip()
LOGIN_CAMPO_USUARIO = os.environ.get("LOGIN_CAMPO_USUARIO", "usuario").strip()
LOGIN_CAMPO_SENHA = os.environ.get("LOGIN_CAMPO_SENHA", "senha").strip()
LOGIN_BOTAO_ENTRAR = os.environ.get("LOGIN_BOTAO_ENTRAR", "btnEntrar").strip()
LOGIN_CARD_DASHBOARD = os.environ.get("LOGIN_CARD_DASHBOARD", "divtxtnotafiscal").strip()
LOGIN_CARD_LISTA_NOTAS = os.environ.get("LOGIN_CARD_LISTA_NOTAS", "divtxtlistanf").strip()
GRID_HEADER_CACHE = {"ts": 0.0, "map": {}}
MSG_CHROME_INIT_FALHA = "FALHA_INICIALIZAR_CHROME"
EXIT_CODE_CHROME_INIT_FALHA = 60
FORCAR_SESSAO_LIMPA_LOGIN = os.environ.get("FORCAR_SESSAO_LIMPA_LOGIN", "0").strip() == "1"
ULTIMO_CHECKBOX_MARCADO = None

# =====================
# CHROME - PERFIL EXCLUSIVO
# =====================

def _default_profile_dir() -> str:
    if os.name == "nt":
        local_appdata = os.environ.get("LOCALAPPDATA", "").strip()
        if local_appdata:
            return os.path.join(local_appdata, "AutomacaoNFSe", "ChromeRobotProfile")
        return os.path.join(str(Path.home()), "AppData", "Local", "AutomacaoNFSe", "ChromeRobotProfile")
    return os.path.join(tempfile.gettempdir(), "AutomacaoNFSe", "ChromeRobotProfile")


def _runtime_chrome_bundle_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(BASE_DIR)


def _resolve_embedded_browser_assets() -> dict[str, str] | None:
    chrome_override = os.environ.get("CHROME_BINARY_PATH", "").strip()
    driver_override = os.environ.get("CHROMEDRIVER_PATH", "").strip()
    if chrome_override and driver_override:
        chrome_path = Path(chrome_override).expanduser().resolve()
        driver_path = Path(driver_override).expanduser().resolve()
        if chrome_path.exists() and driver_path.exists():
            return {
                "chrome": str(chrome_path),
                "driver": str(driver_path),
                "source": "env",
            }

    runtime_root = _runtime_chrome_bundle_root()
    candidate_roots = [
        runtime_root / "browser",
        Path(BASE_DIR) / "browser",
        Path(BASE_DIR) / "installer" / "chrome-for-testing" / "browser",
    ]
    seen: set[str] = set()
    for root in candidate_roots:
        normalized = str(root.resolve()) if root.exists() else str(root)
        if normalized in seen:
            continue
        seen.add(normalized)

        chrome_path = root / "chrome-win64" / "chrome.exe"
        driver_path = root / "chromedriver-win64" / "chromedriver.exe"
        if chrome_path.exists() and driver_path.exists():
            return {
                "chrome": str(chrome_path.resolve()),
                "driver": str(driver_path.resolve()),
                "source": str(root.resolve()),
            }

    return None


prefs = {
    "download.default_directory": TEMP_DOWNLOAD_DIR,  # importante
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False,
    "safebrowsing.disable_download_protection": True,
    # Mantem o viewer interno habilitado para que o portal exponha a URL final do PDF
    # em vez de forcar um download externo que o Chrome pode bloquear.
    "plugins.always_open_pdf_externally": False,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "credentials_enable_service": False,
    "profile.password_manager_enabled": False,
}


def _build_chrome_options(chrome_binary_path: str | None = None) -> Options:
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument(f"--user-data-dir={CHROME_PROFILE_DIR}")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--ignore-ssl-errors")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--disable-features=InsecureDownloadWarnings")
    chrome_options.add_argument("--safebrowsing-disable-download-protection")
    chrome_options.add_argument("--no-sandbox")
    if chrome_binary_path:
        chrome_options.binary_location = chrome_binary_path
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    logging_prefs = {"browser": "ALL"}
    if ENABLE_CHROME_PERFORMANCE_LOGS:
        logging_prefs["performance"] = "ALL"
    chrome_options.set_capability("goog:loggingPrefs", logging_prefs)
    return chrome_options

TEMP_PROFILE_DIR = None
DEFAULT_PROFILE_DIR = _default_profile_dir()
if FORCAR_SESSAO_LIMPA_LOGIN:
    TEMP_PROFILE_DIR = tempfile.mkdtemp(prefix="ChromeRobotProfile_")
    CHROME_PROFILE_DIR = TEMP_PROFILE_DIR
else:
    CHROME_PROFILE_DIR = os.environ.get("CHROME_PROFILE_DIR", DEFAULT_PROFILE_DIR)

driver = None
wait = None
login_wait_seconds = int(os.environ.get("LOGIN_WAIT_SECONDS", "120"))



def pausa_manual(contexto: str):
    """Mantém o Chrome aberto para ações manuais (baixar livros) sem encerrar a execução."""
    if not MODO_APENAS_CADEADO:
        return
    msg = f"[MANUAL] {contexto} | Baixe os livros manualmente (Tomados/Prestados)."
    print(msg)
    try:
        append_txt(caminho_log_manual(), msg)
    except Exception:
        pass

    if aguardar_liberacao_manual_app(contexto, tag="MANUAL"):
        try:
            append_txt(caminho_log_manual(), f"[MANUAL] CONTINUANDO_EXECUCAO_MANUAL | {contexto}")
        except Exception:
            pass
        return

    if PAUSA_MANUAL_SEGUNDOS <= 0:
        try:
            aguardar_enter_robusto("[MANUAL] Pressione ENTER para continuar...", tag="MANUAL")
        except Exception:
            sleep(5)
    else:
        print(f"[MANUAL] Aguardando {PAUSA_MANUAL_SEGUNDOS}s para ações manuais...")
        sleep(PAUSA_MANUAL_SEGUNDOS)
    print(f"[MANUAL] CONTINUANDO_EXECUCAO_MANUAL | {contexto}")


def pausa_final_livros_manual(contexto: str = ""):
    """Pausa final para baixar os livros manualmente. Encerra apenas ao pressionar ENTER."""
    if not PAUSA_MANUAL_FINAL:
        return
    msg = f"[MANUAL-FINAL] {contexto}".strip() if contexto else "[MANUAL-FINAL] Baixe os livros manualmente (Tomados/Prestados)."
    print(msg)
    try:
        append_txt(caminho_log_manual(), msg)
    except Exception:
        pass

    if aguardar_liberacao_manual_app(contexto or "FINAL", tag="MANUAL-FINAL"):
        try:
            append_txt(caminho_log_manual(), f"[MANUAL-FINAL] CONTINUANDO_EXECUCAO_MANUAL | {contexto}")
        except Exception:
            pass
        return

    try:
        aguardar_enter_robusto("[MANUAL-FINAL] Pressione ENTER para encerrar o main.py...", tag="MANUAL-FINAL")
    except Exception:
        sleep(5)
    print(f"[MANUAL-FINAL] CONTINUANDO_EXECUCAO_MANUAL | {contexto}")


def pausa_login_captcha_manual(contexto: str):
    """Registra o login/captcha manual e deixa o monitoramento seguir automaticamente."""
    tag = "MANUAL-LOGIN"
    log_manual = caminho_log_manual()
    msg = f"[{tag}] {contexto}"
    print(msg)
    try:
        append_txt(log_manual, msg)
    except Exception:
        pass

    try:
        salvar_print_evidencia(
            "LOGIN_CAPTCHA",
            "MANUAL_LOGIN_REQUIRED",
            log_path=log_manual,
        )
    except Exception:
        pass

    auto_msg = f"[{tag}] MONITORAMENTO_AUTOMATICO | Apos clicar em 'Entrar', a automacao segue sozinha."
    print(auto_msg)
    try:
        append_txt(log_manual, auto_msg)
    except Exception:
        pass


def encontrar_tr_por_referencia_grid(ref_alvo: str, col_ref: str) -> object:
    """Retorna <tr> cuja célula de Referência (col_ref) seja exatamente ref_alvo."""
    celulas_ref = driver.find_elements(By.XPATH, f"//td[contains(@id,',{col_ref}_grid')]")
    for cel in celulas_ref:
        try:
            if (cel.text or "").strip() != ref_alvo:
                continue
            tr = cel.find_element(By.XPATH, "./ancestor::tr[1]")
            # valida novamente
            cel2 = tr.find_element(By.XPATH, f".//td[contains(@id,',{col_ref}_grid')]")
            if (cel2.text or "").strip() != ref_alvo:
                continue
            return tr
        except Exception:
            continue
    return None


def normalizar_referencia_competencia(s: str) -> str:
    s = (s or "").replace("\u00a0", " ").strip()
    s = re.sub(r"\s+", "", s)
    m = re.match(r"^(\d{1,2})/(\d{4})$", s)
    if m:
        return f"{int(m.group(1)):02d}/{m.group(2)}"
    return s


def obter_referencia_topo_declaracao(timeout=0) -> str:
    def _ler_agora() -> str:
        try:
            el = driver.find_element(By.ID, "_label30")
            return (el.text or "").strip()
        except Exception:
            pass
        try:
            el = driver.find_element(By.XPATH, "//label[normalize-space()='Referência']/following::label[1]")
            return (el.text or "").strip()
        except Exception:
            return ""

    if not timeout:
        return _ler_agora()

    fim = time.time() + timeout
    while time.time() < fim:
        valor = _ler_agora()
        if valor:
            return valor
        sleep(0.20)
    return _ler_agora()


def _linha_grid_parece_ativa(tr) -> bool:
    try:
        classes = _norm_ascii(tr.get_attribute("class") or "")
        if any(token in classes for token in ("selected", "highlight", "active", "current", "focus")):
            return True
    except Exception:
        pass

    try:
        aria_selected = _norm_ascii(tr.get_attribute("aria-selected") or "")
        if aria_selected == "true":
            return True
    except Exception:
        pass

    try:
        for ctrl in tr.find_elements(By.XPATH, ".//input[@type='radio' or @type='checkbox']"):
            try:
                if ctrl.is_selected():
                    return True
            except Exception:
                continue
    except Exception:
        pass

    return False


def garantir_contexto_referencia_grid(
    ref_alvo: str,
    col_ref: str,
    timeout=15,
    log_path: str = "",
    modulo: str = "",
):
    ref_norm = normalizar_referencia_competencia(ref_alvo)
    prefixo = f"{modulo} | {ref_alvo} | " if modulo or ref_alvo else ""
    fim = time.time() + timeout
    ultimo_erro = ""

    while time.time() < fim:
        tr = encontrar_tr_por_referencia_grid(ref_alvo, col_ref)
        if tr is None:
            raise RuntimeError(f"Linha da referência {ref_alvo} não encontrada na grid")

        ref_topo = normalizar_referencia_competencia(obter_referencia_topo_declaracao(timeout=0))
        if ref_topo == ref_norm:
            if log_path:
                append_txt(log_path, f"{prefixo}CONTEXTO_REF=OK | TOPO={ref_topo}")
            return tr

        try:
            alvo_click = tr.find_element(By.XPATH, f".//td[contains(@id,',{col_ref}_grid')]")
        except Exception:
            alvo_click = tr

        for alvo in (alvo_click, tr):
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", alvo)
            except Exception:
                pass
            try:
                click_robusto(alvo)
            except Exception as e:
                ultimo_erro = f"click: {type(e).__name__}"
                continue

            sleep(0.25)
            tr_atual = encontrar_tr_por_referencia_grid(ref_alvo, col_ref) or tr
            ref_topo = normalizar_referencia_competencia(obter_referencia_topo_declaracao(timeout=2))
            if ref_topo == ref_norm:
                if log_path:
                    append_txt(log_path, f"{prefixo}CONTEXTO_REF=OK | TOPO={ref_topo}")
                return tr_atual
            if not ref_topo and _linha_grid_parece_ativa(tr_atual):
                if log_path:
                    append_txt(log_path, f"{prefixo}CONTEXTO_REF=OK | TOPO_VAZIO")
                return tr_atual

            ultimo_erro = f"topo={ref_topo or 'vazio'}"

        sleep(0.20)

    ref_final = normalizar_referencia_competencia(obter_referencia_topo_declaracao(timeout=0))
    raise RuntimeError(
        f"Nao foi possivel ativar a referência {ref_alvo} na grid. "
        f"Topo atual: {ref_final or 'vazio'}. Ultimo erro: {ultimo_erro}"
    )


def clicar_cadeado_na_linha(ref_alvo: str, col_ref: str) -> bool:
    """Clica no cadeado (fa-lock) da linha da competência alvo. Retorna True se clicou."""
    tr = encontrar_tr_por_referencia_grid(ref_alvo, col_ref)
    if not tr:
        return False

    # Título varia (Fechar Movimento / Fechar movimento). Use translate para case-insensitive.
    locks = tr.find_elements(
        By.XPATH,
        ".//span[contains(@class,'fa-lock') and contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'fechar')]"
    )
    if not locks:
        return False

    click_robusto(locks[0])

    # Alguns fluxos disparam alert/confirm
    try:
        WebDriverWait(driver, 2).until(EC.alert_is_present())
        driver.switch_to.alert.accept()
    except Exception:
        pass

    # Aguarda a tela reagir (cadeado sumir ou DOM atualizar). Não é crítico — apenas evita cliques em overlay.
    try:
        WebDriverWait(driver, 20).until(lambda d: len(
            (encontrar_tr_por_referencia_grid(ref_alvo, col_ref) or tr).find_elements(
                By.XPATH,
                ".//span[contains(@class,'fa-lock') and contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'fechar')]"
            )
        ) == 0)
    except Exception:
        pass

    return True
def inicializar_chrome():
    global driver, wait
    embedded_assets = _resolve_embedded_browser_assets()
    try:
        if embedded_assets:
            service = Service(executable_path=embedded_assets["driver"])
            driver = webdriver.Chrome(
                service=service,
                options=_build_chrome_options(embedded_assets["chrome"]),
            )
            print(
                "Chrome embutido iniciado. "
                f"Origem: {embedded_assets['source']} | Perfil: {CHROME_PROFILE_DIR}"
            )
        else:
            driver = webdriver.Chrome(options=_build_chrome_options())
            print(f"Chrome do sistema iniciado. Perfil: {CHROME_PROFILE_DIR}")
        wait = WebDriverWait(driver, 30)
        if ENABLE_CHROME_PERFORMANCE_LOGS:
            try:
                driver.execute_cdp_cmd("Network.enable", {})
            except Exception:
                pass
            try:
                driver.get_log("performance")
            except Exception:
                pass
    except Exception as e:
        if embedded_assets:
            print(f"Falha ao iniciar Chrome embutido, tentando fallback do sistema: {e}")
            try:
                driver = webdriver.Chrome(options=_build_chrome_options())
                wait = WebDriverWait(driver, 30)
                print(f"Chrome do sistema iniciado via fallback. Perfil: {CHROME_PROFILE_DIR}")
                return
            except Exception as fallback_exc:
                raise RuntimeError(f"{MSG_CHROME_INIT_FALHA}: {fallback_exc}") from fallback_exc
        raise RuntimeError(f"{MSG_CHROME_INIT_FALHA}: {e}")



# =====================
# KIT FORENSE (DEBUG)
# =====================
def _agora_ts():
    # 2026-02-25_14-03-22
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def salvar_kit_forense(motivo, exc=None):
    """Salva evidências (screenshot, HTML, URL, logs) em downloads/<empresa>/<MM.AAAA>/_GERAL/_debug/<ts>/.
    Retorna o caminho do diretório criado (ou None se não foi possível).
    """
    try:
        ts = _agora_ts()
        prefixo = _prefixo_cadastro_base()
        nome_dir = f"{prefixo}__{ts}" if prefixo else ts
        debug_dir = os.path.join(pasta_debug(), nome_dir)
        os.makedirs(debug_dir, exist_ok=True)

        # Metadados básicos
        with open(os.path.join(debug_dir, "info.txt"), "w", encoding="utf-8") as f:
            f.write(f"timestamp={ts}\n")
            f.write(f"motivo={motivo}\n")
            f.write(f"cadastro={json.dumps(contexto_cadastro_atual(), ensure_ascii=False)}\n")
            try:
                f.write(f"url={driver.current_url}\n" if driver else "url=<sem driver>\n")
                f.write(f"title={driver.title}\n" if driver else "title=<sem driver>\n")
            except Exception:
                f.write("url/title=<erro ao coletar>\n")

            if exc is not None:
                f.write(f"exception={type(exc).__name__}: {exc}\n")

        # Screenshot
        try:
            if driver is not None:
                driver.save_screenshot(os.path.join(debug_dir, "screenshot.png"))
        except Exception:
            pass

        # HTML
        try:
            if driver is not None:
                with open(os.path.join(debug_dir, "page_source.html"), "w", encoding="utf-8", errors="replace") as f:
                    f.write(driver.page_source or "")
        except Exception:
            pass

        # Console logs (se suportado)
        try:
            if driver is not None:
                logs = driver.get_log("browser")
                with open(os.path.join(debug_dir, "browser_console.json"), "w", encoding="utf-8") as f:
                    json.dump(logs, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

        # Stacktrace (se houver)
        try:
            import traceback
            with open(os.path.join(debug_dir, "traceback.txt"), "w", encoding="utf-8", errors="replace") as f:
                f.write(traceback.format_exc())
        except Exception:
            pass

        print(f"[KIT-FORENSE] Evidências salvas em: {debug_dir}")
        return debug_dir
    except Exception:
        return None


def salvar_dump_bruto_tomados_pdf(tag: str, payload: bytes, meta: dict | None = None, info: dict | None = None, ref_alvo: str = "") -> str:
    """Salva o payload bruto retornado pelo portal para diagnosticar respostas intermediárias do PDF de Tomados."""
    try:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        ano_ref, mes_ref = parse_referencia_competencia(ref_alvo)
        prefixo = _prefixo_cadastro_base()
        nome_dir = f"tomados_pdf_raw_{ts}"
        if prefixo:
            nome_dir = f"{prefixo}__{nome_dir}"
        debug_dir = os.path.join(pasta_debug(ano_alvo=ano_ref, mes_alvo=mes_ref), nome_dir)
        os.makedirs(debug_dir, exist_ok=True)

        info = enriquecer_info_com_cadastro(info)
        payload = payload or b""
        meta = dict(meta or {})
        meta["referencia"] = (ref_alvo or "").strip()
        meta["nf"] = (info.get("nf") or "").strip()
        meta["data_emissao"] = (info.get("data_emissao") or "").strip()
        meta["cadastro"] = contexto_cadastro_atual()
        meta["payload_size"] = len(payload)
        meta["payload_head_hex"] = payload[:96].hex()
        meta["payload_head_text"] = _decodificar_bytes_resposta(payload[:512])[:500]

        with open(os.path.join(debug_dir, f"{tag}_meta.json"), "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        with open(os.path.join(debug_dir, f"{tag}_body.bin"), "wb") as f:
            f.write(payload)

        try:
            with open(os.path.join(debug_dir, f"{tag}_body_preview.txt"), "w", encoding="utf-8", errors="replace") as f:
                f.write(_decodificar_bytes_resposta(payload[:4096]))
        except Exception:
            pass

        try:
            if ref_alvo:
                salvar_log_tomados(
                    "PDF_DEBUG_DUMP",
                    ref_alvo,
                    mensagem=f"TAG={tag} NF={(info.get('nf') or '').strip()} PASTA={debug_dir}",
                )
        except Exception:
            pass

        return debug_dir
    except Exception:
        return ""


def _slug_evidencia(txt: str) -> str:
    base = _norm_ascii(txt or "")
    base = re.sub(r"[^a-z0-9]+", "_", base).strip("_")
    return base or "evidencia"


def _novo_prefixo_evidencia(contexto: str, motivo: str, ano_alvo: int = 0, mes_alvo: int = 0, ref_alvo: str = "") -> str:
    ts_arquivo = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    servico = _servico_evidencia_por_contexto(contexto)
    if servico == "_GERAL":
        base_dir = pasta_geral(ano_alvo=ano_alvo, mes_alvo=mes_alvo)
    else:
        base_dir = pasta_servico(servico, ano_alvo=ano_alvo, mes_alvo=mes_alvo)
    evid_dir = os.path.join(base_dir, "_evidencias")
    os.makedirs(evid_dir, exist_ok=True)

    partes_nome = [_slug_evidencia(contexto), _slug_evidencia(ref_alvo or motivo), ts_arquivo]
    prefixo = _prefixo_cadastro_base()
    if prefixo:
        partes_nome.insert(0, prefixo)
    return os.path.join(evid_dir, "_".join([p for p in partes_nome if p]))


def salvar_print_evidencia(contexto: str, motivo: str, ano_alvo: int = 0, mes_alvo: int = 0, ref_alvo: str = "", log_path: str = "") -> str:
    """Salva screenshot completo com carimbo visível de data/hora para fins de evidência."""
    ts = datetime.now()
    ts_visivel = ts.strftime("%d/%m/%Y %H:%M:%S")
    destino = _novo_prefixo_evidencia(contexto, motivo, ano_alvo=ano_alvo, mes_alvo=mes_alvo, ref_alvo=ref_alvo) + ".png"

    overlay_id = "__codex_evidencia_timestamp__"
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    try:
        try:
            driver.execute_script(
                """
                const old = document.getElementById(arguments[0]);
                if (old) old.remove();
                const tag = document.createElement('div');
                tag.id = arguments[0];
                tag.innerText = arguments[1];
                tag.style.position = 'fixed';
                tag.style.top = '12px';
                tag.style.right = '12px';
                tag.style.zIndex = '2147483647';
                tag.style.background = 'rgba(0,0,0,0.88)';
                tag.style.color = '#fff';
                tag.style.padding = '10px 14px';
                tag.style.font = '700 18px/1.35 monospace';
                tag.style.border = '2px solid #ffd54f';
                tag.style.borderRadius = '8px';
                tag.style.boxShadow = '0 4px 18px rgba(0,0,0,.35)';
                document.body.appendChild(tag);
                """,
                overlay_id,
                f"{contexto} | {motivo} | {ts_visivel}",
            )
            sleep(0.15)
        except Exception:
            pass

        try:
            metrics = driver.execute_cdp_cmd("Page.getLayoutMetrics", {})
            content_size = metrics.get("contentSize", {})
            width = max(1280, int(math.ceil(float(content_size.get("width", 1280) or 1280))))
            height = max(720, int(math.ceil(float(content_size.get("height", 720) or 720))))
            shot = driver.execute_cdp_cmd(
                "Page.captureScreenshot",
                {
                    "format": "png",
                    "fromSurface": True,
                    "captureBeyondViewport": True,
                    "clip": {
                        "x": 0,
                        "y": 0,
                        "width": width,
                        "height": height,
                        "scale": 1,
                    },
                },
            )
            with open(destino, "wb") as f:
                f.write(base64.b64decode(shot["data"]))
        except Exception:
            driver.save_screenshot(destino)

        if log_path:
            append_txt(
                log_path,
                f"EVIDENCIA=OK | CONTEXTO={contexto} | REF={ref_alvo} | MOTIVO={motivo} | ARQ={os.path.basename(destino)} | PASTA={os.path.dirname(destino)}",
            )
        return destino
    finally:
        try:
            driver.execute_script(
                "const el = document.getElementById(arguments[0]); if (el) el.remove();",
                overlay_id,
            )
        except Exception:
            pass


def salvar_evidencia_html(contexto: str, motivo: str, ano_alvo: int = 0, mes_alvo: int = 0, ref_alvo: str = "", log_path: str = "", extra: dict | None = None) -> dict:
    """Salva screenshot e HTML da página atual/default_content para diagnosticar estados de UI frágeis."""
    resultado = {"png": "", "html": "", "meta": ""}
    try:
        resultado["png"] = salvar_print_evidencia(
            contexto,
            motivo,
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            ref_alvo=ref_alvo,
            log_path=log_path,
        )
    except Exception:
        resultado["png"] = ""

    prefixo = os.path.splitext(resultado["png"])[0] if resultado["png"] else _novo_prefixo_evidencia(
        contexto,
        motivo,
        ano_alvo=ano_alvo,
        mes_alvo=mes_alvo,
        ref_alvo=ref_alvo,
    )
    html_path = prefixo + ".html"
    meta_path = prefixo + "_meta.json"

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    try:
        with open(html_path, "w", encoding="utf-8", errors="replace") as f:
            f.write(driver.page_source or "")
        resultado["html"] = html_path
    except Exception:
        resultado["html"] = ""

    meta = {
        "contexto": contexto,
        "motivo": motivo,
        "referencia": (ref_alvo or "").strip(),
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "url": "",
        "title": "",
        "iframe_count": 0,
        "frame_count": 0,
        "extra": extra or {},
    }
    try:
        meta["url"] = driver.current_url or ""
    except Exception:
        pass
    try:
        meta["title"] = driver.title or ""
    except Exception:
        pass
    try:
        meta["iframe_count"] = len(driver.find_elements(By.TAG_NAME, "iframe"))
    except Exception:
        pass
    try:
        meta["frame_count"] = len(driver.find_elements(By.TAG_NAME, "frame"))
    except Exception:
        pass

    try:
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        resultado["meta"] = meta_path
    except Exception:
        resultado["meta"] = ""

    if log_path:
        try:
            if resultado["html"]:
                append_txt(
                    log_path,
                    f"EVIDENCIA_HTML=OK | CONTEXTO={contexto} | REF={ref_alvo} | MOTIVO={motivo} | ARQ={os.path.basename(resultado['html'])} | PASTA={os.path.dirname(resultado['html'])}",
                )
            if resultado["meta"]:
                append_txt(
                    log_path,
                    f"EVIDENCIA_META=OK | CONTEXTO={contexto} | REF={ref_alvo} | MOTIVO={motivo} | ARQ={os.path.basename(resultado['meta'])} | PASTA={os.path.dirname(resultado['meta'])}",
                )
        except Exception:
            pass

    return resultado


def limpar_input(el):
    try:
        el.clear()
    except Exception:
        pass
    try:
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(Keys.DELETE)
    except Exception:
        try:
            el.send_keys(Keys.COMMAND, "a")
            el.send_keys(Keys.DELETE)
        except Exception:
            pass
    try:
        driver.execute_script("arguments[0].value='';", el)
    except Exception:
        pass


def _norm_txt(s: str) -> str:
    s = (s or "").lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s


def captcha_incorreto_na_tela() -> bool:
    """Detecta mensagem do portal quando o captcha foi digitado errado.
    Ex.: <div class="msg-login">Texto da imagem incorreto.</div>
    """
    try:
        src = _norm_txt(driver.page_source or "")
        return "texto da imagem incorreto" in src
    except Exception:
        return False



def credencial_invalida_na_tela() -> bool:
    try:
        src = _norm_txt(driver.page_source or "")
        sinais = [
            "usuario invalido",
            "senha invalida",
            "senha invalida",  # variações sem acento já cobertas
            "senha invalida",  # manter simples/robusto
        ]
        return any(t in src for t in sinais)
    except Exception:
        return False



def sem_modulo_nota_fiscal_no_dashboard() -> bool:
    """Detecta contribuinte sem módulo de Nota Fiscal no dashboard inicial."""
    try:
        # Espera estrutura de botÃµes aparecer antes de concluir ausência.
        if not driver.find_elements(By.ID, "divbotoes"):
            return False
        tem_img = len(driver.find_elements(By.ID, "imgnotafiscal")) > 0
        tem_txt = len(driver.find_elements(By.ID, LOGIN_CARD_DASHBOARD)) > 0
        return not (tem_img or tem_txt)
    except Exception:
        return False


def preencher_login_prefeitura_se_habilitado():
    global PRESTADOS_SEM_MODULO
    if not AUTO_LOGIN_PREFEITURA:
        return "manual"

    PRESTADOS_SEM_MODULO = False
    cnpj = re.sub(r"\D", "", os.environ.get("EMPRESA_CNPJ", ""))
    senha = os.environ.get("EMPRESA_SENHA", "")
    if not cnpj or not senha:
        raise RuntimeError("AUTO_LOGIN_PREFEITURA ativo, mas EMPRESA_CNPJ/EMPRESA_SENHA não informados")

    driver.get(LOGIN_URL_PREFEITURA)
    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, LOGIN_CAMPO_USUARIO)))
    campo_senha = wait.until(EC.presence_of_element_located((By.ID, LOGIN_CAMPO_SENHA)))

    limpar_input(campo_usuario)
    campo_usuario.send_keys(cnpj)
    limpar_input(campo_senha)
    campo_senha.send_keys(senha)

    wait.until(EC.element_to_be_clickable((By.ID, LOGIN_BOTAO_ENTRAR)))

    contexto_login = "LOGIN_CAPTCHA | Resolva o captcha e clique em 'Entrar'. A automacao seguira sozinha."
    print("CNPJ e senha preenchidos. Login manual formal aguardando captcha.")
    pausa_login_captcha_manual(contexto_login)
    print(f"Aguardando dashboard por até {login_wait_seconds}s após a liberação manual...")

    fim = time.time() + login_wait_seconds
    status_login = ""
    while time.time() < fim:
        if credencial_invalida_na_tela():
            try:
                salvar_print_evidencia("LOGIN_CAPTCHA", "CREDENCIAL_INVALIDA")
            except Exception:
                pass
            raise RuntimeError(MSG_CREDENCIAL_INVALIDA)
        if captcha_incorreto_na_tela():
            try:
                salvar_print_evidencia("LOGIN_CAPTCHA", "CAPTCHA_INCORRETO")
            except Exception:
                pass
            raise RuntimeError(MSG_CAPTCHA_INCORRETO)
        if tela_cadastros_relacionados_aberta():
            status_login = "cadastros_relacionados"
            break
        if sem_modulo_nota_fiscal_no_dashboard():
            PRESTADOS_SEM_MODULO = True
            print("Módulo Nota Fiscal não disponível no dashboard. Prestados será ignorado e o fluxo seguirá para Tomados.")
            status_login = "dashboard_sem_modulo"
            break
        if driver.find_elements(By.ID, LOGIN_CARD_DASHBOARD):
            status_login = "dashboard"
            break
        sleep(0.4)
    else:
        try:
            salvar_print_evidencia("LOGIN_CAPTCHA", "CAPTCHA_NAO_RESOLVIDO_NO_TEMPO")
        except Exception:
            pass
        raise RuntimeError(MSG_CAPTCHA_TIMEOUT)

    return status_login or "dashboard"


def aguardar_navegacao_manual_inicial():
    if AUTO_LOGIN_PREFEITURA:
        return
    print("Entre manualmente em: Nota Fiscal - Lista Nota Fiscais")
    print(f"Você tem {login_wait_seconds} segundos para iniciar manualmente.")
    sleep(login_wait_seconds)


def tela_cadastros_relacionados_aberta() -> bool:
    try:
        if driver is None:
            return False
        if not driver.find_elements(By.ID, "_imagebutton1"):
            return False
        if driver.find_elements(By.NAME, "_grid1Selected"):
            return True
        src = _norm_txt(driver.page_source or "")
        return "cadastros relacionados" in src and "selecao do contribuinte" in src
    except Exception:
        return False


def tela_declaracao_fiscal_eletronica_aberta() -> bool:
    try:
        if driver.find_elements(By.ID, "imgprestiss") or driver.find_elements(By.ID, "divtxtprestiss"):
            return True
        if driver.find_elements(By.ID, "imgtomadiss") or driver.find_elements(By.ID, "divtxttomadiss"):
            return True
        src = _norm_txt(driver.page_source or "")
        return "declaracao fiscal eletronica" in src
    except Exception:
        return False


def _linhas_cadastros_relacionados_page():
    try:
        return driver.find_elements(By.XPATH, "//tr[@tipo='tr' and .//input[@name='_grid1Selected']]")
    except Exception:
        return []


def _extrair_row_key_cadastro(tr) -> str:
    try:
        tds = tr.find_elements(By.XPATH, ".//td[contains(@id,'__grid1')]")
    except Exception:
        tds = []
    for td in tds:
        try:
            td_id = (td.get_attribute("id") or "").strip()
        except Exception:
            continue
        m = re.match(r"(\d+),-?\d+__grid1$", td_id)
        if m:
            return m.group(1)
    return ""


def _extrair_cadastro_relacionado_da_linha(tr, page: int, ordem: int, row_index: int) -> dict:
    radio = tr.find_element(By.XPATH, ".//input[@name='_grid1Selected']")
    try:
        identificacao_el = tr.find_element(By.XPATH, ".//td[contains(@id,',2__grid1')]")
        identificacao = re.sub(r"\s+", " ", (identificacao_el.text or "").strip())
    except Exception:
        identificacao = re.sub(r"\s+", " ", (tr.text or "").strip())

    ccm = ""
    m = re.search(r"CCM:\s*([\d.\-\/]+)", identificacao or "", flags=re.IGNORECASE)
    if m:
        ccm = re.sub(r"\D", "", m.group(1))

    return {
        "ordem": ordem,
        "page": page,
        "row_index": row_index,
        "row_key": _extrair_row_key_cadastro(tr),
        "valor": str(radio.get_attribute("value") or "").strip(),
        "ccm": ccm,
        "identificacao": identificacao,
    }


def assinatura_cadastros_relacionados() -> str:
    pagina = pagina_atual_cadastros_relacionados()
    linhas = _linhas_cadastros_relacionados_page()
    tokens = []
    for idx, tr in enumerate(linhas[:12], start=1):
        cadastro = _extrair_cadastro_relacionado_da_linha(tr, pagina or 0, idx, idx)
        tokens.append(
            "::".join(
                [
                    str(cadastro.get("row_key") or ""),
                    str(cadastro.get("valor") or ""),
                    str(cadastro.get("identificacao") or ""),
                ]
            )
        )
    return f"{pagina or 0}|{len(linhas)}|{'|'.join(tokens)}"


def pagina_atual_cadastros_relacionados():
    try:
        val = driver.execute_script("return (document.getElementById('_grid1Page')||{}).value || '';")
        if str(val).strip().isdigit():
            return int(str(val).strip())
    except Exception:
        pass
    try:
        ativos = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'dataTables_paginate')]//li[contains(@class,'paginate_button') and contains(@class,'active')]//span",
        )
        for el in ativos:
            txt = (getattr(el, "text", "") or "").strip()
            if txt.isdigit():
                return int(txt)
    except Exception:
        pass
    return None


def _extrair_pagina_alvo_paginacao_generica(onclick: str, grid_name: str) -> int | None:
    m = re.search(rf"mudarPagina,{re.escape(grid_name)},(\d+)", onclick or "")
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _clicar_botao_paginacao_cadastros(botao, target_page: int | None, assinatura_anterior: str, pagina_anterior: int | None) -> bool:
    click_robusto(botao)
    try:
        WebDriverWait(driver, WAIT_LISTA_TIMEOUT, poll_frequency=GRID_POLL_INTERVAL_SECONDS).until(
            lambda _d: (
                assinatura_cadastros_relacionados() != assinatura_anterior
                or (
                    pagina_atual_cadastros_relacionados() is not None
                    and target_page is not None
                    and pagina_atual_cadastros_relacionados() == target_page
                )
                or (
                    pagina_anterior is not None
                    and pagina_atual_cadastros_relacionados() is not None
                    and pagina_atual_cadastros_relacionados() > pagina_anterior
                )
            )
        )
        return True
    except TimeoutException:
        pagina_depois = pagina_atual_cadastros_relacionados()
        return bool(
            pagina_anterior is not None
            and pagina_depois is not None
            and pagina_depois > pagina_anterior
        )


def ir_para_proxima_pagina_cadastros_relacionados() -> bool:
    assinatura_anterior = assinatura_cadastros_relacionados()
    pagina_anterior = pagina_atual_cadastros_relacionados()

    try:
        botoes_next = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'dataTables_paginate')]"
            "//li[contains(@class,'next') and not(contains(@class,'disabled'))]"
            "//span[contains(@onclick,'mudarPagina,_grid1,')]",
        )
    except Exception:
        botoes_next = []

    for btn in botoes_next:
        alvo = _extrair_pagina_alvo_paginacao_generica(btn.get_attribute("onclick") or "", "_grid1")
        if alvo is None:
            continue
        print(f"Cadastros: avançando de {pagina_anterior} para {alvo}.")
        return _clicar_botao_paginacao_cadastros(btn, alvo, assinatura_anterior, pagina_anterior)

    return False


def _botao_pagina_cadastros_relacionados(pagina_alvo: int):
    try:
        botoes = driver.find_elements(
            By.XPATH,
            f"//div[contains(@class,'dataTables_paginate')]//span[contains(@onclick,'mudarPagina,_grid1,{pagina_alvo}')]",
        )
    except Exception:
        botoes = []
    return botoes[0] if botoes else None


def ir_para_pagina_cadastros_relacionados(pagina_alvo: int) -> None:
    pagina_alvo = int(pagina_alvo or 1)
    while True:
        pagina_atual = pagina_atual_cadastros_relacionados() or 1
        if pagina_atual == pagina_alvo:
            return
        if pagina_atual > pagina_alvo:
            botao = _botao_pagina_cadastros_relacionados(pagina_alvo)
            if botao is None:
                raise RuntimeError(f"Cadastros Relacionados: não foi possível voltar para a página {pagina_alvo}")
            assinatura = assinatura_cadastros_relacionados()
            if _clicar_botao_paginacao_cadastros(botao, pagina_alvo, assinatura, pagina_atual):
                continue
            raise RuntimeError(f"Cadastros Relacionados: clique para a página {pagina_alvo} não mudou a grid")

        botao = _botao_pagina_cadastros_relacionados(pagina_alvo)
        if botao is not None:
            assinatura = assinatura_cadastros_relacionados()
            if _clicar_botao_paginacao_cadastros(botao, pagina_alvo, assinatura, pagina_atual):
                continue
            raise RuntimeError(f"Cadastros Relacionados: clique para a página {pagina_alvo} não mudou a grid")

        if not ir_para_proxima_pagina_cadastros_relacionados():
            raise RuntimeError(f"Cadastros Relacionados: não foi possível alcançar a página {pagina_alvo}")


def coletar_cadastros_relacionados() -> list[dict]:
    if not tela_cadastros_relacionados_aberta():
        return []

    cadastros = []
    ordem = 1
    pagina = pagina_atual_cadastros_relacionados() or 1
    while True:
        linhas = _linhas_cadastros_relacionados_page()
        if not linhas:
            break
        for idx, tr in enumerate(linhas, start=1):
            cadastros.append(_extrair_cadastro_relacionado_da_linha(tr, pagina, ordem, idx))
            ordem += 1
        if not ir_para_proxima_pagina_cadastros_relacionados():
            break
        pagina = pagina_atual_cadastros_relacionados() or (pagina + 1)
    return cadastros


def _normalizar_identificacao_cadastro_relacionado(txt: str) -> str:
    token = _norm_txt(re.sub(r"\s+", " ", str(txt or "").strip()))
    token = re.sub(r"[^a-z0-9]+", " ", token)
    return re.sub(r"\s+", " ", token).strip()


def _cadastro_relacionado_tem_contexto(cadastro: dict | None) -> bool:
    if not isinstance(cadastro, dict):
        return False
    try:
        row_index = int(cadastro.get("row_index") or 0)
    except Exception:
        row_index = 0
    return row_index > 0


def _mesclar_cadastro_relacionado(cadastro_base: dict, cadastro_encontrado: dict) -> dict:
    base = dict(cadastro_base or {})
    encontrado = dict(cadastro_encontrado or {})
    merged = dict(base)

    for campo in ("page", "row_index", "row_key", "valor", "ccm", "identificacao"):
        valor = encontrado.get(campo)
        if valor not in (None, ""):
            merged[campo] = valor

    try:
        ordem_base = int(base.get("ordem") or 0)
    except Exception:
        ordem_base = 0
    try:
        ordem_encontrada = int(encontrado.get("ordem") or 0)
    except Exception:
        ordem_encontrada = 0
    merged["ordem"] = ordem_base or ordem_encontrada

    for campo in ("page", "row_index"):
        try:
            merged[campo] = int(merged.get(campo) or 0)
        except Exception:
            merged[campo] = 0

    merged["row_key"] = str(merged.get("row_key") or "").strip()
    merged["valor"] = str(merged.get("valor") or "").strip()
    merged["ccm"] = re.sub(r"\D", "", str(merged.get("ccm") or ""))
    merged["identificacao"] = re.sub(r"\s+", " ", str(merged.get("identificacao") or "").strip())
    return merged


def _conteudo_linha_cadastro_relacionado_confere(
    cadastro_esperado: dict,
    cadastro_encontrado: dict,
    *,
    validar_row_index: bool = False,
) -> bool:
    if not cadastro_esperado or not cadastro_encontrado:
        return False

    try:
        idx_esperado = int(cadastro_esperado.get("row_index") or 0)
    except Exception:
        idx_esperado = 0
    try:
        idx_encontrado = int(cadastro_encontrado.get("row_index") or 0)
    except Exception:
        idx_encontrado = 0

    try:
        page_esperada = int(cadastro_esperado.get("page") or 0)
    except Exception:
        page_esperada = 0
    try:
        page_encontrada = int(cadastro_encontrado.get("page") or 0)
    except Exception:
        page_encontrada = 0

    if idx_esperado <= 0 or idx_encontrado <= 0:
        return False
    if idx_encontrado != idx_esperado:
        return False
    if page_esperada > 0 and page_encontrada > 0 and page_encontrada != page_esperada:
        return False
    if validar_row_index:
        return True
    return True


def _aguardar_grade_cadastros_relacionados_carregada(timeout: int = 40, minimo_linhas: int = 1):
    minimo_linhas = max(1, int(minimo_linhas or 1))

    def _grid_pronta(_d):
        if not tela_cadastros_relacionados_aberta():
            return False
        linhas = _linhas_cadastros_relacionados_page()
        if len(linhas) < minimo_linhas:
            return False
        return linhas

    try:
        return WebDriverWait(
            driver,
            max(5, int(timeout or 40)),
            poll_frequency=GRID_POLL_INTERVAL_SECONDS,
        ).until(_grid_pronta)
    except Exception:
        linhas = _linhas_cadastros_relacionados_page()
        if len(linhas) >= minimo_linhas:
            return linhas
        raise RuntimeError(
            f"Cadastros Relacionados: a grade não carregou com linhas suficientes (necessário row_index={minimo_linhas})."
        )


def _sincronizar_contexto_cadastro_relacionado(cadastro: dict) -> dict:
    global MULTI_CADASTRO_TARDIO_ESTADO

    atualizado = ativar_contexto_cadastro(cadastro)
    if not multicadastro_tardio_ativo():
        return atualizado

    cadastros = [dict(item) for item in (MULTI_CADASTRO_TARDIO_ESTADO.get("cadastros") or [])]
    if not cadastros:
        return atualizado

    try:
        idx = int(MULTI_CADASTRO_TARDIO_ESTADO.get("indice_atual") or -1)
    except Exception:
        idx = -1
    if 0 <= idx < len(cadastros):
        merged = dict(cadastros[idx])
        merged.update({k: v for k, v in atualizado.items() if v not in (None, "") or k in {"ordem", "page", "row_index"}})
        cadastros[idx] = merged
        MULTI_CADASTRO_TARDIO_ESTADO["cadastros"] = cadastros
    return atualizado


def _localizar_linha_cadastro_relacionado(cadastro: dict, timeout: int = 40, recoletas: int = 2):
    pagina_alvo = max(1, int(cadastro.get("page") or 1))
    try:
        row_index_alvo = int(cadastro.get("row_index") or 0)
    except Exception:
        row_index_alvo = 0
    if row_index_alvo <= 0:
        raise RuntimeError(
            f"Cadastros Relacionados: cadastro sem row_index válido para ordem={cadastro.get('ordem')}"
        )
    ultimo_erro = (
        f"Cadastros Relacionados: página {pagina_alvo} sem a linha ordinal {row_index_alvo} "
        f"para ordem={cadastro.get('ordem')}"
    )

    for tentativa in range(1, max(1, recoletas) + 1):
        try:
            ir_para_pagina_cadastros_relacionados(pagina_alvo)
            linhas = _aguardar_grade_cadastros_relacionados_carregada(
                timeout=min(timeout, 20),
                minimo_linhas=row_index_alvo,
            )
        except Exception as exc:
            ultimo_erro = f"Falha ao carregar a grade da página {pagina_alvo}: {type(exc).__name__}: {str(exc)[:180]}"
            continue

        pagina_atual = pagina_atual_cadastros_relacionados() or pagina_alvo
        if pagina_atual != pagina_alvo:
            ultimo_erro = (
                f"Cadastros Relacionados: página atual divergente ao retomar a linha ordinal "
                f"{row_index_alvo}. Esperado={pagina_alvo} Atual={pagina_atual}"
            )
            continue

        if len(linhas) < row_index_alvo:
            ultimo_erro = (
                f"Cadastros Relacionados: página {pagina_atual} possui {len(linhas)} linha(s), "
                f"mas era esperado alcançar a linha ordinal {row_index_alvo}."
            )
            continue

        alvo = linhas[row_index_alvo - 1]
        ordem = int(cadastro.get("ordem") or row_index_alvo)
        candidato = _mesclar_cadastro_relacionado(
            cadastro,
            _extrair_cadastro_relacionado_da_linha(alvo, pagina_atual, ordem, row_index_alvo),
        )
        if _conteudo_linha_cadastro_relacionado_confere(cadastro, candidato, validar_row_index=True):
            return alvo, candidato
        ultimo_erro = (
            f"Cadastros Relacionados: a linha ordinal {row_index_alvo} da página {pagina_atual} "
            "não pôde ser confirmada."
        )

    raise RuntimeError(ultimo_erro)


def _linha_cadastro_relacionado_esta_selecionada(cadastro: dict) -> bool:
    if not tela_cadastros_relacionados_aberta():
        return False

    try:
        row_index_alvo = int(cadastro.get("row_index") or 0)
    except Exception:
        row_index_alvo = 0
    if row_index_alvo <= 0:
        return False

    try:
        pagina_esperada = int(cadastro.get("page") or 0)
    except Exception:
        pagina_esperada = 0
    pagina_atual = pagina_atual_cadastros_relacionados() or pagina_esperada or 1
    if pagina_esperada > 0 and pagina_atual != pagina_esperada:
        return False

    linhas = _linhas_cadastros_relacionados_page()
    if len(linhas) < row_index_alvo:
        return False

    tr = linhas[row_index_alvo - 1]
    ordem = int(cadastro.get("ordem") or row_index_alvo)
    candidato = _mesclar_cadastro_relacionado(
        cadastro,
        _extrair_cadastro_relacionado_da_linha(tr, pagina_atual, ordem, row_index_alvo),
    )
    if not _conteudo_linha_cadastro_relacionado_confere(cadastro, candidato, validar_row_index=True):
        return False

    try:
        classes = (tr.get_attribute("class") or "").lower()
    except Exception:
        classes = ""
    if "selected" in classes:
        return True

    try:
        radio = tr.find_element(By.XPATH, ".//input[@name='_grid1Selected']")
        checked = radio.is_selected() or str(radio.get_attribute("checked") or "").strip().lower() in {"true", "checked"}
        if checked:
            return True
    except Exception:
        pass
    return False


def selecionar_cadastro_relacionado(cadastro: dict, timeout: int = 40) -> dict:
    if not tela_cadastros_relacionados_aberta():
        raise RuntimeError("Cadastros Relacionados não está aberto para seleção")
    if not _cadastro_relacionado_tem_contexto(cadastro):
        raise RuntimeError("Cadastros Relacionados: cadastro sem contexto suficiente para seleção.")

    timeout_local = max(5, int(timeout or 40))
    alvo, cadastro_resolvido = _localizar_linha_cadastro_relacionado(cadastro, timeout=timeout_local, recoletas=2)

    clicavel = None
    try:
        clicavel = alvo.find_element(By.XPATH, ".//td[contains(@id,',2__grid1')]")
    except Exception:
        pass
    if clicavel is None:
        try:
            celulas = alvo.find_elements(By.XPATH, ".//td[not(contains(@style,'display:none'))]")
            clicavel = celulas[0] if celulas else alvo
        except Exception:
            clicavel = alvo

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", clicavel)
    except Exception:
        pass
    click_robusto(clicavel)

    WebDriverWait(driver, timeout).until(
        lambda _d: _linha_cadastro_relacionado_esta_selecionada(cadastro_resolvido)
    )

    cadastro_resolvido = _sincronizar_contexto_cadastro_relacionado(cadastro_resolvido)
    btn_continuar = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "_imagebutton1")))
    click_robusto(btn_continuar)

    WebDriverWait(driver, timeout).until(
        lambda _d: (
            not tela_cadastros_relacionados_aberta()
            and (
                tela_declaracao_fiscal_eletronica_aberta()
                or len(_d.find_elements(By.ID, LOGIN_CARD_DASHBOARD)) > 0
                or sem_modulo_nota_fiscal_no_dashboard()
            )
        )
    )
    return dict(cadastro_resolvido)


def retornar_para_cadastros_relacionados(timeout: int = 40) -> None:
    if tela_cadastros_relacionados_aberta():
        return

    def _esperar():
        WebDriverWait(driver, timeout).until(lambda _d: tela_cadastros_relacionados_aberta())

    def _achar_breadcrumb_cadastros(d):
        candidatos = d.find_elements(By.CSS_SELECTOR, "a.historic-item")
        for el in candidatos:
            try:
                txt = _norm_txt((el.text or "").strip())
                title = _norm_txt((el.get_attribute("title") or "").strip())
                if "cadastros relacion" in txt or "cadastros relacion" in title:
                    return el
            except Exception:
                continue
        return False

    try:
        link = WebDriverWait(driver, min(timeout, 8)).until(_achar_breadcrumb_cadastros)
        click_robusto(link)
        _esperar()
        return
    except Exception:
        pass

    botoes_voltar = []
    try:
        botoes_voltar.extend(driver.find_elements(By.XPATH, "//button[contains(normalize-space(.),'Voltar')]"))
        botoes_voltar.extend(driver.find_elements(By.XPATH, "//a[contains(normalize-space(.),'Voltar')]"))
    except Exception:
        pass
    for btn in botoes_voltar:
        try:
            click_robusto(btn)
            _esperar()
            return
        except Exception:
            continue

    driver.get(LOGIN_URL_PREFEITURA)
    _esperar()


def garantir_lista_nota_fiscal_contexto(timeout: int = 40) -> bool:
    global PRESTADOS_SEM_MODULO

    try:
        status_lista = esperar_lista_ou_sem_checkbox(timeout=2)
        if status_lista in {"checkboxes", "sem_checkbox", "sem_checkbox_com_data"}:
            return True
    except Exception:
        pass

    try:
        clicar_inicio_para_dashboard(timeout=timeout)
    except Exception:
        pass

    if sem_modulo_nota_fiscal_no_dashboard():
        PRESTADOS_SEM_MODULO = True
        print("Módulo Nota Fiscal não disponível para o cadastro atual. Prestados será ignorado.")
        return False

    navegar_para_lista_nota_fiscal()
    return True

# =====================
# HELPERS â€“ CLIQUE / LISTA / 502
# =====================
def click_robusto(el):
    try:
        el.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", el)


def fechar_alerta_mensagens_prefeitura_se_aberto(timeout=5, log_path: str = "", modulo: str = "", ref_alvo: str = "") -> bool:
    """Fecha o popup informativo da prefeitura que pode abrir ao entrar em Tomados/Prestados."""
    marcadores = (
        "alerta de mensagens",
        "declaracao sem pagamento",
        "efetue o recolhimento",
        "cobranca judicial",
        "parcele seus debitos",
    )
    botoes_xpaths = [
        "//button[contains(normalize-space(.),'Fechar')]",
        "//a[contains(normalize-space(.),'Fechar')]",
        "//button[contains(@class,'ui-dialog-titlebar-close') or contains(@class,'close') or @title='Fechar']",
        "//a[contains(@class,'ui-dialog-titlebar-close') or @title='Fechar']",
        "//*[contains(@class,'ui-icon-closethick')]/ancestor::button[1]",
        "//*[contains(@class,'ui-icon-closethick')]/ancestor::a[1]",
    ]

    def _contexto_tem_alerta() -> bool:
        try:
            src = _norm_txt(driver.page_source or "")
        except Exception:
            return False
        return any(marker in src for marker in marcadores)

    def _buscar_botoes_fechar():
        encontrados = []
        for xp in botoes_xpaths:
            try:
                encontrados.extend(driver.find_elements(By.XPATH, xp))
            except Exception:
                continue
        return encontrados

    def _iterar_contextos():
        yield ("default", None, None)
        try:
            outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            outer_frames = []

        for idx_outer, outer in enumerate(outer_frames):
            try:
                fid = (outer.get_attribute("id") or "").strip()
                fname = (outer.get_attribute("name") or "").strip()
                if not outer.is_displayed():
                    continue
            except Exception:
                continue

            if "_iFilho" not in fid and "_iFilho" not in fname:
                continue

            yield ("iframe", idx_outer, None)
            try:
                driver.switch_to.default_content()
                outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
                if idx_outer >= len(outer_frames):
                    continue
                driver.switch_to.frame(outer_frames[idx_outer])
                inner_frames = driver.find_elements(By.TAG_NAME, "frame")
            except Exception:
                inner_frames = []
            finally:
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass

            for idx_inner in range(len(inner_frames)):
                yield ("frame", idx_outer, idx_inner)

    fim = time.time() + timeout
    fechou = False

    while time.time() < fim:
        encontrou = False
        for tipo_ctx, idx_outer, idx_inner in _iterar_contextos():
            try:
                driver.switch_to.default_content()
                if tipo_ctx == "iframe":
                    outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
                    if idx_outer is None or idx_outer >= len(outer_frames):
                        continue
                    driver.switch_to.frame(outer_frames[idx_outer])
                elif tipo_ctx == "frame":
                    outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
                    if idx_outer is None or idx_outer >= len(outer_frames):
                        continue
                    driver.switch_to.frame(outer_frames[idx_outer])
                    inner_frames = driver.find_elements(By.TAG_NAME, "frame")
                    if idx_inner is None or idx_inner >= len(inner_frames):
                        continue
                    driver.switch_to.frame(inner_frames[idx_inner])

                if not _contexto_tem_alerta():
                    continue

                encontrou = True
                botoes = _buscar_botoes_fechar()
                botao = None
                for item in botoes:
                    try:
                        if item.is_displayed():
                            botao = item
                            break
                    except Exception:
                        continue
                if botao is None and botoes:
                    botao = botoes[0]
                if botao is None:
                    continue

                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", botao)
                except Exception:
                    pass
                click_robusto(botao)
                fechou = True
                break
            except Exception:
                continue
            finally:
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass

        try:
            driver.switch_to.default_content()
        except Exception:
            pass

        if fechou:
            try:
                WebDriverWait(driver, 4).until(lambda d: not _contexto_tem_alerta())
            except Exception:
                pass
            if log_path:
                prefixo = f"{modulo} | {ref_alvo} | " if modulo or ref_alvo else ""
                append_txt(log_path, f"{prefixo}ALERTA_MENSAGENS=FECHADO")
            return True

        if not encontrou:
            return False

        sleep(0.2)

    return fechou


def navegar_para_lista_nota_fiscal():
    print("Login confirmado. Navegando automaticamente: Nota Fiscal -> Lista Nota Fiscais...")

    # 1) Dashboard inicial: clicar em Nota Fiscal.
    card_nf = wait.until(EC.element_to_be_clickable((By.ID, LOGIN_CARD_DASHBOARD)))
    click_robusto(card_nf)

    # 2) Segundo dashboard: clicar no card Lista Nota Fiscais (id confirmado pelo escritório).
    try:
        card_lista = WebDriverWait(driver, 12).until(EC.element_to_be_clickable((By.ID, LOGIN_CARD_LISTA_NOTAS)))
        click_robusto(card_lista)
        status_lista = esperar_lista_ou_sem_checkbox(timeout=WAIT_LISTA_TIMEOUT)
        if status_lista.startswith("sem_checkbox"):
            print("Tela de lista carregada sem checkbox; empresa será concluída sem competência.")
        else:
            print("Tela de lista de notas carregada automaticamente.")
        return
    except Exception as e:
        ultimo_erro = e

    # Fallback por texto para maior resiliência se o id variar.
    seletores_lista = [
        (By.XPATH, "//a[contains(normalize-space(.), 'Lista Nota Fiscais')]"),
        (By.XPATH, "//span[contains(normalize-space(.), 'Lista Nota Fiscais')]"),
        (By.XPATH, "//*[contains(normalize-space(.), 'Lista Nota Fiscais')]"),
    ]

    for by, sel in seletores_lista:
        try:
            el = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((by, sel)))
            click_robusto(el)
            status_lista = esperar_lista_ou_sem_checkbox(timeout=WAIT_LISTA_TIMEOUT)
            if status_lista.startswith("sem_checkbox"):
                print("Tela de lista carregada sem checkbox (fallback); empresa será concluída sem competência.")
            else:
                print("Tela de lista de notas carregada automaticamente (fallback por texto).")
            return
        except Exception as e:
            ultimo_erro = e

    raise RuntimeError(f"NAO_FOI_POSSIVEL_ABRIR_LISTA_NOTAS: {ultimo_erro}")

def esperar_lista(timeout=WAIT_LISTA_TIMEOUT):
    WebDriverWait(driver, timeout, poll_frequency=GRID_POLL_INTERVAL_SECONDS).until(
        EC.presence_of_all_elements_located((By.NAME, "gridListaCheck"))
    )

def obter_checkboxes_lista(timeout=0):
    checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
    if checkboxes or timeout <= 0:
        return checkboxes

    try:
        status_lista = esperar_lista_ou_sem_checkbox(timeout=timeout)
    except TimeoutException:
        return []

    if status_lista != "checkboxes":
        return []

    return driver.find_elements(By.NAME, "gridListaCheck")

def esperar_lista_ou_sem_checkbox(timeout=WAIT_LISTA_TIMEOUT):
    """
    Aguarda a tela de lista ficar pronta e classifica:
      - "checkboxes": há notas selecionáveis
      - "sem_checkbox_com_data": lista sem checkbox, mas com data de emissão visível
      - "sem_checkbox": lista carregada sem checkbox (fallback por page size)
    """
    def cond(_):
        checks = driver.find_elements(By.NAME, "gridListaCheck")
        if len(checks) > 0:
            return "checkboxes"

        # Caso real observado: há linha e radio selecionado, mas sem checkbox de marcação.
        radios_linha = driver.find_elements(By.NAME, "gridListaSelected")
        celulas_linha = driver.find_elements(By.CSS_SELECTOR, "td[id*=',-1_gridLista']")

        # Se já existe coluna de data da grid, a lista carregou mesmo sem checkboxes.
        datas = driver.find_elements(By.CSS_SELECTOR, "td[id*=',10_gridLista']")
        if len(datas) > 0 and len(checks) == 0:
            return "sem_checkbox_com_data"

        if (len(radios_linha) > 0 or len(celulas_linha) > 0) and len(checks) == 0:
            return "sem_checkbox"

        # Fallback: page size presente também indica lista carregada.
        has_pagesize = len(driver.find_elements(By.ID, "gridListaPageSize")) > 0
        if has_pagesize and len(checks) == 0:
            return "sem_checkbox"

        return False

    return WebDriverWait(driver, timeout, poll_frequency=GRID_POLL_INTERVAL_SECONDS).until(cond)

def lista_ativa():
    try:
        return len(obter_checkboxes_lista(timeout=0)) > 0
    except Exception:
        return False

def assinatura_lista():
    """Assinatura leve da grid para detectar refresh/paginação."""
    try:
        checks = obter_checkboxes_lista(timeout=0)
        values = [c.get_attribute("value") or "" for c in checks[:5]]
        return f"{len(checks)}|{'|'.join(values)}"
    except Exception:
        return ""

def esperar_troca_de_grid(assinatura_anterior, timeout=WAIT_LISTA_TIMEOUT):
    def mudou(_):
        return assinatura_lista() != assinatura_anterior
    WebDriverWait(driver, timeout, poll_frequency=GRID_POLL_INTERVAL_SECONDS).until(mudou)

def is_502_page():
    """
    Detecção ESTRITA de 502 (evita falso positivo).
    Só retorna True se tiver sinais claros de "Bad Gateway" e a lista NÃO estiver ativa.
    """
    try:
        if lista_ativa():
            return False

        title = (driver.title or "").lower()
        src = (driver.page_source or "").lower()

        if ("bad gateway" in title and "502" in title):
            return True

        # padrÃµes comuns de página 502
        if "502 bad gateway" in src:
            return True
        if "http error 502" in src:
            return True
        if ("bad gateway" in src and "502" in src):
            return True

        return False
    except Exception:
        return False

def recover_from_502():
    print("Detectado 502 REAL. Fazendo refresh e aguardando lista voltar...")
    try:
        driver.refresh()
    except Exception:
        try:
            driver.get(driver.current_url)
        except Exception:
            pass

    esperar_lista(timeout=WAIT_LISTA_TIMEOUT)
    print("Lista recuperada apos 502.")

# =====================
# MODAL 'AVISO' (1 NOTA POR VEZ)
# =====================
def fechar_aviso_se_existir(timeout=2):
    try:
        modal = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((
                By.XPATH,
                "//*[contains(.,'ATENÃ‡ÃƒO!') and contains(.,'apenas a exportação de uma nota')]"
                "/ancestor::*[contains(@class,'modal')][1]"
            ))
        )
        ok = modal.find_element(By.XPATH, ".//button[contains(.,'OK') or contains(.,'Ok')]")
        click_robusto(ok)
        sleep(0.2)
        return True
    except Exception:
        return False

def desmarcar_todas_notas():
    checks = driver.find_elements(By.NAME, "gridListaCheck")
    for c in checks:
        try:
            if c.is_selected():
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", c)
                click_robusto(c)
                sleep(0.02)
        except Exception:
            pass


def marcar_somente_checkbox(checkbox, full_reset=False):
    global ULTIMO_CHECKBOX_MARCADO

    if full_reset:
        desmarcar_todas_notas()
        ULTIMO_CHECKBOX_MARCADO = None
    elif ULTIMO_CHECKBOX_MARCADO is not None:
        try:
            if ULTIMO_CHECKBOX_MARCADO != checkbox and ULTIMO_CHECKBOX_MARCADO.is_selected():
                click_robusto(ULTIMO_CHECKBOX_MARCADO)
        except Exception:
            ULTIMO_CHECKBOX_MARCADO = None

    try:
        if not checkbox.is_selected():
            click_robusto(checkbox)
        ULTIMO_CHECKBOX_MARCADO = checkbox
        return
    except Exception:
        pass

    if not full_reset:
        try:
            desmarcar_todas_notas()
            if not checkbox.is_selected():
                click_robusto(checkbox)
            ULTIMO_CHECKBOX_MARCADO = checkbox
        except Exception:
            ULTIMO_CHECKBOX_MARCADO = None

# =====================
# EXTRAIR INFO DA LINHA (NF/DATA/RPS/CHAVE)
# =====================
def td_text(tr, col):
    el = tr.find_element(By.CSS_SELECTOR, f"td[id*=',{col}_gridLista']")
    return (el.text or "").strip()


def _texto_para_numero_limpo(txt: str) -> str:
    return re.sub(r"\D", "", txt or "")


def _nf_valido(nf: str) -> bool:
    n = _texto_para_numero_limpo(nf)
    return len(n) > 0


def _inferir_nf_por_proximidade(tr, col_data: int, col_nf: int, col_rps: int, nf_atual: str, rps_atual: str, data_emissao_atual: str):
    """Fallback para recuperar NF quando coluna vem deslocada e NF fica vazia/inválida."""
    if _nf_valido(nf_atual):
        return nf_atual

    rps_limpo = _texto_para_numero_limpo(rps_atual)
    data_limpa = _texto_para_numero_limpo(data_emissao_atual)
    data_yyyymmdd = ""
    parsed_data = parse_data_emissao_site(data_emissao_atual or "")
    if parsed_data:
        yyyy, mm, dd, _, _ = parsed_data
        data_yyyymmdd = f"{yyyy}{mm}{dd}"

    candidatos = []

    # Prioriza colunas explicitamente de NF antes de heurística por proximidade.
    for c in (col_nf, col_nf + 1, 6, 7):
        if c >= 0:
            candidatos.append(c)

    # Em layouts conhecidos, NF costuma estar ~4 colunas antes de Data Emissão.
    for delta in (-4, -5, -3, -6, -2, 0, 1, -1):
        c = col_data + delta
        if c < 0:
            continue
        candidatos.append(c)

    vistos = set()
    candidatos_validos = []
    for c in candidatos:
        if c in vistos:
            continue
        vistos.add(c)
        try:
            txt = td_text(tr, c)
        except Exception:
            continue

        num = _texto_para_numero_limpo(txt)
        if not num:
            continue
        if rps_limpo and num == rps_limpo:
            continue
        if data_limpa and num == data_limpa:
            continue
        if data_yyyymmdd and num == data_yyyymmdd:
            continue
        if len(num) > 12:
            continue
        candidatos_validos.append(num)

    for num in candidatos_validos:
        if len(num) >= 3:
            return num
    if candidatos_validos:
        return candidatos_validos[0]

    return nf_atual


def _normalizar_titulo_coluna(txt: str) -> str:
    txt = (txt or "").replace("\xa0", " ").strip()
    txt = txt.replace("Âº", "o")
    txt = re.sub(r"\s+", " ", txt)
    return _norm_ascii(txt)


def mapa_colunas_grid_lista(force=False):
    """
    Mapeia colunas pelo THEAD usando títulos normalizados para evitar leitura deslocada
    (empresa/tema diferente). O `columnorder` do TH começa em 0 após a coluna de checkbox,
    então no TD o índice final é `columnorder - 1`.
    """
    global GRID_HEADER_CACHE

    now = time.time()
    if not force and GRID_HEADER_CACHE["map"] and (now - GRID_HEADER_CACHE["ts"]) < 10:
        return GRID_HEADER_CACHE["map"]

    col_map = {}
    headers = driver.find_elements(By.CSS_SELECTOR, "#_gridListaTHeadLinhas th[columnorder]")
    for th in headers:
        try:
            order_raw = th.get_attribute("columnorder")
            if order_raw is None:
                continue
            order = int(order_raw)
            idx_td = order - 1
            if idx_td < 0:
                continue

            titulo = _normalizar_titulo_coluna(th.text)
            if not titulo:
                continue

            if "data emissao" in titulo and "rps" not in titulo:
                col_map["data_emissao"] = idx_td
            elif "situacao" in titulo:
                col_map["situacao"] = idx_td
            elif "chave de validacao/acesso" in titulo or "chave de validacao" in titulo:
                col_map["chave"] = idx_td
            elif "rps" in titulo and "data" not in titulo and "serie" not in titulo:
                col_map["rps"] = idx_td
            elif titulo == "nf":
                col_map["nf"] = idx_td
        except Exception:
            continue

    GRID_HEADER_CACHE = {"ts": now, "map": col_map}
    return col_map


def primeira_data_emissao_visivel_sem_checkbox():
    """Lê a primeira data da coluna Data Emissão quando não há checkbox."""
    col_data = mapa_colunas_grid_lista().get("data_emissao", 10)
    try:
        el = driver.find_element(By.CSS_SELECTOR, f"td[id*=',{col_data}_gridLista']")
        return (el.text or "").strip()
    except Exception:
        return ""

def extrair_info_linha(checkbox):
    tr = checkbox.find_element(By.XPATH, "./ancestor::tr[1]")

    mapa = mapa_colunas_grid_lista()
    col_nf = mapa.get("nf", 6)
    col_situacao = mapa.get("situacao", 9)
    col_data = mapa.get("data_emissao", 10)
    col_rps = mapa.get("rps", 11)
    col_chave = mapa.get("chave", 14)

    info = {
        "id_interno": checkbox.get_attribute("value") or "",
        "nf": td_text(tr, col_nf),
        "situacao": td_text(tr, col_situacao),
        "data_emissao": td_text(tr, col_data),  # dd/mm/aaaa
        "rps": td_text(tr, col_rps),
        "chave": td_text(tr, col_chave),
    }

    # Fallback para grids que chegam deslocadas em +1 coluna.
    if not parse_data_emissao_site(info.get("data_emissao", "")):
        alt = {
            "nf": td_text(tr, col_nf + 1),
            "situacao": td_text(tr, col_situacao + 1),
            "data_emissao": td_text(tr, col_data + 1),
            "rps": td_text(tr, col_rps + 1),
            "chave": td_text(tr, col_chave + 1),
        }
        if parse_data_emissao_site(alt.get("data_emissao", "")):
            info.update(alt)

    info["nf"] = _inferir_nf_por_proximidade(
        tr,
        col_data=col_data,
        col_nf=col_nf,
        col_rps=col_rps,
        nf_atual=info.get("nf", ""),
        rps_atual=info.get("rps", ""),
        data_emissao_atual=info.get("data_emissao", ""),
    )

    return enriquecer_info_com_cadastro(info)


def extrair_info_linha_tomados(checkbox):
    """Extrai dados da grid de Serviços Tomados com layout próprio."""
    tr = checkbox.find_element(By.XPATH, "./ancestor::tr[1]")

    info = {
        "id_interno": checkbox.get_attribute("value") or "",
        "nf": td_text(tr, 4),
        "situacao": td_text(tr, 5),
        "cnpj": td_text(tr, 6),
        "nome": td_text(tr, 7),
        "data_emissao": td_text(tr, 8),
        "cfps": td_text(tr, 9),
        "atividade": td_text(tr, 10),
        "chave": "",
    }

    info["chave"] = "|".join(
        [
            (info.get("id_interno") or "").strip(),
            (info.get("nf") or "").strip(),
            (info.get("data_emissao") or "").strip(),
        ]
    )
    return enriquecer_info_com_cadastro(info)

def nota_cancelada(info: dict) -> bool:
    situacao = (info.get("situacao") or "").strip().lower()
    return "cancelad" in situacao

def chave_unica(info: dict) -> str:
    info = enriquecer_info_com_cadastro(info)
    ch = (info.get("chave") or "").strip()
    if ch:
        base = ch
    else:
        iid = (info.get("id_interno") or "").strip()
        if iid:
            base = f"ID:{iid}"
        else:
            nf = (info.get("nf") or "").strip()
            rps = (info.get("rps") or "").strip()
            dt = (info.get("data_emissao") or "").strip()
            base = f"NF:{nf}_RPS:{rps}_DT:{dt}"
    prefixo = _prefixo_cadastro_base(
        {
            "ordem": info.get("cadastro_ordem"),
            "ccm": info.get("cadastro_ccm"),
            "identificacao": info.get("cadastro_identificacao"),
        }
    )
    base_norm = _normalizar_token_chave(base)
    return f"{prefixo}__{base_norm}" if prefixo else base_norm

# =====================
# DATA (SITE) -> ISO + COMPETENCIA
# =====================
def parse_data_emissao_site(data_emissao: str):
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", data_emissao or "")
    if not m:
        return None
    dd, mm, yyyy = m.group(1), m.group(2), m.group(3)
    iso = f"{yyyy}-{mm}-{dd}"
    competencia = f"{mm}.{yyyy}"
    return yyyy, mm, dd, iso, competencia


def competencia_nota(info: dict) -> tuple[int, int] | None:
    parsed = parse_data_emissao_site(info.get("data_emissao", ""))
    if not parsed:
        return None
    return int(parsed[0]), int(parsed[1])

def calcular_mes_alvo(apuracao_ref: str = ""):
    """
    Retorna (ano_alvo, mes_alvo) para download.
    Regra: sempre mês anterior ao mês de apuração.
    apuracao_ref esperado: MM/AAAA.
    """
    if apuracao_ref:
        m = re.match(r"^(\d{2})/(\d{4})$", apuracao_ref)
        if not m:
            raise ValueError(f"APURACAO_REFERENCIA invalida: {apuracao_ref}. Use MM/AAAA")
        mes_ap, ano_ap = int(m.group(1)), int(m.group(2))
    else:
        hoje = datetime.now()
        mes_ap, ano_ap = hoje.month, hoje.year

    if not (1 <= mes_ap <= 12):
        raise ValueError(f"Mes de apuracao invalido: {mes_ap}")

    if mes_ap == 1:
        return ano_ap - 1, 12
    return ano_ap, mes_ap - 1

def comparar_competencia_nota(info: dict, ano_alvo: int, mes_alvo: int):
    """
    Retorna:
      1  -> nota mais nova que o mês alvo
      0  -> nota no mês alvo
     -1  -> nota mais antiga que o mês alvo
      None -> data inválida/ausente
    """
    comp_nota = competencia_nota(info)
    if comp_nota is None:
        return None
    comp_alvo = (ano_alvo, mes_alvo)

    if comp_nota == comp_alvo:
        return 0
    if comp_nota > comp_alvo:
        return 1
    return -1


def comparar_competencia_nota_intervalo(
    info: dict,
    ano_inicio: int,
    mes_inicio: int,
    ano_fim: int,
    mes_fim: int,
):
    comp_nota = competencia_nota(info)
    if comp_nota is None:
        return None
    if comp_nota > (ano_fim, mes_fim):
        return 1
    if comp_nota < (ano_inicio, mes_inicio):
        return -1
    return 0

def nota_no_mes_alvo(info: dict, ano_alvo: int, mes_alvo: int) -> bool:
    return comparar_competencia_nota(info, ano_alvo, mes_alvo) == 0

def extrair_dhproc_do_xml(caminho_xml):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        dh = root.find(".//{*}dhProc")
        if dh is not None and dh.text:
            return dh.text.strip()
    except Exception:
        pass
    return ""

# =====================
# DOWNLOAD â€“ esperar XML NOVO (na pasta _tmp)
# =====================
def aguardar_xml_novo(pasta, timeout=80, xmls_antes=None):
    inicio_perf = time.perf_counter()
    fim = time.time() + timeout
    xmls_antes = set(xmls_antes or [])

    while time.time() < fim:
        arquivos = list(os.scandir(pasta))

        if any(e.name.lower().endswith(".crdownload") for e in arquivos):
            sleep(DOWNLOAD_POLL_INTERVAL_SECONDS)
            continue

        novo_mais_recente = None
        novo_mais_recente_mtime = -1.0
        for entry in arquivos:
            nome = entry.name
            if not nome.lower().endswith(".xml") or nome in xmls_antes:
                continue
            try:
                mtime = entry.stat().st_mtime
            except Exception:
                continue
            if mtime > novo_mais_recente_mtime:
                novo_mais_recente = nome
                novo_mais_recente_mtime = mtime

        if novo_mais_recente:
            perf_log("download.xml_novo", inicio_perf, f"arquivo={novo_mais_recente}")
            return os.path.join(pasta, novo_mais_recente)

        sleep(DOWNLOAD_POLL_INTERVAL_SECONDS)

    perf_log("download.xml_novo_timeout", inicio_perf)
    raise TimeoutError("Download nao finalizou (xml novo nao apareceu)")


def _nome_arquivo_por_content_disposition(content_disposition: str, fallback: str = "arquivo.xml") -> str:
    cd = (content_disposition or "").strip()
    nome = ""
    if cd:
        try:
            m = re.search(r"filename\*\s*=\s*UTF-8''([^;]+)", cd, flags=re.IGNORECASE)
            if m:
                nome = urllib_parse.unquote(m.group(1).strip().strip('"').strip("'"))
            if not nome:
                m = re.search(r'filename\s*=\s*"([^"]+)"', cd, flags=re.IGNORECASE)
                if m:
                    nome = m.group(1).strip()
            if not nome:
                m = re.search(r"filename\s*=\s*([^;]+)", cd, flags=re.IGNORECASE)
                if m:
                    nome = m.group(1).strip().strip('"').strip("'")
        except Exception:
            nome = ""
    nome = nome or fallback
    nome = os.path.basename(nome).strip() or fallback
    nome = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nome)
    return nome


def _header_ci(headers: dict | None, nome: str) -> str:
    alvo = (nome or "").strip().lower()
    for k, v in (headers or {}).items():
        if str(k).strip().lower() == alvo:
            return str(v or "").strip()
    return ""


def _salvar_bytes_tmp(bytes_arquivo: bytes, nome_arquivo: str) -> str:
    nome = (nome_arquivo or "arquivo.bin").strip() or "arquivo.bin"
    destino = os.path.join(TEMP_DOWNLOAD_DIR, nome)
    if os.path.exists(destino):
        base, ext = os.path.splitext(nome)
        k = 1
        while True:
            cand = os.path.join(TEMP_DOWNLOAD_DIR, f"{base}_{k}{ext}")
            if not os.path.exists(cand):
                destino = cand
                break
            k += 1
    with open(destino, "wb") as f:
        f.write(bytes_arquivo or b"")
    return destino


def _garantir_nome_pdf(nome_arquivo: str, fallback: str = "arquivo.pdf") -> str:
    nome = os.path.basename((nome_arquivo or "").strip()) or fallback
    if not os.path.splitext(nome)[1]:
        nome = f"{nome}.pdf"
    elif not nome.lower().endswith(".pdf"):
        nome = f"{os.path.splitext(nome)[0]}.pdf"
    return nome


def _nome_pdf_por_url(pdf_url: str, fallback: str = "arquivo.pdf") -> str:
    path = os.path.basename(urllib_parse.urlparse((pdf_url or "").strip()).path or "").strip()
    return _garantir_nome_pdf(path or fallback, fallback=fallback)


def _salvar_pdf_tmp_validado(
    payload: bytes,
    nome_fallback: str = "arquivo.pdf",
    content_disposition: str = "",
    pdf_url: str = "",
) -> str:
    data = payload or b""
    if not data or len(data) < 20:
        raise RuntimeError("Resposta de PDF vazia ou muito pequena")

    if not data.lstrip().startswith(b"%PDF"):
        raise RuntimeError("Resposta nao parece PDF valido")

    nome_base = _nome_pdf_por_url(pdf_url, fallback=nome_fallback)
    nome = _nome_arquivo_por_content_disposition(content_disposition, fallback=nome_base)
    nome = _garantir_nome_pdf(nome, fallback=nome_base)
    return _salvar_bytes_tmp(data, nome)


def _limpar_logs_performance():
    if not ENABLE_CHROME_PERFORMANCE_LOGS:
        return
    try:
        driver.get_log("performance")
    except Exception:
        pass


def _decodificar_bytes_resposta(payload: bytes) -> str:
    bruto = payload or b""
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            return bruto.decode(enc)
        except Exception:
            continue
    return bruto.decode("utf-8", errors="ignore")


def _extrair_urls_pdf_de_texto(texto: str, base_url: str = "") -> list[str]:
    bruto = (texto or "").strip()
    if not bruto:
        return []

    variacoes = [
        bruto,
        html.unescape(bruto),
        bruto.replace("\\/", "/"),
        html.unescape(bruto.replace("\\/", "/")),
    ]
    vistas = set()
    urls = []
    patterns = [
        r"https?://[^\"'()<>\s]+?\.pdf(?:\?[^\"'()<>\s]*)?",
        r"/resultados/[^\"'()<>\s]+?\.pdf(?:\?[^\"'()<>\s]*)?",
    ]

    for texto_var in variacoes:
        for pat in patterns:
            for m in re.finditer(pat, texto_var, flags=re.IGNORECASE):
                raw = (m.group(0) or "").strip()
                norm = _normalizar_url_candidata(raw, base_url or _base_url_portal())
                if not norm or norm in vistas:
                    continue
                vistas.add(norm)
                urls.append(norm)

    urls.sort(key=lambda u: 0 if "/resultados/" in u.lower() else 1)
    return urls


def _obter_body_cdp(request_id: str) -> bytes:
    if not request_id:
        return b""
    try:
        result = driver.execute_cdp_cmd("Network.getResponseBody", {"requestId": request_id}) or {}
        body = result.get("body") or ""
        if not body:
            return b""
        if result.get("base64Encoded"):
            return base64.b64decode(body)
        return body.encode("latin-1", errors="ignore")
    except Exception:
        return b""


def _is_pdf_response(status: int, content_type: str, url: str) -> bool:
    ct = (content_type or "").lower()
    u = (url or "").lower()
    return int(status or 0) == 200 and ("application/pdf" in ct or u.endswith(".pdf") or "/resultados/" in u)


class _NoRedirectHandler(urllib_request.HTTPRedirectHandler):
    def http_error_301(self, req, fp, code, msg, headers):
        return fp

    def http_error_302(self, req, fp, code, msg, headers):
        return fp

    def http_error_303(self, req, fp, code, msg, headers):
        return fp

    def http_error_307(self, req, fp, code, msg, headers):
        return fp

    def http_error_308(self, req, fp, code, msg, headers):
        return fp


def baixar_arquivo_via_submit_form(
    button,
    timeout=120,
    nome_fallback="arquivo.xml",
    tipo_esperado="xml",
    debug_dump_tag: str = "",
    debug_info: dict | None = None,
    debug_ref_alvo: str = "",
) -> str:
    """
    Executa o submit relacionado ao botão dentro da sessão autenticada e
    baixa o payload via fetch (sem usar o gerenciador de downloads do Chrome).
    """
    timeout_ms = max(5000, int(timeout * 1000))
    script = r"""
const btn = arguments[0];
const timeoutMs = arguments[1];
const settleMs = arguments[2];
const done = arguments[arguments.length - 1];

(async () => {
  function b64FromBytes(bytes) {
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }

  try {
    if (!btn) {
      done({ ok: false, error: "botao_nao_informado" });
      return;
    }

    const form = btn.form || document.getElementById("form") || document.forms[0];
    if (!form) {
      done({ ok: false, error: "form_nao_encontrado" });
      return;
    }

    const originalSubmit = form.submit;
    try {
      form.submit = function () {};
    } catch (e) {}

    try {
      if (typeof btn.onclick === "function") {
        btn.onclick();
      } else {
        const oc = btn.getAttribute("onclick");
        if (oc) {
          (new Function("event", oc)).call(btn, null);
        }
      }
    } catch (e) {}

    await new Promise((r) => setTimeout(r, Math.max(0, settleMs || 0)));

    const action = form.getAttribute("action") || "/servlet/controle";
    const method = (form.getAttribute("method") || "POST").toUpperCase();
    const body = new URLSearchParams(new FormData(form)).toString();

    try {
      form.submit = originalSubmit;
    } catch (e) {}

    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), timeoutMs);

    const resp = await fetch(action, {
      method,
      headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
      body,
      credentials: "same-origin",
      signal: ctrl.signal
    });

    clearTimeout(t);
    const bytes = new Uint8Array(await resp.arrayBuffer());
    done({
      ok: resp.ok,
      status: resp.status,
      url: resp.url || "",
      contentType: resp.headers.get("content-type") || "",
      contentDisposition: resp.headers.get("content-disposition") || "",
      dataB64: b64FromBytes(bytes)
    });
  } catch (err) {
    done({ ok: false, error: String((err && err.message) || err) });
  }
})();
"""

    result = driver.execute_async_script(script, button, timeout_ms, SUBMIT_FORM_SETTLE_MS)
    if not isinstance(result, dict):
        raise RuntimeError("Resposta inválida ao baixar arquivo via submit")
    if not result.get("ok"):
        raise RuntimeError(f"Falha no submit/fetch: {result.get('error') or result}")

    data_b64 = result.get("dataB64") or ""
    if not data_b64:
        raise RuntimeError("Payload vazio no download via submit")

    try:
        payload = base64.b64decode(data_b64)
    except Exception as e:
        raise RuntimeError(f"Falha ao decodificar payload Base64: {e}")

    if not payload:
        raise RuntimeError("Payload decodificado vazio")

    content_type = (result.get("contentType") or "").lower()
    nome = _nome_arquivo_por_content_disposition(result.get("contentDisposition") or "", fallback=nome_fallback)

    tipo = (tipo_esperado or "xml").strip().lower()
    head = payload[:200].lstrip()

    if tipo == "pdf":
        if "pdf" in content_type and not nome.lower().endswith(".pdf"):
            nome = f"{os.path.splitext(nome)[0]}.pdf"
        elif not os.path.splitext(nome)[1]:
            nome = f"{nome}.pdf"

        if not head.startswith(b"%PDF"):
            dump_dir = ""
            if debug_dump_tag:
                dump_dir = salvar_dump_bruto_tomados_pdf(
                    debug_dump_tag,
                    payload,
                    meta={
                        "origem": "submit_form",
                        "status": result.get("status"),
                        "url": result.get("url") or "",
                        "content_type": result.get("contentType") or "",
                        "content_disposition": result.get("contentDisposition") or "",
                    },
                    info=debug_info,
                    ref_alvo=debug_ref_alvo,
                )
            raise RuntimeError(
                "Resposta do submit nao parece PDF valido "
                f"(content-type={content_type or 'n/a'} url={result.get('url') or 'n/a'}"
                f"{f' dump={dump_dir}' if dump_dir else ''})"
            )
    else:
        if "xml" in content_type and not nome.lower().endswith(".xml"):
            nome = f"{os.path.splitext(nome)[0]}.xml"
        elif not os.path.splitext(nome)[1]:
            nome = f"{nome}.xml"

        if not (
            head.startswith(b"<?xml")
            or head.startswith(b"<")
            or b"<nfe" in payload[:2048].lower()
            or b"<nfse" in payload[:2048].lower()
        ):
            raise RuntimeError(
                "Resposta do submit nao parece XML valido "
                f"(content-type={content_type or 'n/a'} url={result.get('url') or 'n/a'})"
            )

    return _salvar_bytes_tmp(payload, nome)


def montar_submit_form(button) -> dict:
    """Monta action/method/body do submit do botão sem depender do download do Chrome."""
    script = r"""
const btn = arguments[0];
try {
  if (!btn) {
    return { ok: false, error: "botao_nao_informado" };
  }
  const form = btn.form || document.getElementById("form") || document.forms[0];
  if (!form) {
    return { ok: false, error: "form_nao_encontrado" };
  }

  const originalSubmit = form.submit;
  try { form.submit = function () {}; } catch (e) {}

  try {
    if (typeof btn.onclick === "function") {
      btn.onclick();
    } else {
      const oc = btn.getAttribute("onclick");
      if (oc) {
        (new Function("event", oc)).call(btn, null);
      }
    }
  } catch (e) {}

  const action = new URL(form.getAttribute("action") || "/servlet/controle", window.location.href).href;
  const method = (form.getAttribute("method") || "POST").toUpperCase();
  const body = new URLSearchParams(new FormData(form)).toString();
  const currentUrl = window.location.href || "";

  try { form.submit = originalSubmit; } catch (e) {}

  return { ok: true, action, method, body, currentUrl };
} catch (err) {
  return { ok: false, error: String((err && err.message) || err) };
}
"""
    result = driver.execute_script(script, button)
    if not isinstance(result, dict) or not result.get("ok"):
        raise RuntimeError(f"Falha ao montar submit do botao: {result}")
    return result


def _cookies_header_from_driver() -> str:
    return "; ".join(
        f"{c.get('name','')}={c.get('value','')}"
        for c in (driver.get_cookies() or [])
        if c.get("name")
    )


def _abrir_request_sem_redirect(url: str, method="GET", data: bytes | None = None, headers: dict | None = None, timeout=120):
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    opener = urllib_request.build_opener(
        urllib_request.HTTPHandler(),
        urllib_request.HTTPSHandler(context=ssl_ctx),
        _NoRedirectHandler(),
    )
    req = urllib_request.Request(url, data=data, headers=headers or {}, method=method)
    try:
        return opener.open(req, timeout=timeout)
    except urllib_error.HTTPError as e:
        return e


def baixar_pdf_tomados_via_post_controle(button, info: dict, timeout=120, ref_alvo: str = "") -> str:
    """Reproduz o POST /servlet/controle capturado no DevTools e usa o Location para baixar o PDF."""
    submit = montar_submit_form(button)
    action = (submit.get("action") or "").strip()
    body = (submit.get("body") or "").strip()
    current_url = (submit.get("currentUrl") or "").strip()
    if not action or not body:
        raise RuntimeError("Submit do PDF de Tomados nao retornou action/body")

    cookie_header = _cookies_header_from_driver()
    user_agent = ""
    try:
        user_agent = driver.execute_script("return navigator.userAgent;") or ""
    except Exception:
        user_agent = ""

    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
        "Cache-Control": "no-cache",
        "Content-Type": "application/x-www-form-urlencoded",
        "Origin": _base_url_portal().rstrip("/"),
        "Pragma": "no-cache",
        "Referer": current_url or action,
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": user_agent or "Mozilla/5.0",
    }
    if cookie_header:
        headers["Cookie"] = cookie_header

    resp = _abrir_request_sem_redirect(
        action,
        method="POST",
        data=body.encode("utf-8"),
        headers=headers,
        timeout=timeout,
    )
    status = int(getattr(resp, "status", 0) or getattr(resp, "code", 0) or 0)
    location = resp.headers.get("Location", "") if hasattr(resp, "headers") else ""
    content_type = (resp.headers.get("Content-Type", "") if hasattr(resp, "headers") else "").strip()
    payload = b""
    try:
        payload = resp.read() or b""
    except Exception:
        payload = b""
    if hasattr(resp, "close"):
        try:
            resp.close()
        except Exception:
            pass

    if status not in {200, 301, 302, 303, 307, 308}:
        raise RuntimeError(f"POST controle retornou status inesperado: {status}")

    pdf_url = _normalizar_url_candidata(location, action)
    if not pdf_url and status == 200:
        texto = _decodificar_bytes_resposta(payload)
        urls = _extrair_urls_pdf_de_texto(texto, action)
        if urls:
            pdf_url = urls[0]
        elif payload.startswith(b"%PDF"):
            nome = f"{re.sub(r'\\D', '', info.get('nf') or '') or 'nfse_tomados'}.pdf"
            return _salvar_bytes_tmp(payload, nome)

    if not pdf_url:
        dump_dir = salvar_dump_bruto_tomados_pdf(
            "post_controle_invalido",
            payload,
            meta={
                "origem": "post_controle",
                "action": action,
                "current_url": current_url,
                "status": status,
                "location": location,
                "content_type": content_type,
                "request_headers": headers,
            },
            info=info,
            ref_alvo=ref_alvo,
        )
        raise RuntimeError(
            "POST controle nao retornou URL utilizavel para o PDF "
            f"(status={status} content-type={content_type or 'n/a'}"
            f"{f' dump={dump_dir}' if dump_dir else ''})"
        )

    return baixar_pdf_por_url_com_cookies(pdf_url, timeout=timeout)


def capturar_fluxo_pdf_via_logs(timeout=20) -> dict:
    """Captura o fluxo real do PDF após o clique: POST /controle -> 302 + Location -> GET /resultados/*.pdf."""
    if not ENABLE_CHROME_PERFORMANCE_LOGS:
        return {
            "controle_302": None,
            "pdf_response": None,
        }

    inicio_perf = time.perf_counter()
    fim = time.time() + timeout
    fluxo = {
        "controle_302": None,
        "pdf_response": None,
    }
    controle_seen_at = None
    pdf_seen_at = None

    while time.time() < fim:
        try:
            logs = driver.get_log("performance")
        except Exception:
            logs = []

        for entry in logs:
            try:
                msg = json.loads(entry.get("message") or "{}").get("message") or {}
            except Exception:
                continue

            method = (msg.get("method") or "").strip()
            params = msg.get("params") or {}

            if method == "Network.requestWillBeSent":
                request = params.get("request") or {}
                request_url = _normalizar_url_candidata(request.get("url") or "", _base_url_portal())

                redirect = params.get("redirectResponse") or {}
                if redirect:
                    redirect_url = _normalizar_url_candidata(redirect.get("url") or "", _base_url_portal())
                    redirect_status = int(redirect.get("status") or 0)
                    redirect_headers = redirect.get("headers") or {}
                    location = _normalizar_url_candidata(_header_ci(redirect_headers, "Location"), redirect_url or _base_url_portal())
                    if redirect_url.lower().endswith("/servlet/controle") and redirect_status == 302 and location:
                        fluxo["controle_302"] = {
                            "status": redirect_status,
                            "url": redirect_url,
                            "location": location,
                            "headers": redirect_headers,
                        }
                        if not controle_seen_at:
                            controle_seen_at = time.time()

            elif method == "Network.responseReceived":
                request_id = (params.get("requestId") or "").strip()
                response = params.get("response") or {}
                headers = response.get("headers") or {}
                url = _normalizar_url_candidata(response.get("url") or "", _base_url_portal())
                status = int(response.get("status") or 0)
                mime = str(response.get("mimeType") or _header_ci(headers, "Content-Type") or "").strip()

                if url.lower().endswith("/servlet/controle") and status == 302:
                    location = _normalizar_url_candidata(_header_ci(headers, "Location"), url or _base_url_portal())
                    if location:
                        fluxo["controle_302"] = {
                            "status": status,
                            "url": url,
                            "location": location,
                            "headers": headers,
                        }
                        if not controle_seen_at:
                            controle_seen_at = time.time()

                if _is_pdf_response(status, mime, url):
                    fluxo["pdf_response"] = {
                        "requestId": request_id,
                        "url": url,
                        "status": status,
                        "contentType": mime,
                        "contentDisposition": _header_ci(headers, "Content-Disposition"),
                        "headers": headers,
                    }
                    if not pdf_seen_at:
                        pdf_seen_at = time.time()

            elif method == "Network.loadingFinished":
                request_id = (params.get("requestId") or "").strip()
                cand = fluxo.get("pdf_response") or {}
                if request_id and request_id == cand.get("requestId"):
                    body = _obter_body_cdp(request_id)
                    if body:
                        cand["body"] = body
                        fluxo["pdf_response"] = cand

        controle = fluxo.get("controle_302") or {}
        pdf_response = fluxo.get("pdf_response") or {}
        if controle.get("location") and (pdf_response.get("body") or pdf_response.get("url")):
            perf_log("captura_fluxo_pdf_logs", inicio_perf, "resultado=completo")
            return fluxo
        if controle.get("location") and controle_seen_at and (time.time() - controle_seen_at) >= 1.0:
            perf_log("captura_fluxo_pdf_logs", inicio_perf, "resultado=controle_302")
            return fluxo
        if pdf_response.get("url") and pdf_seen_at and (time.time() - pdf_seen_at) >= 1.0:
            perf_log("captura_fluxo_pdf_logs", inicio_perf, "resultado=pdf_response")
            return fluxo
        sleep(NETWORK_LOG_POLL_INTERVAL_SECONDS)

    perf_log("captura_fluxo_pdf_logs_timeout", inicio_perf)
    return fluxo


def capturar_fluxo_pdf_tomados_via_logs(timeout=20) -> dict:
    return capturar_fluxo_pdf_via_logs(timeout=timeout)


def materializar_pdf_de_fluxo_rede(fluxo: dict, nome_fallback: str = "arquivo.pdf", timeout=120) -> tuple[str, str]:
    controle = fluxo.get("controle_302") or {}
    pdf_response = fluxo.get("pdf_response") or {}

    controle_url = _normalizar_url_candidata(controle.get("url") or "", _base_url_portal())
    controle_location = _normalizar_url_candidata(controle.get("location") or "", controle_url or _base_url_portal())
    response_url = _normalizar_url_candidata(pdf_response.get("url") or "", _base_url_portal())
    payload = pdf_response.get("body") or b""

    if payload:
        try:
            return (
                _salvar_pdf_tmp_validado(
                    payload,
                    nome_fallback=nome_fallback,
                    content_disposition=pdf_response.get("contentDisposition") or "",
                    pdf_url=response_url or controle_location,
                ),
                "logs_rede_body",
            )
        except Exception:
            texto = _decodificar_bytes_resposta(payload)
            urls = _extrair_urls_pdf_de_texto(texto, response_url or controle_location or _base_url_portal())
            url_atual_valida = bool(
                response_url
                and (
                    response_url.lower().endswith(".pdf")
                    or "/resultados/" in response_url.lower()
                )
            )
            if urls and not controle_location and not url_atual_valida:
                response_url = urls[0]

    pdf_url = controle_location or response_url
    if not pdf_url:
        return "", ""

    referer = controle_url or _base_url_portal()
    return baixar_pdf_por_url_com_cookies(pdf_url, timeout=timeout, referer=referer), "logs_rede_url"


def map_runtime_error_to_exit_code(msg: str):
    texto = str(msg or "")
    mapping = (
        (MSG_CAPTCHA_TIMEOUT, EXIT_CODE_CAPTCHA_TIMEOUT),
        (MSG_CAPTCHA_INCORRETO, EXIT_CODE_CAPTCHA_TIMEOUT),
        (MSG_SEM_COMPETENCIA, EXIT_CODE_SEM_COMPETENCIA),
        (MSG_SEM_SERVICOS, EXIT_CODE_SEM_SERVICOS),
        (MSG_CREDENCIAL_INVALIDA, EXIT_CODE_CREDENCIAL_INVALIDA),
        (MSG_MULTI_CADASTRO_REVISAO_MANUAL, EXIT_CODE_MULTI_CADASTRO_REVISAO_MANUAL),
        (MSG_EMPRESA_MULTIPLA, EXIT_CODE_EMPRESA_MULTIPLA),
        (MSG_TOMADOS_PDF_REVISAO_MANUAL, EXIT_CODE_TOMADOS_PDF_REVISAO_MANUAL),
        (MSG_APURACAO_COMPLETA_REVISAO_MANUAL, EXIT_CODE_APURACAO_COMPLETA_REVISAO_MANUAL),
        (MSG_TOMADOS_FALHA, EXIT_CODE_TOMADOS_FALHA),
        (MSG_CHROME_INIT_FALHA, EXIT_CODE_CHROME_INIT_FALHA),
    )
    for marker, code in mapping:
        if marker in texto:
            return code
    return None


def aguardar_pdf_novo(pasta, timeout=120, pdfs_antes=None):
    inicio_perf = time.perf_counter()
    fim = time.time() + timeout
    pdfs_antes = set(pdfs_antes or [])
    estado_cr = {}

    while time.time() < fim:
        arquivos = list(os.scandir(pasta))

        novo_mais_recente = None
        novo_mais_recente_mtime = -1.0
        for entry in arquivos:
            nome = entry.name
            if not nome.lower().endswith(".pdf") or nome in pdfs_antes:
                continue
            try:
                mtime = entry.stat().st_mtime
            except Exception:
                continue
            if mtime > novo_mais_recente_mtime:
                novo_mais_recente = nome
                novo_mais_recente_mtime = mtime

        if novo_mais_recente:
            perf_log("download.pdf_novo", inicio_perf, f"arquivo={novo_mais_recente}")
            return os.path.join(pasta, novo_mais_recente)

        # Alguns ambientes bloqueiam o download e deixam apenas ".crdownload" com conteúdo PDF válido.
        # Quando o tamanho estabiliza por alguns ciclos e o cabeçalho é %PDF, copiamos para .pdf.
        agora = time.time()
        for entry in arquivos:
            nome = entry.name
            if not nome.lower().endswith(".crdownload"):
                continue
            try:
                st = entry.stat()
                tam = int(st.st_size)
            except Exception:
                continue

            if tam <= 0:
                continue

            prev = estado_cr.get(nome)
            if not prev:
                estado_cr[nome] = {"size": tam, "stable": 0}
                continue

            if prev.get("size") != tam:
                estado_cr[nome] = {"size": tam, "stable": 0}
                continue

            stable = int(prev.get("stable", 0)) + 1
            estado_cr[nome] = {"size": tam, "stable": stable}
            if tam < 1024 or stable < 3:
                continue

            origem = os.path.join(pasta, nome)
            try:
                with open(origem, "rb") as f:
                    head = f.read(5)
                if not head.startswith(b"%PDF"):
                    continue
            except Exception:
                continue

            base = re.sub(r"\.crdownload$", "", nome, flags=re.IGNORECASE).strip() or f"livro_{int(agora)}"
            if not base.lower().endswith(".pdf"):
                base = f"{base}.pdf"
            destino = os.path.join(pasta, base)
            if os.path.exists(destino):
                b, ext = os.path.splitext(base)
                k = 1
                while True:
                    cand = os.path.join(pasta, f"{b}_{k}{ext}")
                    if not os.path.exists(cand):
                        destino = cand
                        break
                    k += 1
            try:
                shutil.copy2(origem, destino)
                perf_log("download.pdf_crdownload_convertido", inicio_perf, f"arquivo={os.path.basename(destino)}")
                return destino
            except Exception:
                continue

        sleep(DOWNLOAD_POLL_INTERVAL_SECONDS)

    perf_log("download.pdf_novo_timeout", inicio_perf)
    raise TimeoutError("Download nao finalizou (pdf novo nao apareceu)")


def limpar_tmp_livros_pdf():
    """Remove resíduos de PDFs de livro no _tmp para evitar ambiguidade de detecção."""
    try:
        for nome in os.listdir(TEMP_DOWNLOAD_DIR):
            low = (nome or "").lower()
            if low.endswith(".pdf") or low.endswith(".crdownload"):
                try:
                    os.remove(os.path.join(TEMP_DOWNLOAD_DIR, nome))
                except Exception:
                    pass
    except Exception:
        pass


def fechar_janelas_extras(handles_base):
    base = list(handles_base or [])
    base_set = set(base)

    try:
        atuais = list(driver.window_handles)
    except Exception:
        return

    for h in atuais:
        if h in base_set:
            continue
        try:
            driver.switch_to.window(h)
            driver.close()
        except Exception:
            pass

    for h in base:
        try:
            driver.switch_to.window(h)
            return
        except Exception:
            continue

    try:
        if driver.window_handles:
            driver.switch_to.window(driver.window_handles[0])
    except Exception:
        pass


def _base_url_portal() -> str:
    try:
        p = urllib_parse.urlparse(LOGIN_URL_PREFEITURA)
        if p.scheme and p.netloc:
            return f"{p.scheme}://{p.netloc}/"
    except Exception:
        pass
    try:
        p = urllib_parse.urlparse(driver.current_url or "")
        if p.scheme and p.netloc:
            return f"{p.scheme}://{p.netloc}/"
    except Exception:
        pass
    return ""


def _normalizar_url_candidata(url: str, base_url: str = "") -> str:
    u = (url or "").strip()
    if not u:
        return ""

    if u.startswith("chrome-extension://"):
        try:
            q = urllib_parse.parse_qs(urllib_parse.urlparse(u).query)
            for k in ("file", "src", "url"):
                for v in q.get(k, []):
                    vv = urllib_parse.unquote(v or "").strip()
                    if vv.startswith("http://") or vv.startswith("https://"):
                        return vv
        except Exception:
            pass
        return ""

    if u.startswith("blob:"):
        return ""
    if u.startswith("//"):
        return f"https:{u}"
    if u.startswith("http://") or u.startswith("https://"):
        try:
            parsed = urllib_parse.urlparse(u)
            host = (parsed.netloc or "").lower()
            portal_host = ""
            try:
                portal_host = (urllib_parse.urlparse(_base_url_portal()).netloc or "").lower()
            except Exception:
                portal_host = ""

            # O portal retorna Location em http no POST do PDF, mas o navegador
            # imediatamente promove para https antes de abrir o arquivo.
            # Subimos direto para https para não depender do redirect intermediário.
            if parsed.scheme == "http" and host and host == portal_host:
                return urllib_parse.urlunparse(parsed._replace(scheme="https"))
        except Exception:
            pass
        return u
    if u.startswith("/"):
        return urllib_parse.urljoin(base_url or _base_url_portal(), u)
    return urllib_parse.urljoin(base_url or _base_url_portal(), u)


def coletar_urls_pdf_pos_visualizar(handles_base, timeout=12):
    """Coleta URLs candidatas ao PDF em janelas/tabs novas e estruturas internas."""
    fim = time.time() + timeout
    base = list(handles_base or [])
    vistos = set()
    coletadas = []

    while time.time() < fim:
        try:
            handles = list(driver.window_handles)
        except Exception:
            sleep(0.2)
            continue

        novos = [h for h in handles if h not in base]
        candidatos = novos + [h for h in handles if h in base]

        for h in candidatos:
            try:
                driver.switch_to.window(h)
            except Exception:
                continue

            try:
                atual = (driver.current_url or "").strip()
            except Exception:
                atual = ""

            try:
                dados = driver.execute_script(
                    """
                    const frames = Array.from(document.querySelectorAll("frame[src],iframe[src]"))
                      .map(f => f.getAttribute("src") || f.src || "");
                    const links = Array.from(document.querySelectorAll("a[href]"))
                      .map(a => a.getAttribute("href") || a.href || "");
                    return { frames, links };
                    """
                ) or {}
            except Exception:
                dados = {}

            candidatos_url = [atual]
            for u in (dados.get("frames") or []):
                candidatos_url.append(u)
            for u in (dados.get("links") or []):
                candidatos_url.append(u)

            for raw in candidatos_url:
                norm = _normalizar_url_candidata(raw, atual)
                if not norm:
                    continue
                if norm in vistos:
                    continue
                vistos.add(norm)
                coletadas.append(norm)

        if any(".pdf" in u.lower() or "/resultados/" in u.lower() for u in coletadas):
            break

        sleep(0.25)

    coletadas.sort(key=lambda u: 0 if (".pdf" in u.lower() or "/resultados/" in u.lower()) else 1)
    return coletadas


def aguardar_popup_pdf_tomados(handles_base, timeout=20) -> dict:
    """Espera a nova aba/janela do PDF de Tomados e captura a URL final do portal."""
    fim = time.time() + timeout
    base = list(handles_base or [])
    ultimo_url = ""
    ultimo_titulo = ""

    while time.time() < fim:
        try:
            handles = list(driver.window_handles)
        except Exception:
            sleep(0.2)
            continue

        novos = [h for h in handles if h not in base]
        for h in novos:
            try:
                driver.switch_to.window(h)
            except Exception:
                continue

            try:
                atual = (driver.current_url or "").strip()
            except Exception:
                atual = ""
            try:
                titulo = (driver.title or "").strip()
            except Exception:
                titulo = ""

            ultimo_url = atual or ultimo_url
            ultimo_titulo = titulo or ultimo_titulo

            candidatos = [atual]
            try:
                extra = driver.execute_script(
                    """
                    return [
                      window.location.href || "",
                      document.URL || "",
                      Array.from(document.querySelectorAll("embed[src],iframe[src],frame[src],object[data]"))
                        .map(el => el.getAttribute("src") || el.getAttribute("data") || ""),
                      Array.from(document.querySelectorAll("a[href]"))
                        .map(el => el.getAttribute("href") || ""),
                      performance.getEntries().map(e => e.name || "")
                    ].flat();
                    """
                ) or []
                candidatos.extend(extra if isinstance(extra, list) else [])
            except Exception:
                pass

            try:
                html_popup = driver.page_source or ""
            except Exception:
                html_popup = ""

            for raw in _extrair_urls_pdf_de_texto(html_popup, atual or _base_url_portal()):
                candidatos.append(raw)

            for raw in candidatos:
                norm = _normalizar_url_candidata(raw, atual or _base_url_portal())
                if norm and ("/resultados/" in norm.lower() or norm.lower().endswith(".pdf")):
                    try:
                        WebDriverWait(driver, 5).until(
                            lambda d: (d.execute_script("return document.readyState") or "").strip().lower() in {"interactive", "complete"}
                        )
                    except Exception:
                        pass
                    return {"handle": h, "url": norm, "title": titulo}

        sleep(DOWNLOAD_POLL_INTERVAL_SECONDS)

    raise RuntimeError(
        "Nova aba do PDF nao expôs URL utilizavel "
        f"(ultima_url={ultimo_url or 'n/a'} titulo={ultimo_titulo or 'n/a'})"
    )


def _extrair_url_pdf_de_fluxo(fluxo: dict) -> str:
    controle = fluxo.get("controle_302") or {}
    pdf_response = fluxo.get("pdf_response") or {}

    controle_url = _normalizar_url_candidata(
        controle.get("location") or "",
        controle.get("url") or _base_url_portal(),
    )
    if controle_url and ("/resultados/" in controle_url.lower() or controle_url.lower().endswith(".pdf")):
        return controle_url

    response_url = _normalizar_url_candidata(pdf_response.get("url") or "", _base_url_portal())
    if response_url and _is_pdf_response(
        int(pdf_response.get("status") or 0),
        pdf_response.get("contentType") or "",
        response_url,
    ):
        return response_url

    payload = pdf_response.get("body") or b""
    if payload:
        texto = _decodificar_bytes_resposta(payload)
        urls = _extrair_urls_pdf_de_texto(texto, response_url or controle_url or _base_url_portal())
        if urls:
            return urls[0]

    return ""


def baixar_pdf_via_popup_ou_logs(
    botao,
    handles_base,
    timeout=120,
    nome_fallback="arquivo.pdf",
    aceitar_alerta=False,
) -> tuple[str, dict]:
    """Usa o clique real do botão e captura o PDF via popup /resultados/*.pdf ou logs de rede."""
    inicio_perf = time.perf_counter()
    pdfs_antes = [
        nome
        for nome in os.listdir(TEMP_DOWNLOAD_DIR)
        if nome.lower().endswith(".pdf") or nome.lower().endswith(".crdownload")
    ]
    _limpar_logs_performance()
    click_robusto(botao)

    if aceitar_alerta:
        try:
            WebDriverWait(driver, 2).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
        except Exception:
            pass

    try:
        popup_info = aguardar_popup_pdf_tomados(
            handles_base,
            timeout=bounded_timeout(timeout, PDF_POPUP_WAIT_TIMEOUT, floor=4),
        )
        pdf_url = _normalizar_url_candidata((popup_info or {}).get("url") or "", _base_url_portal())
        if not pdf_url:
            raise RuntimeError("Popup do PDF nao expôs URL final utilizavel")

        popup_info = dict(popup_info or {})
        popup_info["url"] = pdf_url
        popup_info["source"] = "popup_url"
        pdf_tmp = baixar_pdf_por_url_com_cookies(pdf_url, timeout=timeout)
        perf_log("pdf_fetch", inicio_perf, "origem=popup_url")
        return pdf_tmp, popup_info
    except Exception as exc_popup:
        try:
            pdf_tmp = aguardar_pdf_novo(
                TEMP_DOWNLOAD_DIR,
                timeout=bounded_timeout(timeout, PDF_FILE_FALLBACK_TIMEOUT, floor=3),
                pdfs_antes=pdfs_antes,
            )
            perf_log("pdf_fetch", inicio_perf, "origem=arquivo_tmp")
            return pdf_tmp, {
                "handle": "",
                "url": "",
                "title": "",
                "source": "arquivo_tmp",
                "popup_error": f"{type(exc_popup).__name__}: {str(exc_popup)[:140]}",
            }
        except Exception as exc_file:
            if not ENABLE_CHROME_PERFORMANCE_LOGS:
                perf_log(
                    "pdf_fetch_falha",
                    inicio_perf,
                    f"popup={type(exc_popup).__name__} arquivo={type(exc_file).__name__}",
                )
                raise exc_popup

        fluxo = capturar_fluxo_pdf_via_logs(
            timeout=bounded_timeout(timeout, PDF_NETWORK_CAPTURE_TIMEOUT, floor=4)
        )
        pdf_tmp, origem_fluxo = materializar_pdf_de_fluxo_rede(
            fluxo,
            nome_fallback=nome_fallback,
            timeout=timeout,
        )
        if not pdf_tmp:
            perf_log("pdf_fetch_falha", inicio_perf, f"popup={type(exc_popup).__name__}")
            raise exc_popup
        perf_log("pdf_fetch", inicio_perf, f"origem={origem_fluxo or 'logs_rede'}")
        return pdf_tmp, {
            "handle": "",
            "url": _extrair_url_pdf_de_fluxo(fluxo),
            "title": "",
            "source": origem_fluxo or "logs_rede_url",
            "popup_error": f"{type(exc_popup).__name__}: {str(exc_popup)[:140]}",
        }


def baixar_pdf_tomados_via_popup(btn_pdf, handles_base, timeout=120, nome_fallback="arquivo.pdf") -> tuple[str, dict]:
    """Compatibilidade do fluxo de tomados individual com a helper genérica de popup."""
    return baixar_pdf_via_popup_ou_logs(
        btn_pdf,
        handles_base,
        timeout=timeout,
        nome_fallback=nome_fallback,
        aceitar_alerta=False,
    )


def baixar_pdf_por_url_com_cookies(pdf_url: str, timeout=120, referer: str = "") -> str:
    """Baixa PDF diretamente pela URL, reaproveitando cookies da sessão Selenium."""
    url = (pdf_url or "").strip()
    if not url:
        raise RuntimeError("URL do PDF vazia")

    cookie_header = "; ".join(
        f"{c.get('name','')}={c.get('value','')}"
        for c in (driver.get_cookies() or [])
        if c.get("name")
    )
    user_agent = ""
    try:
        user_agent = driver.execute_script("return navigator.userAgent;") or ""
    except Exception:
        user_agent = ""

    headers = {
        "User-Agent": user_agent or "Mozilla/5.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "Upgrade-Insecure-Requests": "1",
    }
    if cookie_header:
        headers["Cookie"] = cookie_header
    referer_header = (referer or "").strip()
    if not referer_header:
        referer_header = _base_url_portal().rstrip("/")
        if referer_header:
            referer_header = referer_header + "/"
    if referer_header:
        headers["Referer"] = referer_header

    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    req = urllib_request.Request(url, headers=headers, method="GET")
    with urllib_request.urlopen(req, timeout=timeout, context=ssl_ctx) as resp:
        data = resp.read()

    if not data or len(data) < 20:
        raise RuntimeError("Resposta de PDF vazia ou muito pequena")
    if not data.startswith(b"%PDF"):
        raise RuntimeError("Resposta não parece PDF válido")

    nome = os.path.basename(urllib_parse.urlparse(url).path) or f"livro_{int(time.time())}.pdf"
    if not nome.lower().endswith(".pdf"):
        nome = f"{nome}.pdf"

    destino = os.path.join(TEMP_DOWNLOAD_DIR, nome)
    if os.path.exists(destino):
        base, ext = os.path.splitext(nome)
        k = 1
        while True:
            cand = os.path.join(TEMP_DOWNLOAD_DIR, f"{base}_{k}{ext}")
            if not os.path.exists(cand):
                destino = cand
                break
            k += 1

    with open(destino, "wb") as f:
        f.write(data)

    return destino


def fechar_modal_tomados_visualizacao(timeout=8) -> bool:
    """Fecha o modal de visualização da NFSe em Tomados."""
    try:
        botoes = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'modal') and not(contains(@style,'display: none'))]"
            "//button[contains(@class,'close') or @data-dismiss='modal' or contains(.,'Fechar')]"
        )
        for btn in botoes:
            try:
                if not btn.is_displayed():
                    continue
                click_robusto(btn)
                break
            except Exception:
                continue
    except Exception:
        pass

    try:
        driver.execute_script(
            """
            try { $('.modal').modal('hide'); } catch (e) {}
            try { $('.modal-backdrop').remove(); } catch (e) {}
            try { $('.ui-widget-overlay').remove(); } catch (e) {}
            try { document.body.classList.remove('modal-open'); } catch (e) {}
            """
        )
    except Exception:
        pass

    try:
        WebDriverWait(driver, timeout).until(
            lambda d: len(
                [
                    el for el in d.find_elements(By.XPATH, "//button[@id='btnPDF' or @id='btnXML']")
                    if el.is_displayed()
                ]
            ) == 0
        )
    except Exception:
        pass

    return True

# =====================
# FECHAR MODAL EXPORTACAO
# =====================
def fechar_modal_exportacao():
    try:
        driver.execute_script("""
            try{ $('#modalboxexportarnotas').modal('hide'); }catch(e){}
            try{ $('.modal').modal('hide'); }catch(e){}
            try{ $('.modal-backdrop').remove(); }catch(e){}
            try{ $('.ui-widget-overlay').remove(); }catch(e){}
        """)
        WebDriverWait(driver, 8).until(
            EC.invisibility_of_element_located((By.ID, "modalboxexportarnotas"))
        )
    except Exception:
        pass


def baixar_pdf_nota_tomados_por_indice(i: int, ano_alvo: int, mes_alvo: int, ref_alvo: str) -> str:
    """Baixa o PDF da NFSe de uma linha de Tomados por request autenticado, sem depender do viewer do Chrome."""
    checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
    if i >= len(checkboxes):
        return "EOF"

    checkbox = checkboxes[i]
    marcar_somente_checkbox(checkbox, full_reset=True)
    WebDriverWait(driver, 10).until(
        lambda d: len([c for c in d.find_elements(By.NAME, "gridListaCheck") if c.is_selected()]) == 1
    )

    info = extrair_info_linha_tomados(checkbox)
    tr = checkbox.find_element(By.XPATH, "./ancestor::tr[1]")
    handles_base = []
    modal_aberto = False

    try:
        handles_base = list(driver.window_handles)
    except Exception:
        handles_base = []

    nf = info.get("nf", "")
    data_emissao = info.get("data_emissao", "")
    nome_fallback_pdf = f"NFT_{re.sub(r'\\D', '', nf or '') or 'nfse_tomados'}.pdf"
    pdf_url = ""
    falhas = []

    def _reobter_btn_pdf():
        btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "btnPDF"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        return btn

    try:
        icone_nfse = tr.find_element(
            By.XPATH,
            ".//span[contains(@title,'Visualizar NFSe') or contains(@title,'Visualizar NFSe')]"
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", icone_nfse)
        click_robusto(icone_nfse)
        modal_aberto = True

        btn_pdf = _reobter_btn_pdf()
        pdf_tmp = ""
        origem = ""
        fluxo = {}
        popup_info = {}

        try:
            pdf_tmp, popup_info = baixar_pdf_tomados_via_popup(
                btn_pdf,
                handles_base,
                timeout=120,
                nome_fallback=nome_fallback_pdf,
            )
            origem = popup_info.get("source") or "popup_url"
            pdf_url = popup_info.get("url") or ""
            salvar_log_tomados(
                "PDF_POPUP_URL",
                ref_alvo,
                mensagem=(
                    f"NF={nf} DATA={data_emissao} "
                    f"URL={pdf_url} TITULO={popup_info.get('title') or ''}"
                ),
            )
            salvar_log_tomados(
                "PDF_POPUP_FETCH",
                ref_alvo,
                mensagem=f"NF={nf} DATA={data_emissao} via {origem}",
            )
        except Exception as exc_popup:
            falhas.append(f"popup={type(exc_popup).__name__}:{str(exc_popup)[:140]}")
            salvar_log_tomados(
                "PDF_POPUP_FAIL",
                ref_alvo,
                mensagem=f"NF={nf} DATA={data_emissao} | {type(exc_popup).__name__}: {str(exc_popup)[:140]}",
            )

        if not pdf_tmp:
            try:
                fluxo = capturar_fluxo_pdf_via_logs(timeout=20)
                pdf_tmp, origem_fluxo = materializar_pdf_de_fluxo_rede(
                    fluxo,
                    nome_fallback=nome_fallback_pdf,
                    timeout=120,
                )
                if pdf_tmp:
                    origem = origem_fluxo
                    controle = fluxo.get("controle_302") or {}
                    pdf_response = fluxo.get("pdf_response") or {}
                    pdf_url = _normalizar_url_candidata(
                        (controle.get("location") or "") or (pdf_response.get("url") or ""),
                        _base_url_portal(),
                    )
                    salvar_log_tomados(
                        "PDF_NET_CAPTURE",
                        ref_alvo,
                        mensagem=f"NF={nf} DATA={data_emissao} via {origem_fluxo} URL={pdf_url}",
                    )
            except Exception as exc_diag:
                falhas.append(f"diag_logs={type(exc_diag).__name__}:{str(exc_diag)[:140]}")
                salvar_log_tomados(
                    "PDF_NET_CAPTURE_FAIL",
                    ref_alvo,
                    mensagem=f"NF={nf} DATA={data_emissao} | {type(exc_diag).__name__}: {str(exc_diag)[:140]}",
                )
            finally:
                try:
                    fechar_janelas_extras(handles_base)
                except Exception:
                    pass

        if not pdf_tmp:
            try:
                btn_pdf = _reobter_btn_pdf()
                pdf_tmp = baixar_arquivo_via_submit_form(
                    btn_pdf,
                    timeout=120,
                    nome_fallback=nome_fallback_pdf,
                    tipo_esperado="pdf",
                    debug_dump_tag="submit_form_pdf_tomados",
                    debug_info=info,
                    debug_ref_alvo=ref_alvo,
                )
                origem = "submit_form"
                salvar_log_tomados(
                    "PDF_SUBMIT_FETCH",
                    ref_alvo,
                    mensagem=f"NF={nf} DATA={data_emissao} via submit_form",
                )
            except Exception as exc_submit:
                falhas.append(f"submit_form={type(exc_submit).__name__}:{str(exc_submit)[:140]}")
                salvar_log_tomados(
                    "PDF_SUBMIT_FETCH_FAIL",
                    ref_alvo,
                    mensagem=f"NF={nf} DATA={data_emissao} | {type(exc_submit).__name__}: {str(exc_submit)[:140]}",
                )
                try:
                    fechar_janelas_extras(handles_base)
                except Exception:
                    pass

        if not pdf_tmp:
            try:
                btn_pdf = _reobter_btn_pdf()
                pdf_tmp = baixar_pdf_tomados_via_post_controle(btn_pdf, info, timeout=120, ref_alvo=ref_alvo)
                origem = "post_controle"
                salvar_log_tomados(
                    "PDF_POST_CONTROLE",
                    ref_alvo,
                    mensagem=f"NF={nf} DATA={data_emissao} via post_controle",
                )
            except Exception as exc_post:
                falhas.append(f"post_controle={type(exc_post).__name__}:{str(exc_post)[:140]}")
                salvar_log_tomados(
                    "PDF_POST_CONTROLE_FAIL",
                    ref_alvo,
                    mensagem=f"NF={nf} DATA={data_emissao} | {type(exc_post).__name__}: {str(exc_post)[:140]}",
                )
                try:
                    fechar_janelas_extras(handles_base)
                except Exception:
                    pass

        if not pdf_tmp:
            controle = fluxo.get("controle_302") or {}
            pdf_response = fluxo.get("pdf_response") or {}
            controle_location = _normalizar_url_candidata(controle.get("location") or "", controle.get("url") or _base_url_portal())
            if controle_location and ("/resultados/" in controle_location.lower() and controle_location.lower().endswith(".pdf")):
                pdf_url = controle_location
                salvar_log_tomados(
                    "PDF_CONTROLE_302",
                    ref_alvo,
                    mensagem=f"NF={nf} LOCATION={controle_location}",
                )
            else:
                controle_location = ""

            pdf_response_url = _normalizar_url_candidata(pdf_response.get("url") or "", _base_url_portal())
            if pdf_response_url and _is_pdf_response(
                int(pdf_response.get("status") or 0),
                pdf_response.get("contentType") or "",
                pdf_response_url,
            ):
                if not pdf_url:
                    pdf_url = pdf_response_url
                salvar_log_tomados(
                    "PDF_NET_CAPTURE",
                    ref_alvo,
                    mensagem=f"NF={nf} URL={pdf_response_url}",
                )
            else:
                pdf_response_url = ""

            if controle_location or pdf_response_url:
                bucket = CODIGO_ERRO_TOMADOS_PDF_RESPOSTA_INVALIDA
                detalhes = "; ".join(falhas[-3:]) or "Falha ao materializar PDF apos resposta valida do portal"
                registrar_status_tomados_pdf(
                    bucket,
                    detalhes,
                    ref_alvo,
                    nf=nf,
                    data_emissao=data_emissao,
                    pdf_url=pdf_url,
                )
                raise RuntimeError(f"{bucket}: {detalhes}")

            bucket = CODIGO_ERRO_TOMADOS_PDF_SEM_LOCATION
            detalhes = "; ".join(falhas[-3:]) or "Fluxo nao retornou 302 + Location nem GET /resultados/*.pdf"
            registrar_status_tomados_pdf(
                bucket,
                detalhes,
                ref_alvo,
                nf=nf,
                data_emissao=data_emissao,
            )
            raise RuntimeError(f"{bucket}: {detalhes}")

        destino_pdf = salvar_pdf_tomados_nota(pdf_tmp, info, ano_alvo, mes_alvo)
        limpar_status_tomados_pdf(referencia=ref_alvo)
        salvar_log_tomados(
            "PDF_OK",
            ref_alvo,
            arquivo_pdf=destino_pdf,
            mensagem=f"NF={nf} DATA={data_emissao} via {origem}",
        )
        return destino_pdf
    except Exception as exc:
        bucket = _bucket_tomados_pdf_from_message(str(exc)) or CODIGO_ERRO_TOMADOS_PDF_NAO_GERADO
        if not os.path.exists(caminho_status_tomados_pdf(referencia=ref_alvo)):
            detalhes = f"{type(exc).__name__}: {str(exc)[:180]}"
            registrar_status_tomados_pdf(
                bucket,
                detalhes,
                ref_alvo,
                nf=nf,
                data_emissao=data_emissao,
                pdf_url=pdf_url,
            )
        raise
    finally:
        try:
            fechar_janelas_extras(handles_base)
        except Exception:
            pass
        if modal_aberto:
            try:
                fechar_modal_tomados_visualizacao(timeout=8)
            except Exception:
                pass
        try:
            driver.switch_to.default_content()
        except Exception:
            pass


def baixar_pdfs_servicos_tomados(ano_alvo: int, mes_alvo: int, ref_alvo: str) -> int:
    """Percorre a grid de Tomados e baixa o PDF de todas as notas do período."""
    total_pdfs = 0
    visitados = set()
    pagina = 1

    try:
        desmarcar_todas_notas()
    except Exception:
        pass

    while True:
        checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
        total = len(checkboxes)
        print(f"[TOMADOS] Pagina {pagina}: {total} notas para baixar PDF.")

        if total == 0:
            break

        i = 0
        while i < total:
            checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
            if i >= len(checkboxes):
                break

            info = extrair_info_linha_tomados(checkboxes[i])
            key = chave_unica(info)
            if key in visitados:
                i += 1
                continue

            visitados.add(key)
            try:
                destino_pdf = baixar_pdf_nota_tomados_por_indice(i, ano_alvo, mes_alvo, ref_alvo)
                if destino_pdf != "EOF":
                    total_pdfs += 1
                    print(f"[TOMADOS] PDF OK -> {os.path.basename(destino_pdf)}")
            except Exception as e:
                msg = f"NF={info.get('nf','')} DATA={info.get('data_emissao','')} | {type(e).__name__}: {str(e)[:160]}"
                salvar_log_tomados("PDF_ERRO", ref_alvo, mensagem=msg)
                salvar_print_evidencia(
                    "SERVICOS_TOMADOS",
                    "PDF_NFSE_FALHA",
                    ano_alvo=ano_alvo,
                    mes_alvo=mes_alvo,
                    ref_alvo=ref_alvo,
                    log_path=tomados_log_path_empresa(ref_alvo),
                )
                raise RuntimeError(f"{MSG_TOMADOS_PDF_REVISAO_MANUAL}: {msg}")
            i += 1

        if not ir_para_proxima_pagina():
            break
        pagina += 1

    salvar_log_tomados("PDFS_OK", ref_alvo, mensagem=f"Total de PDFs baixados: {total_pdfs}")
    return total_pdfs

def detectar_nome_empresa_da_tela():
    try:
        els = driver.find_elements(By.XPATH, "//*[contains(.,'Empresa:')]")
        for el in els:
            t = (el.text or "").strip().replace("\n", " ")
            if not t or len(t) > 250:
                continue
            m = re.search(r"Empresa:\s*.+?\s*-\s*(.+)$", t)
            if m:
                return m.group(1).strip()
    except Exception:
        pass
    return ""

EMPRESA_PASTA = EMPRESA_PASTA_FORCADA.strip() or "EMPRESA_DESCONHECIDA"


def resolver_pasta_empresa_atual() -> str:
    global EMPRESA_PASTA

    if EMPRESA_PASTA_FORCADA.strip():
        EMPRESA_PASTA = EMPRESA_PASTA_FORCADA.strip()
    else:
        nome_detectado = normalizar_nome_empresa(detectar_nome_empresa_da_tela())
        if nome_detectado:
            EMPRESA_PASTA = nome_detectado

    print(f"Pasta da empresa: {EMPRESA_PASTA}")
    return EMPRESA_PASTA


def resetar_contexto_cadastro() -> None:
    global CONTEXTO_CADASTRO_ATUAL, MULTI_CADASTRO_ATIVO
    CONTEXTO_CADASTRO_ATUAL = {}
    MULTI_CADASTRO_ATIVO = False
    resetar_multicadastro_tardio()


def _estado_multicadastro_tardio_vazio() -> dict:
    return {
        "ativo": False,
        "cadastros": [],
        "indice_atual": -1,
        "resultados": [],
        "concluidos": [],
        "falhos": [],
    }


def resetar_multicadastro_tardio() -> None:
    global MULTI_CADASTRO_TARDIO_ESTADO
    MULTI_CADASTRO_TARDIO_ESTADO = _estado_multicadastro_tardio_vazio()


def estado_multicadastro_tardio() -> dict:
    estado = dict(MULTI_CADASTRO_TARDIO_ESTADO or {})
    estado["cadastros"] = [dict(item) for item in (estado.get("cadastros") or [])]
    estado["resultados"] = [dict(item) for item in (estado.get("resultados") or []) if isinstance(item, dict)]
    estado["concluidos"] = list(estado.get("concluidos") or [])
    estado["falhos"] = list(estado.get("falhos") or [])
    return estado


def multicadastro_tardio_ativo() -> bool:
    estado = MULTI_CADASTRO_TARDIO_ESTADO or {}
    return bool(estado.get("ativo") and estado.get("cadastros"))


def _fluxo_multicadastro_tardio_local_ativo() -> bool:
    estado = MULTI_CADASTRO_TARDIO_ESTADO or {}
    return bool(estado.get("ativo") and estado.get("execucao_local"))


def _definir_fluxo_multicadastro_tardio_local(ativo: bool) -> None:
    global MULTI_CADASTRO_TARDIO_ESTADO
    estado = dict(MULTI_CADASTRO_TARDIO_ESTADO or {})
    estado["execucao_local"] = bool(ativo)
    MULTI_CADASTRO_TARDIO_ESTADO = estado


def cadastro_multicadastro_tardio_atual() -> dict:
    estado = MULTI_CADASTRO_TARDIO_ESTADO or {}
    cadastros = estado.get("cadastros") or []
    try:
        indice = int(estado.get("indice_atual"))
    except Exception:
        indice = -1
    if 0 <= indice < len(cadastros):
        return dict(cadastros[indice])
    return {}


def inicializar_multicadastro_tardio(cadastros: list[dict], indice_atual: int = 0) -> dict:
    global MULTI_CADASTRO_TARDIO_ESTADO
    lista = [dict(item) for item in (cadastros or [])]
    if not lista:
        resetar_multicadastro_tardio()
        return {}

    indice = max(0, min(int(indice_atual or 0), len(lista) - 1))
    MULTI_CADASTRO_TARDIO_ESTADO = {
        "ativo": True,
        "cadastros": lista,
        "indice_atual": indice,
        "resultados": [],
        "concluidos": [],
        "falhos": [],
    }
    return ativar_cadastro_multicadastro_tardio(indice)


def ativar_cadastro_multicadastro_tardio(indice: int) -> dict:
    global MULTI_CADASTRO_TARDIO_ESTADO
    cadastros = [dict(item) for item in (MULTI_CADASTRO_TARDIO_ESTADO.get("cadastros") or [])]
    if not cadastros:
        raise RuntimeError("Multi-cadastro tardio não inicializado.")

    idx = int(indice or 0)
    if idx < 0 or idx >= len(cadastros):
        raise RuntimeError(f"Índice de cadastro tardio inválido: {idx}")

    cadastro = dict(cadastros[idx])
    MULTI_CADASTRO_TARDIO_ESTADO["ativo"] = True
    MULTI_CADASTRO_TARDIO_ESTADO["indice_atual"] = idx
    MULTI_CADASTRO_TARDIO_ESTADO["cadastros"] = cadastros
    ativar_contexto_cadastro(cadastro)
    return cadastro


def registrar_resultado_multicadastro_tardio(resultado: dict) -> None:
    global MULTI_CADASTRO_TARDIO_ESTADO

    if not isinstance(resultado, dict):
        return

    cadastro = dict(resultado.get("cadastro") or {})
    try:
        ordem = int(cadastro.get("ordem") or 0)
    except Exception:
        ordem = 0

    resultados = [
        dict(item)
        for item in (MULTI_CADASTRO_TARDIO_ESTADO.get("resultados") or [])
        if int((item.get("cadastro") or {}).get("ordem") or 0) != ordem
    ]
    resultados.append(dict(resultado))
    MULTI_CADASTRO_TARDIO_ESTADO["resultados"] = resultados

    concluidos = [int(item) for item in (MULTI_CADASTRO_TARDIO_ESTADO.get("concluidos") or []) if int(item or 0) != ordem]
    falhos = [int(item) for item in (MULTI_CADASTRO_TARDIO_ESTADO.get("falhos") or []) if int(item or 0) != ordem]
    if ordem > 0:
        if resultado.get("status") == "FALHA":
            falhos.append(ordem)
        else:
            concluidos.append(ordem)
    MULTI_CADASTRO_TARDIO_ESTADO["concluidos"] = concluidos
    MULTI_CADASTRO_TARDIO_ESTADO["falhos"] = falhos


def ativar_contexto_cadastro(cadastro: dict | None) -> dict:
    global CONTEXTO_CADASTRO_ATUAL, MULTI_CADASTRO_ATIVO

    if not cadastro:
        resetar_contexto_cadastro()
        return {}

    try:
        ordem = int(cadastro.get("ordem") or 0)
    except Exception:
        ordem = 0
    try:
        page = int(cadastro.get("page") or 0)
    except Exception:
        page = 0
    try:
        row_index = int(cadastro.get("row_index") or 0)
    except Exception:
        row_index = 0

    contexto = {
        "ordem": ordem,
        "page": page,
        "row_index": row_index,
        "row_key": str(cadastro.get("row_key") or "").strip(),
        "valor": str(cadastro.get("valor") or "").strip(),
        "ccm": re.sub(r"\D", "", str(cadastro.get("ccm") or "")),
        "identificacao": re.sub(r"\s+", " ", str(cadastro.get("identificacao") or "").strip()),
    }
    CONTEXTO_CADASTRO_ATUAL = contexto
    MULTI_CADASTRO_ATIVO = True
    return dict(contexto)


def contexto_cadastro_atual() -> dict:
    return dict(CONTEXTO_CADASTRO_ATUAL or {})


def _prefixo_cadastro_base(cadastro: dict | None = None) -> str:
    ctx = dict(cadastro or contexto_cadastro_atual())
    if not ctx:
        return ""

    partes = []
    ccm = re.sub(r"\D", "", str(ctx.get("ccm") or ctx.get("cadastro_ccm") or ""))
    if ccm:
        partes.append(f"CCM_{ccm}")

    try:
        ordem = int(ctx.get("ordem") or ctx.get("cadastro_ordem") or 0)
    except Exception:
        ordem = 0
    if ordem > 0:
        partes.append(f"CAD_{ordem:02d}")
    elif ctx.get("identificacao") or ctx.get("cadastro_identificacao"):
        partes.append("CAD_00")

    return "__".join(partes)


def prefixo_cadastro_arquivo(cadastro: dict | None = None) -> str:
    base = _prefixo_cadastro_base(cadastro)
    return f"{base}__" if base else ""


def prefixar_nome_arquivo_cadastro(nome_arquivo: str, cadastro: dict | None = None) -> str:
    prefixo = prefixo_cadastro_arquivo(cadastro)
    return f"{prefixo}{nome_arquivo}" if prefixo else nome_arquivo


def enriquecer_info_com_cadastro(info: dict | None) -> dict:
    payload = dict(info or {})
    ctx = contexto_cadastro_atual()
    if not ctx:
        return payload

    payload.setdefault("cadastro_ordem", ctx.get("ordem") or "")
    payload.setdefault("cadastro_ccm", ctx.get("ccm") or "")
    payload.setdefault("cadastro_identificacao", ctx.get("identificacao") or "")
    return payload


def _partes_contexto_cadastro(info: dict | None = None) -> list[str]:
    payload = dict(info or {})
    ordem = payload.get("cadastro_ordem")
    ccm = payload.get("cadastro_ccm")
    identificacao = payload.get("cadastro_identificacao")

    if not any([ordem, ccm, identificacao]):
        ctx = contexto_cadastro_atual()
        ordem = ctx.get("ordem") or ordem
        ccm = ctx.get("ccm") or ccm
        identificacao = ctx.get("identificacao") or identificacao

    partes = []
    if ordem:
        partes.append(f"CADASTRO_ORDEM={ordem}")
    if ccm:
        partes.append(f"CADASTRO_CCM={ccm}")
    if identificacao:
        partes.append(f"CADASTRO_IDENTIFICACAO={(identificacao or '').replace('|', '/')}")
    return partes


def _normalizar_token_chave(txt: str) -> str:
    token = re.sub(r"\s+", "", str(txt or "").strip())
    token = token.replace("|", "_")
    token = re.sub(r"[^A-Za-z0-9._:-]+", "_", token)
    return token.strip("_") or "SEM_CHAVE"

# =====================
# LOG (1 por empresa) + IDP
# =====================
def pasta_empresa():
    return os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA)


def pasta_competencia(ano_alvo: int = 0, mes_alvo: int = 0, competencia: str = "") -> str:
    base_dir = pasta_empresa()
    competencia = (competencia or "").strip()
    if competencia:
        return os.path.join(base_dir, competencia)
    if ano_alvo and mes_alvo:
        return os.path.join(base_dir, f"{mes_alvo:02d}.{ano_alvo}")
    return base_dir


def normalizar_pasta_servico(servico: str = "") -> str:
    nome = (servico or "").strip().upper()
    if nome == "PRESTADOS":
        return "PRESTADOS"
    if nome == "TOMADOS":
        return "TOMADOS"
    return "_GERAL"


def pasta_servico(servico: str = "", ano_alvo: int = 0, mes_alvo: int = 0, competencia: str = "") -> str:
    return os.path.join(
        pasta_competencia(ano_alvo=ano_alvo, mes_alvo=mes_alvo, competencia=competencia),
        normalizar_pasta_servico(servico),
    )


def parse_referencia_competencia(ref_alvo: str) -> tuple[int, int]:
    ref = normalizar_referencia_competencia(ref_alvo)
    m = re.fullmatch(r"(\d{2})/(\d{4})", ref)
    if not m:
        return 0, 0
    return int(m.group(2)), int(m.group(1))


def competencia_execucao_padrao(apurar_completo: bool | None = None) -> tuple[int, int]:
    if apurar_completo is None:
        apurar_completo = APURAR_COMPLETO

    apuracao_ref = APURACAO_REFERENCIA or datetime.now().strftime("%m/%Y")
    try:
        competencias = runtime_competencias_alvo(apuracao_ref, apurar_completo=apurar_completo)
        if competencias:
            if apurar_completo:
                return competencias[-1]
            return competencias[0]
    except Exception:
        pass

    agora = datetime.now()
    if agora.month == 1:
        return agora.year - 1, 12
    return agora.year, agora.month - 1


def pasta_geral(ano_alvo: int = 0, mes_alvo: int = 0, competencia: str = "", apurar_completo: bool | None = None) -> str:
    if not (ano_alvo and mes_alvo) and not (competencia or "").strip():
        ano_alvo, mes_alvo = competencia_execucao_padrao(apurar_completo=apurar_completo)
    return pasta_servico("_GERAL", ano_alvo=ano_alvo, mes_alvo=mes_alvo, competencia=competencia)


def pasta_debug(ano_alvo: int = 0, mes_alvo: int = 0, competencia: str = "", apurar_completo: bool | None = None) -> str:
    return os.path.join(
        pasta_geral(
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            competencia=competencia,
            apurar_completo=apurar_completo,
        ),
        "_debug",
    )


def caminho_arquivo_geral(
    nome_arquivo: str,
    ano_alvo: int = 0,
    mes_alvo: int = 0,
    competencia: str = "",
    apurar_completo: bool | None = None,
) -> str:
    return os.path.join(
        pasta_geral(
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            competencia=competencia,
            apurar_completo=apurar_completo,
        ),
        nome_arquivo,
    )


def caminho_log_manual(ref_alvo: str = "", ano_alvo: int = 0, mes_alvo: int = 0) -> str:
    if ref_alvo and not (ano_alvo and mes_alvo):
        ano_alvo, mes_alvo = parse_referencia_competencia(ref_alvo)
    return caminho_arquivo_geral("log_fechamento_manual.txt", ano_alvo=ano_alvo, mes_alvo=mes_alvo)


def caminho_log_apuracao_completa(ref_alvo: str = "", ano_alvo: int = 0, mes_alvo: int = 0) -> str:
    if ref_alvo and not (ano_alvo and mes_alvo):
        ano_alvo, mes_alvo = parse_referencia_competencia(ref_alvo)
    return caminho_arquivo_geral(
        "log_apuracao_completa.txt",
        ano_alvo=ano_alvo,
        mes_alvo=mes_alvo,
        apurar_completo=True,
    )


def _servico_evidencia_por_contexto(contexto: str) -> str:
    contexto_up = (contexto or "").strip().upper()
    if "TOMADOS" in contexto_up:
        return "TOMADOS"
    if "PRESTADOS" in contexto_up:
        return "PRESTADOS"
    return "_GERAL"


def caminho_status_tomados_pdf(referencia: str = "", ano_alvo: int = 0, mes_alvo: int = 0) -> str:
    if referencia and not (ano_alvo and mes_alvo):
        ano_alvo, mes_alvo = parse_referencia_competencia(referencia)
    nome_arquivo = prefixar_nome_arquivo_cadastro(ARQUIVO_STATUS_TOMADOS_PDF)
    return caminho_arquivo_geral(nome_arquivo, ano_alvo=ano_alvo, mes_alvo=mes_alvo)


def limpar_status_tomados_pdf(referencia: str = "", ano_alvo: int = 0, mes_alvo: int = 0):
    try:
        path = caminho_status_tomados_pdf(referencia=referencia, ano_alvo=ano_alvo, mes_alvo=mes_alvo)
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def registrar_status_tomados_pdf(codigo_erro: str, detalhes: str = "", referencia: str = "", nf: str = "", data_emissao: str = "", pdf_url: str = ""):
    payload = {
        "codigo_erro": (codigo_erro or "").strip(),
        "detalhes": (detalhes or "").strip(),
        "referencia": (referencia or "").strip(),
        "nf": (nf or "").strip(),
        "data_emissao": (data_emissao or "").strip(),
        "pdf_url": (pdf_url or "").strip(),
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "cadastro": contexto_cadastro_atual(),
    }
    try:
        path_status = caminho_status_tomados_pdf(referencia=referencia)
        os.makedirs(os.path.dirname(path_status), exist_ok=True)
        with open(path_status, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _bucket_tomados_pdf_from_message(msg: str) -> str:
    texto = (msg or "").strip()
    for codigo in (
        CODIGO_ERRO_TOMADOS_PDF_SEM_LOCATION,
        CODIGO_ERRO_TOMADOS_PDF_RESPOSTA_INVALIDA,
        CODIGO_ERRO_TOMADOS_PDF_NAO_GERADO,
    ):
        if codigo in texto:
            return codigo
    return ""


def log_path_empresa(info: dict | None = None):
    info = info or {}
    if APURAR_COMPLETO:
        parsed = parse_data_emissao_site(info.get("data_emissao", ""))
        if parsed:
            _, _, _, _, competencia = parsed
            return caminho_arquivo_geral(LOG_FILENAME, competencia=competencia)
    return caminho_arquivo_geral(LOG_FILENAME)

def garantir_log_com_header(path_):
    """Agora garante um .txt simples (uma linha por evento)."""
    os.makedirs(os.path.dirname(path_), exist_ok=True)
    if not os.path.exists(path_):
        with open(path_, "w", encoding="utf-8") as f:
            f.write("# Log NFSe (uma linha por evento)\n")

def carregar_chaves_ok(path_):
    """Carrega chaves já baixadas com STATUS=OK a partir do log .txt.
    Espera linhas contendo 'CHAVE_UNICA=<valor>' e 'STATUS=OK'.
    """
    chaves = set()
    if not os.path.exists(path_):
        return chaves

    try:
        with open(path_, "r", encoding="utf-8") as f:
            for line in f:
                if "STATUS=OK" not in line:
                    continue
                m = re.search(r"CHAVE_UNICA=([^\s\|]+)", line)
                if m:
                    chaves.add(m.group(1).strip())
    except Exception:
        pass

    return chaves

def salvar_log(destino_xml, info, status="OK", mensagem=""):
    info = enriquecer_info_com_cadastro(info)
    path_ = log_path_empresa(info)
    garantir_log_com_header(path_)

    key = chave_unica(info)
    destino_existe = bool(destino_xml) and os.path.isfile(destino_xml)
    arquivo_xml = os.path.basename(destino_xml) if destino_existe else ""
    pasta_xml = os.path.dirname(destino_xml) if destino_existe else ""

    parts = [
        f"STATUS={status}",
        f"CHAVE_UNICA={key}",
        f"NF={info.get('nf','')}",
        f"DATA={info.get('data_emissao','')}",
        f"RPS={info.get('rps','')}",
        f"CHAVE={info.get('chave','')}",
        f"ID_INTERNO={info.get('id_interno','')}",
    ]
    parts.extend(_partes_contexto_cadastro(info))
    if arquivo_xml:
        parts.append(f"ARQ={arquivo_xml}")
    if pasta_xml:
        parts.append(f"PASTA={pasta_xml}")
    if mensagem:
        parts.append(f"MSG={(mensagem or '').replace('|','/')}")

    append_txt(path_, " | ".join(parts))
    return path_

def mover_com_retry(src, dst, tentativas=6):
    last = None
    for i in range(tentativas):
        try:
            shutil.move(src, dst)
            return
        except Exception as e:
            last = e
            sleep(0.3 + i * 0.4)
    raise last


def _competencia_saida_geral(ano_alvo: int = 0, mes_alvo: int = 0, competencia: str = "") -> str:
    competencia = (competencia or "").strip()
    if competencia:
        return competencia
    if ano_alvo and mes_alvo:
        return f"{mes_alvo:02d}.{ano_alvo}"
    return ""


def pasta_saida_geral(raiz: str, ano_alvo: int = 0, mes_alvo: int = 0, competencia: str = "") -> str:
    competencia_dir = _competencia_saida_geral(ano_alvo=ano_alvo, mes_alvo=mes_alvo, competencia=competencia)
    if not competencia_dir:
        competencia_dir = "SEM_COMPETENCIA"
    return os.path.join(OUTPUT_BASE_DIR, "SAIDAS_GERAIS", (raiz or "").strip().upper(), competencia_dir)


def _identidade_saida_geral_atual() -> tuple[str, str]:
    ctx = contexto_cadastro_atual()
    codigo = normalizar_codigo_empresa(ctx.get("ccm") or ctx.get("codigo") or "")
    slug_origem = (
        globals().get("EMPRESA_PASTA")
        or ctx.get("identificacao")
        or ""
    )
    slug = normalizar_nome_empresa(slug_origem)
    if not slug:
        slug = "EMPRESA_DESCONHECIDA"
    if codigo and slug.startswith(f"{codigo}_"):
        codigo = ""
    return codigo, slug


def _nome_arquivo_saida_geral(caminho_origem: str, tipo_documento: str, ano_alvo: int, mes_alvo: int) -> str:
    codigo, slug = _identidade_saida_geral_atual()
    partes = []
    if codigo:
        partes.append(codigo)
    if slug and slug not in partes:
        partes.append(slug)
    if not partes:
        partes.append("EMPRESA_DESCONHECIDA")

    tipo = re.sub(r"[^A-Z0-9_]+", "_", (tipo_documento or "").strip().upper()).strip("_") or "DOCUMENTO"
    competencia = f"{mes_alvo:02d}-{ano_alvo}" if ano_alvo and mes_alvo else "SEM_COMPETENCIA"
    partes.extend([tipo, competencia])

    nome_base = re.sub(r"_+", "_", "_".join(partes)).strip("_")
    ext = os.path.splitext(caminho_origem)[1] or ""
    return f"{nome_base}{ext}"


def espelhar_arquivo_em_saida_geral(
    caminho_origem: str,
    raiz_saida_geral: str,
    ano_alvo: int,
    mes_alvo: int,
    tipo_documento: str = "",
    competencia: str = "",
) -> str:
    codigo, slug = _identidade_saida_geral_atual()
    competencia_log = _competencia_saida_geral(ano_alvo=ano_alvo, mes_alvo=mes_alvo, competencia=competencia)
    empresa_log = codigo or slug or "EMPRESA_DESCONHECIDA"

    try:
        if not caminho_origem or not os.path.isfile(caminho_origem):
            raise FileNotFoundError(f"Arquivo original inexistente para espelhamento: {caminho_origem}")

        pasta_destino = pasta_saida_geral(
            raiz_saida_geral,
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            competencia=competencia,
        )
        os.makedirs(pasta_destino, exist_ok=True)

        nome_final = _nome_arquivo_saida_geral(caminho_origem, tipo_documento or raiz_saida_geral, ano_alvo, mes_alvo)
        destino_final = os.path.join(pasta_destino, nome_final)
        if os.path.exists(destino_final):
            base, ext = os.path.splitext(nome_final)
            k = 1
            while True:
                cand = os.path.join(pasta_destino, f"{base}_{k}{ext}")
                if not os.path.exists(cand):
                    destino_final = cand
                    break
                k += 1

        shutil.copy2(caminho_origem, destino_final)
        print(
            f"[SAIDAS_GERAIS] {raiz_saida_geral}=OK | EMPRESA={empresa_log} | COMPETENCIA={competencia_log or 'SEM_COMPETENCIA'} "
            f"| DESTINO={destino_final}"
        )
        return destino_final
    except Exception as exc:
        print(
            f"[SAIDAS_GERAIS] {raiz_saida_geral}=ERRO | EMPRESA={empresa_log} | COMPETENCIA={competencia_log or 'SEM_COMPETENCIA'} "
            f"| {type(exc).__name__}: {str(exc)[:180]}"
        )
        return ""

def organizar_xml_por_pasta(caminho_xml, info):
    """
    Move para:
      downloads/EMPRESA_PASTA/MM.AAAA/PRESTADOS/NFS_<NF>_<DD-MM-YYYY>.xml
    Data/competência vem da linha do site (fallback: dhProc).
    """
    parsed = parse_data_emissao_site(info.get("data_emissao", ""))
    if parsed:
        _, _, _, data_iso, competencia = parsed
    else:
        dhproc = extrair_dhproc_do_xml(caminho_xml)
        if dhproc and len(dhproc) >= 10:
            data_iso = dhproc[:10]
            ano, mes, _ = data_iso.split("-")
            competencia = f"{mes}.{ano}"
        else:
            data_iso, competencia = "SEM_DATA", "SEM_COMPETENCIA"

    info = enriquecer_info_com_cadastro(info)
    nf = (info.get("nf") or "").strip() or "SEM_NUMERO"
    destino_dir = pasta_servico("PRESTADOS", competencia=competencia)
    os.makedirs(destino_dir, exist_ok=True)

    if data_iso != "SEM_DATA":
        try:
            yyyy, mm, dd = data_iso.split("-")
            data_nome = f"{dd}-{mm}-{yyyy}"
        except Exception:
            data_nome = data_iso
        novo_nome = f"NFS_{nf}_{data_nome}.xml"
    else:
        novo_nome = f"NFS_{nf}.xml"
    novo_nome = prefixar_nome_arquivo_cadastro(novo_nome, cadastro=info)

    destino_final = os.path.join(destino_dir, novo_nome)

    # não sobrescrever
    if os.path.exists(destino_final):
        base, ext = os.path.splitext(novo_nome)
        k = 1
        while True:
            cand = os.path.join(destino_dir, f"{base}_{k}{ext}")
            if not os.path.exists(cand):
                destino_final = cand
                break
            k += 1

    mover_com_retry(caminho_xml, destino_final)
    return destino_final

# =====================
# PAGINACAO
# =====================
def page_size_atual():
    try:
        el = driver.find_element(By.ID, "gridListaPageSize")
        return int((el.get_attribute("value") or "0").strip() or "0")
    except Exception:
        return 0

def definir_page_size(max_por_pagina=100):
    alvo = min(max_por_pagina, 100)
    atual = page_size_atual()
    if atual == alvo:
        print(f"Page size ja esta em {alvo}.")
        return

    assinatura_anterior = assinatura_lista()
    inp = wait.until(EC.element_to_be_clickable((By.ID, "gridListaPageSize")))
    inp.click()
    try:
        inp.send_keys(Keys.CONTROL, "a")
    except Exception:
        inp.send_keys(Keys.COMMAND, "a")
    inp.send_keys(str(alvo))
    inp.send_keys(Keys.ENTER)

    try:
        esperar_troca_de_grid(assinatura_anterior, timeout=WAIT_LISTA_TIMEOUT)
    except TimeoutException:
        # fallback: ao menos garantir que a lista segue ativa
        esperar_lista(timeout=WAIT_LISTA_TIMEOUT)

    atual = page_size_atual()
    print(f"Page size configurado: {atual}")

def pagina_atual():
    try:
        val = driver.execute_script("return (document.getElementById('gridListaPage')||{}).value || '';")
        if str(val).strip().isdigit():
            return int(str(val).strip())
    except Exception:
        pass
    try:
        ativos = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'dataTables_paginate')]//li[contains(@class,'paginate_button') and contains(@class,'active')]//span",
        )
        for el in ativos:
            txt = (getattr(el, "text", "") or "").strip()
            if txt.isdigit():
                return int(txt)
    except Exception:
        pass
    return None


def _extrair_pagina_alvo_paginacao(onclick: str) -> int | None:
    m = re.search(r"mudarPagina,gridLista,(\d+)", onclick or "")
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def ir_para_proxima_pagina():
    global ULTIMO_CHECKBOX_MARCADO
    assinatura_anterior = assinatura_lista()
    pag_antes = pagina_atual()

    target_page = None
    botao_next = None

    try:
        botoes_next = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'dataTables_paginate')]"
            "//li[contains(@class,'next') and not(contains(@class,'disabled'))]"
            "//span[contains(@onclick,'mudarPagina,gridLista,')]",
        )
    except Exception:
        botoes_next = []

    for btn in botoes_next:
        alvo = _extrair_pagina_alvo_paginacao(btn.get_attribute("onclick") or "")
        if alvo is not None:
            botao_next = btn
            target_page = alvo
            break

    if botao_next is None:
        try:
            botoes_numericos = driver.find_elements(
                By.XPATH,
                "//div[contains(@class,'dataTables_paginate')]"
                "//span[contains(@onclick,'mudarPagina,gridLista,')]",
            )
        except Exception:
            botoes_numericos = []

        candidatos = []
        for btn in botoes_numericos:
            alvo = _extrair_pagina_alvo_paginacao(btn.get_attribute("onclick") or "")
            if alvo is None:
                continue
            if pag_antes is not None and alvo <= pag_antes:
                continue
            candidatos.append((alvo, btn))

        if candidatos:
            candidatos.sort(key=lambda item: item[0])
            target_page, botao_next = candidatos[0]

    if botao_next is None:
        print(f"Paginacao: nenhum botao de proxima pagina encontrado. Pagina atual={pag_antes}.")
        return False

    print(f"Paginacao: avançando de {pag_antes} para {target_page}.")
    click_robusto(botao_next)

    try:
        WebDriverWait(driver, WAIT_LISTA_TIMEOUT, poll_frequency=GRID_POLL_INTERVAL_SECONDS).until(
            lambda _d: (
                assinatura_lista() != assinatura_anterior
                or (
                    pagina_atual() is not None
                    and target_page is not None
                    and pagina_atual() == target_page
                )
                or (
                    pag_antes is not None
                    and pagina_atual() is not None
                    and pagina_atual() > pag_antes
                )
            )
        )
        ULTIMO_CHECKBOX_MARCADO = None
        return True
    except TimeoutException:
        pag_depois = pagina_atual()
        if pag_antes is not None and pag_depois is not None and pag_depois > pag_antes:
            ULTIMO_CHECKBOX_MARCADO = None
            return True
        print(
            f"Paginacao: clique executado, mas a pagina nao mudou. "
            f"Antes={pag_antes} Depois={pag_depois} Alvo={target_page}."
        )
        return False

# =====================
# PROCESSAR UMA NOTA (sem falso 502)
# =====================
# =====================
# SERVIÃ‡OS TOMADOS (DECLARAÃ‡ÃƒO FISCAL)
# =====================
def tomados_log_path_empresa(referencia: str = "", ano_alvo: int = 0, mes_alvo: int = 0):
    if referencia and not (ano_alvo and mes_alvo):
        ano_alvo, mes_alvo = parse_referencia_competencia(referencia)
    return caminho_arquivo_geral(TOMADOS_LOG_FILENAME, ano_alvo=ano_alvo, mes_alvo=mes_alvo)

def garantir_log_tomados_com_header(path_):
    """Agora garante um .txt simples para Tomados (uma linha por evento)."""
    os.makedirs(os.path.dirname(path_), exist_ok=True)
    if not os.path.exists(path_):
        with open(path_, "w", encoding="utf-8") as f:
            f.write("# Log Serviços Tomados (uma linha por evento)\n")

def salvar_log_tomados(
    status: str,
    referencia: str,
    arquivo_xml: str = "",
    arquivo_pdf: str = "",
    mensagem: str = "",
):
    ano_alvo, mes_alvo = parse_referencia_competencia(referencia)
    path_ = caminho_arquivo_geral(TOMADOS_LOG_FILENAME, ano_alvo=ano_alvo, mes_alvo=mes_alvo)
    garantir_log_tomados_com_header(path_)

    arquivo_base = os.path.basename(arquivo_xml) if arquivo_xml else ""
    pasta_xml = os.path.dirname(arquivo_xml) if arquivo_xml else ""
    arquivo_pdf_base = os.path.basename(arquivo_pdf) if arquivo_pdf else ""
    pasta_pdf = os.path.dirname(arquivo_pdf) if arquivo_pdf else ""

    parts = [
        f"STATUS={status}",
        f"REFERENCIA={referencia}",
    ]
    parts.extend(_partes_contexto_cadastro())
    if arquivo_base:
        parts.append(f"ARQ={arquivo_base}")
    if pasta_xml:
        parts.append(f"PASTA={pasta_xml}")
    if arquivo_pdf_base:
        parts.append(f"ARQ_PDF={arquivo_pdf_base}")
    if pasta_pdf:
        parts.append(f"PASTA_PDF={pasta_pdf}")
    if mensagem:
        parts.append(f"MSG={(mensagem or '').replace('|','/')}")

    append_txt(path_, " | ".join(parts))
    return path_

def _norm_ascii(s: str) -> str:
    s = s or ""
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)).lower()

def limpar_tmp_servicos_tomados():
    """Remove arquivos antigos de Serviçostomados.* no _tmp para garantir detecção de download novo."""
    try:
        for nome in os.listdir(TEMP_DOWNLOAD_DIR):
            n = _norm_ascii(nome)
            if n.startswith("servicostomados") and (n.endswith(".xml") or n.endswith(".crdownload")):
                try:
                    os.remove(os.path.join(TEMP_DOWNLOAD_DIR, nome))
                except Exception:
                    pass
    except Exception:
        pass

def dashboard_aberto() -> bool:
    try:
        return (
            len(driver.find_elements(By.ID, "imgdeclaracaofiscal")) > 0
            or len(driver.find_elements(By.ID, "divtxtdeclaracaofiscal")) > 0
        )
    except Exception:
        return False


def _achar_link_inicio(d):
    candidatos = d.find_elements(By.CSS_SELECTOR, "a.historic-item")
    for a in candidatos:
        try:
            txt = (a.text or "").strip()
            title = (a.get_attribute("title") or "").strip()
            if _norm_ascii(txt) == "inicio" or _norm_ascii(title) == "inicio":
                return a
        except Exception:
            continue
    els = d.find_elements(By.CSS_SELECTOR, "a.historic-item[title='Início']")
    return els[0] if els else False


def _clicar_breadcrumb_inicio(timeout=30):
    link_inicio = WebDriverWait(driver, timeout).until(_achar_link_inicio)
    click_robusto(link_inicio)


def _destino_contexto_atingido(destino: str) -> bool:
    destino_norm = (destino or "").strip().lower()
    if destino_norm == "dashboard":
        return dashboard_aberto()
    if destino_norm == "declaracao_fiscal":
        return tela_declaracao_fiscal_eletronica_aberta()
    raise ValueError(f"Destino de contexto não suportado: {destino}")


def _descricao_destino_contexto(destino: str) -> str:
    destino_norm = (destino or "").strip().lower()
    if destino_norm == "dashboard":
        return "Dashboard"
    if destino_norm == "declaracao_fiscal":
        return "Declaração Fiscal"
    return destino or "destino desconhecido"


def bootstrap_multicadastro_tardio(
    destino: str,
    log_path: str = "",
    contexto_evidencia: str = "CADASTROS_RELACIONADOS",
    ano_alvo: int = 0,
    mes_alvo: int = 0,
    ref_alvo: str = "",
) -> dict:
    if not tela_cadastros_relacionados_aberta():
        return {}

    cadastro = cadastro_multicadastro_tardio_atual()
    if cadastro:
        try:
            return ativar_cadastro_multicadastro_tardio(int((MULTI_CADASTRO_TARDIO_ESTADO or {}).get("indice_atual") or 0))
        except Exception:
            pass

    cadastros = coletar_cadastros_relacionados()
    if not cadastros:
        salvar_evidencia_html(
            contexto_evidencia,
            "RECUPERACAO_CADASTRO_FALHOU",
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            ref_alvo=ref_alvo,
            log_path=log_path,
            extra={
                "destino": _descricao_destino_contexto(destino),
                "erro": "GRID_NAO_IDENTIFICADA",
                "multi_cadastro_ativo": MULTI_CADASTRO_ATIVO,
            },
        )
        raise RuntimeError("Cadastros Relacionados reapareceu, mas a grade não pôde ser coletada para bootstrap tardio.")

    cadastro = inicializar_multicadastro_tardio(cadastros, indice_atual=0)
    destino_desc = _descricao_destino_contexto(destino)
    extra = {
        "destino": destino_desc,
        "cadastro": cadastro,
        "multi_cadastro_ativo": MULTI_CADASTRO_ATIVO,
        "bootstrap_tardio": True,
        "total_cadastros": len(cadastros),
    }
    if log_path:
        append_txt(
            log_path,
            f"{contexto_evidencia} | {ref_alvo} | CADASTROS_RELACIONADOS_DESCOBERTO_TARDIO | DESTINO={destino_desc} | TOTAL={len(cadastros)} | ORDEM={cadastro.get('ordem') or ''}",
        )
    salvar_evidencia_html(
        contexto_evidencia,
        "CADASTROS_RELACIONADOS_DESCOBERTO_TARDIO",
        ano_alvo=ano_alvo,
        mes_alvo=mes_alvo,
        ref_alvo=ref_alvo,
        log_path=log_path,
        extra=extra,
    )
    return cadastro


def _resolver_cadastro_para_recuperacao(
    destino: str,
    log_path: str = "",
    contexto_evidencia: str = "CADASTROS_RELACIONADOS",
    ano_alvo: int = 0,
    mes_alvo: int = 0,
    ref_alvo: str = "",
) -> dict:
    cadastro = contexto_cadastro_atual()
    if _cadastro_relacionado_tem_contexto(cadastro):
        return _sincronizar_contexto_cadastro_relacionado(cadastro)

    cadastro_tardio = cadastro_multicadastro_tardio_atual()
    if cadastro_tardio:
        try:
            return ativar_cadastro_multicadastro_tardio(int((MULTI_CADASTRO_TARDIO_ESTADO or {}).get("indice_atual") or 0))
        except Exception:
            pass

    if (destino or "").strip().lower() == "declaracao_fiscal":
        return bootstrap_multicadastro_tardio(
            destino,
            log_path=log_path,
            contexto_evidencia=contexto_evidencia,
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            ref_alvo=ref_alvo,
        )
    return {}


def recuperar_contexto_cadastro_relacionado(
    destino: str,
    timeout=40,
    tentativas: int = 2,
    abrir_destino=None,
    log_path: str = "",
    contexto_evidencia: str = "CADASTROS_RELACIONADOS",
    ano_alvo: int = 0,
    mes_alvo: int = 0,
    ref_alvo: str = "",
) -> bool:
    if _destino_contexto_atingido(destino):
        return True
    if not tela_cadastros_relacionados_aberta():
        return False

    cadastro = _resolver_cadastro_para_recuperacao(
        destino,
        log_path=log_path,
        contexto_evidencia=contexto_evidencia,
        ano_alvo=ano_alvo,
        mes_alvo=mes_alvo,
        ref_alvo=ref_alvo,
    )
    destino_desc = _descricao_destino_contexto(destino)
    extra_base = {
        "destino": destino_desc,
        "cadastro": cadastro,
        "multi_cadastro_ativo": MULTI_CADASTRO_ATIVO,
    }

    salvar_evidencia_html(
        contexto_evidencia,
        "CADASTROS_RELACIONADOS_REAPARECEU",
        ano_alvo=ano_alvo,
        mes_alvo=mes_alvo,
        ref_alvo=ref_alvo,
        log_path=log_path,
        extra=extra_base,
    )

    if not cadastro:
        motivo = f"Cadastros Relacionados reapareceu sem cadastro ativo para retomar {destino_desc}."
        if log_path:
            append_txt(log_path, f"{contexto_evidencia} | {ref_alvo} | RECUPERACAO_CADASTRO_FALHOU | {motivo}")
        salvar_evidencia_html(
            contexto_evidencia,
            "RECUPERACAO_CADASTRO_FALHOU",
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            ref_alvo=ref_alvo,
            log_path=log_path,
            extra={**extra_base, "erro": "CADASTRO_ATIVO_AUSENTE"},
        )
        raise RuntimeError(motivo)

    timeout_local = max(5, min(int(timeout), 20))
    ultimo_erro = "Cadastros Relacionados permaneceu aberto."
    for tentativa in range(1, max(1, tentativas) + 1):
        try:
            cadastro = _resolver_cadastro_para_recuperacao(
                destino,
                log_path=log_path,
                contexto_evidencia=contexto_evidencia,
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
            )
            if not cadastro:
                ultimo_erro = "Cadastro ativo ausente."
                continue
            extra_base["cadastro"] = cadastro

            if tela_cadastros_relacionados_aberta():
                cadastro = selecionar_cadastro_relacionado(cadastro, timeout=timeout_local)
                extra_base["cadastro"] = cadastro

            if _destino_contexto_atingido(destino):
                if log_path:
                    append_txt(log_path, f"{contexto_evidencia} | {ref_alvo} | RECUPERACAO_CADASTRO_OK | DESTINO={destino_desc} | TENTATIVA={tentativa}")
                return True

            if abrir_destino is not None and not tela_cadastros_relacionados_aberta():
                abrir_destino(timeout=timeout_local)

            WebDriverWait(driver, timeout_local).until(
                lambda _d: _destino_contexto_atingido(destino) or tela_cadastros_relacionados_aberta()
            )
            if _destino_contexto_atingido(destino):
                if log_path:
                    append_txt(log_path, f"{contexto_evidencia} | {ref_alvo} | RECUPERACAO_CADASTRO_OK | DESTINO={destino_desc} | TENTATIVA={tentativa}")
                return True

            ultimo_erro = "Cadastros Relacionados reapareceu novamente após a retomada."
        except Exception as exc:
            ultimo_erro = f"{type(exc).__name__}: {str(exc)[:220]}"

    salvar_evidencia_html(
        contexto_evidencia,
        "RECUPERACAO_CADASTRO_FALHOU",
        ano_alvo=ano_alvo,
        mes_alvo=mes_alvo,
        ref_alvo=ref_alvo,
        log_path=log_path,
        extra={**extra_base, "erro": ultimo_erro},
    )
    raise RuntimeError(
        f"Cadastros Relacionados reapareceu e não foi possível retomar {destino_desc}. {ultimo_erro}"
    )


def garantir_declaracao_fiscal_contexto(
    timeout=40,
    log_path: str = "",
    contexto_evidencia: str = "DECLARACAO_FISCAL",
    ano_alvo: int = 0,
    mes_alvo: int = 0,
    ref_alvo: str = "",
) -> bool:
    if tela_declaracao_fiscal_eletronica_aberta():
        return True

    try:
        WebDriverWait(driver, timeout).until(
            lambda _d: tela_declaracao_fiscal_eletronica_aberta() or tela_cadastros_relacionados_aberta()
        )
    except Exception:
        pass

    if tela_cadastros_relacionados_aberta():
        return recuperar_contexto_cadastro_relacionado(
            "declaracao_fiscal",
            timeout=timeout,
            tentativas=2,
            abrir_destino=abrir_declaracao_fiscal,
            log_path=log_path,
            contexto_evidencia=contexto_evidencia,
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            ref_alvo=ref_alvo,
        )

    if tela_declaracao_fiscal_eletronica_aberta():
        return True

    salvar_evidencia_html(
        contexto_evidencia,
        "RECUPERACAO_CADASTRO_FALHOU",
        ano_alvo=ano_alvo,
        mes_alvo=mes_alvo,
        ref_alvo=ref_alvo,
        log_path=log_path,
        extra={"erro": "DECLARACAO_FISCAL_NAO_ABERTA"},
    )
    raise RuntimeError("Declaração Fiscal não foi aberta e o contexto do cadastro não pôde ser confirmado.")


def clicar_inicio_para_dashboard(timeout=30):
    """Volta segura para o dashboard usando o breadcrumb 'Início'."""
    try:
        if dashboard_aberto():
            return
    except Exception:
        pass

    if tela_cadastros_relacionados_aberta():
        recuperar_contexto_cadastro_relacionado(
            "dashboard",
            timeout=timeout,
            tentativas=2,
            abrir_destino=_clicar_breadcrumb_inicio,
            contexto_evidencia="DASHBOARD",
        )
        WebDriverWait(driver, timeout).until(lambda _d: dashboard_aberto())
        return

    _clicar_breadcrumb_inicio(timeout=timeout)

    WebDriverWait(driver, timeout).until(
        lambda _d: dashboard_aberto() or tela_cadastros_relacionados_aberta()
    )

    if tela_cadastros_relacionados_aberta():
        recuperar_contexto_cadastro_relacionado(
            "dashboard",
            timeout=timeout,
            tentativas=2,
            abrir_destino=_clicar_breadcrumb_inicio,
            contexto_evidencia="DASHBOARD",
        )
        WebDriverWait(driver, timeout).until(lambda _d: dashboard_aberto())

def abrir_declaracao_fiscal(timeout=30):
    # Preferência: IMG (mais "único"), fallback: DIV
    if driver.find_elements(By.ID, "imgdeclaracaofiscal"):
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "imgdeclaracaofiscal")))
        click_robusto(el)
        return
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "divtxtdeclaracaofiscal")))
    click_robusto(el)

def abrir_servicos_tomados(timeout=30):
    if tela_cadastros_relacionados_aberta():
        raise RuntimeError("Cadastros Relacionados ainda está aberto antes de abrir Serviços Tomados.")
    if driver.find_elements(By.ID, "imgtomadiss"):
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "imgtomadiss")))
        click_robusto(el)
        return
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "divtxttomadiss")))
    click_robusto(el)


def abrir_servicos_prestados(timeout=30):
    # tile Serviços Prestados
    if tela_cadastros_relacionados_aberta():
        raise RuntimeError("Cadastros Relacionados ainda está aberto antes de abrir Serviços Prestados.")
    if driver.find_elements(By.ID, "imgprestiss"):
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "imgprestiss")))
        click_robusto(el)
        return
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "divtxtprestiss")))
    click_robusto(el)


def retornar_para_declaracao_fiscal(
    timeout: int = 40,
    log_path: str = "",
    contexto_evidencia: str = "DECLARACAO_FISCAL",
    ano_alvo: int = 0,
    mes_alvo: int = 0,
    ref_alvo: str = "",
) -> None:
    if tela_declaracao_fiscal_eletronica_aberta():
        return
    if driver is None:
        garantir_declaracao_fiscal_contexto(
            timeout=timeout,
            log_path=log_path,
            contexto_evidencia=contexto_evidencia,
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            ref_alvo=ref_alvo,
        )
        return

    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    try:
        if fechar_modal_livro_se_aberto(timeout=min(timeout, 5)):
            if log_path:
                append_txt(log_path, f"{contexto_evidencia} | {ref_alvo} | MODAL_RESIDUAL=FECHADO")
    except Exception:
        pass

    def _achar_breadcrumb_declaracao(d):
        candidatos = d.find_elements(By.CSS_SELECTOR, "a.historic-item")
        for el in candidatos:
            try:
                txt = _norm_txt((el.text or "").strip())
                title = _norm_txt((el.get_attribute("title") or "").strip())
                if "declaracao fiscal" in txt or "declaracao fiscal" in title:
                    return el
            except Exception:
                continue
        return False

    navegou = False
    try:
        link = WebDriverWait(driver, min(timeout, 8)).until(_achar_breadcrumb_declaracao)
        click_robusto(link)
        navegou = True
    except Exception:
        pass

    if not navegou:
        botoes_voltar = []
        try:
            botoes_voltar.extend(driver.find_elements(By.XPATH, "//button[contains(normalize-space(.),'Voltar')]"))
            botoes_voltar.extend(driver.find_elements(By.XPATH, "//a[contains(normalize-space(.),'Voltar')]"))
        except Exception:
            pass
        for btn in botoes_voltar:
            try:
                click_robusto(btn)
                navegou = True
                break
            except Exception:
                continue

    if not navegou and dashboard_aberto():
        abrir_declaracao_fiscal(timeout=min(timeout, 15))
        navegou = True

    if navegou:
        try:
            WebDriverWait(driver, timeout).until(
                lambda _d: tela_declaracao_fiscal_eletronica_aberta() or tela_cadastros_relacionados_aberta()
            )
        except Exception:
            pass

    garantir_declaracao_fiscal_contexto(
        timeout=timeout,
        log_path=log_path,
        contexto_evidencia=contexto_evidencia,
        ano_alvo=ano_alvo,
        mes_alvo=mes_alvo,
        ref_alvo=ref_alvo,
    )


def declaracao_fiscal_tem_servicos_prestados() -> bool:
    try:
        return (
            len(driver.find_elements(By.ID, "imgprestiss")) > 0
            or len(driver.find_elements(By.ID, "divtxtprestiss")) > 0
        )
    except Exception:
        return False

def esperar_grid_declaracao_fiscal(timeout=40) -> str:
    """Aguarda a grid da Declaração Fiscal e retorna o columnorder da coluna 'Referência'."""
    def _achar_th_referencia(d):
        ths = d.find_elements(By.XPATH, "//th[@columnorder]")
        for th in ths:
            try:
                txt = _norm_ascii((th.text or '').replace('\xa0', ' ').strip())
                if "referencia" in txt:
                    return th
            except Exception:
                continue
        return False

    th_ref = WebDriverWait(driver, timeout).until(_achar_th_referencia)
    col_ref = (th_ref.get_attribute("columnorder") or "7").strip() or "7"

    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, f"//td[contains(@id,',{col_ref}_grid')]"))
    )
    return col_ref

def tomados_grid_sem_registros(timeout=30) -> bool:
    """Detecta grid vazia na Declaração de serviços tomados.

    Considera 'sem tomados' quando:
      - tbody #_gridTBodyLinhas não tem <tr>
      - existe div .dataTables_empty com texto 'NENHUM REGISTRO...'
    Aguarda a grid 'decidir' (ou aparece linha, ou aparece empty state) para evitar falso positivo por carregamento.
    """
    state = {"val": None}

    def _decidiu(d):
        # Se tiver linha, não é vazio
        if d.find_elements(By.CSS_SELECTOR, "#_gridTBodyLinhas tr"):
            state["val"] = False
            return True

        empties = d.find_elements(By.CSS_SELECTOR, "#_gridTable .dataTables_empty")
        if empties:
            txt = _norm_ascii((empties[0].text or "").replace("\xa0", " ").strip())
            if "nenhum registro" in txt:
                # Confirma tbody realmente vazio
                if not d.find_elements(By.CSS_SELECTOR, "#_gridTBodyLinhas tr"):
                    state["val"] = True
                    return True
        return False

    WebDriverWait(driver, timeout).until(_decidiu)
    return bool(state["val"])



def encontrar_botao_laranja_por_referencia(ref_alvo: str, col_ref: str = "7"):
    """Encontra o botão laranja (fa-bars) da linha cuja Referência == ref_alvo.
    Travas:
      - célula de Referência da linha deve ser exatamente ref_alvo
      - botão deve estar na coluna 0 e ser fa-bars
      - título deve conter o ref_alvo e a palavra 'notas' (case-insensitive)
    """
    celulas_ref = driver.find_elements(By.XPATH, f"//td[contains(@id,',{col_ref}_grid')]")
    for cel in celulas_ref:
        try:
            if (cel.text or "").strip() != ref_alvo:
                continue
            tr = cel.find_element(By.XPATH, "./ancestor::tr[1]")

            # valida novamente referência na mesma linha
            try:
                cel_ref_linha = tr.find_element(By.XPATH, f".//td[contains(@id,',{col_ref}_grid')]")
                if (cel_ref_linha.text or '').strip() != ref_alvo:
                    continue
            except Exception:
                continue

            btn = tr.find_element(By.XPATH, ".//td[contains(@id,',0_grid')]//span[contains(@class,'fa-bars')]")
            title = (btn.get_attribute("title") or "").strip()
            tnorm = _norm_ascii(title)
            if (ref_alvo not in title) or ("notas" not in tnorm):
                continue
            return btn
        except Exception:
            continue
    return None


def encontrar_icone_guia_por_referencia(ref_alvo: str, col_ref: str = ""):
    """Encontra o ícone de impressão da guia ISS para a competência alvo."""
    ref = (ref_alvo or "").strip()
    if not ref:
        return None

    # 1) Melhor caminho: title do próprio ícone já contém "Imprimir Guia MM/AAAA".
    try:
        candidatos = driver.find_elements(
            By.XPATH,
            (
                "//span[contains(@class,'fa-print') "
                "and contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'imprimir guia') "
                f"and contains(@title,'{ref}')]"
            ),
        )
        for el in candidatos:
            try:
                if el.is_displayed():
                    return el
            except Exception:
                continue
        if candidatos:
            return candidatos[0]
    except Exception:
        pass

    # 2) Fallback: usa linha encontrada pela coluna de referência mapeada.
    try:
        if col_ref:
            tr = encontrar_tr_por_referencia_grid(ref, col_ref)
            if tr is not None:
                candidatos = tr.find_elements(
                    By.XPATH,
                    ".//span[contains(@class,'fa-print') and contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'imprimir guia')]",
                )
                for el in candidatos:
                    try:
                        if el.is_displayed():
                            return el
                    except Exception:
                        continue
                if candidatos:
                    return candidatos[0]
    except Exception:
        pass

    # 3) Fallback robusto: qualquer linha com célula igual à referência alvo.
    try:
        candidatos = driver.find_elements(
            By.XPATH,
            (
                f"//tr[.//td[normalize-space()='{ref}']]"
                "//span[contains(@class,'fa-print') "
                "and contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'imprimir guia')]"
            ),
        )
        for el in candidatos:
            try:
                if el.is_displayed():
                    return el
            except Exception:
                continue
        if candidatos:
            return candidatos[0]
    except Exception:
        pass

    return None


def _primeiro_visivel_no_contexto(xpath: str):
    for el in driver.find_elements(By.XPATH, xpath):
        try:
            if el.is_displayed():
                return el
        except Exception:
            continue
    return None


def esperar_controles_livro(timeout=20):
    """Aguarda os controles do livro no contexto atual (main/frame/iframe)."""
    def _achar(_):
        exercicio = _primeiro_visivel_no_contexto("//select[@id='exercicio' or @name='exercicio']")
        mes = _primeiro_visivel_no_contexto("//select[@id='mes' or @name='mes']")
        visualizar = _primeiro_visivel_no_contexto(
            "//button[@id='_imagebutton1' or contains(normalize-space(.),'Visualizar')]"
        )
        if exercicio and mes and visualizar:
            return {"exercicio": exercicio, "mes": mes, "visualizar": visualizar}
        return False

    return WebDriverWait(driver, timeout).until(_achar)


def _listar_iframes_ifilho_visiveis():
    candidatos = []
    for idx, frame in enumerate(driver.find_elements(By.TAG_NAME, "iframe")):
        try:
            fid = (frame.get_attribute("id") or "").strip()
            fname = (frame.get_attribute("name") or "").strip()
            if not frame.is_displayed():
                continue
            if "_iFilho" not in fid and "_iFilho" not in fname:
                continue
            candidatos.append({
                "idx": idx,
                "id": fid,
                "name": fname,
            })
        except Exception:
            continue
    candidatos.sort(key=lambda item: item["idx"], reverse=True)
    return candidatos


def entrar_contexto_livro(timeout=20):
    """Entra no contexto correto do livro: iframe externo e, se existir, frame interno 'inferior'."""
    fim = time.time() + timeout
    ultimo_erro = ""

    while time.time() < fim:
        driver.switch_to.default_content()
        iframe_candidatos = _listar_iframes_ifilho_visiveis()
        if not iframe_candidatos:
            sleep(0.20)
            continue

        for iframe_info in iframe_candidatos:
            try:
                outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
                if int(iframe_info["idx"]) >= len(outer_frames):
                    ultimo_erro = f"iframe-desapareceu-{iframe_info.get('idx')}"
                    continue
                driver.switch_to.frame(outer_frames[int(iframe_info["idx"])])

                try:
                    controles = esperar_controles_livro(timeout=2)
                    return {"iframe": iframe_info, "inner_frame": None, "controls": controles}
                except Exception as e:
                    ultimo_erro = f"iframe-{iframe_info.get('idx')}-sem-controles: {type(e).__name__}"

                inner_frames = driver.find_elements(By.TAG_NAME, "frame")
                candidatos = []
                for idx, frame in enumerate(inner_frames):
                    try:
                        candidatos.append({
                            "idx": idx,
                            "id": (frame.get_attribute("id") or "").strip(),
                            "name": (frame.get_attribute("name") or "").strip(),
                        })
                    except Exception:
                        candidatos.append({"idx": idx, "id": "", "name": ""})

                candidatos.sort(key=lambda item: (0 if item["id"] == "inferior" or item["name"] == "inferior" else 1, -item["idx"]))

                for inner in candidatos:
                    try:
                        driver.switch_to.default_content()
                        outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
                        if int(iframe_info["idx"]) >= len(outer_frames):
                            ultimo_erro = f"iframe-desapareceu-{iframe_info.get('idx')}"
                            continue
                        driver.switch_to.frame(outer_frames[int(iframe_info["idx"])])
                        inner_frames = driver.find_elements(By.TAG_NAME, "frame")
                        if int(inner["idx"]) >= len(inner_frames):
                            ultimo_erro = f"frame-desapareceu-{inner.get('idx')}"
                            continue
                        driver.switch_to.frame(inner_frames[int(inner["idx"])])
                        controles = esperar_controles_livro(timeout=2)
                        return {"iframe": iframe_info, "inner_frame": inner, "controls": controles}
                    except Exception as e:
                        ultimo_erro = f"iframe-{iframe_info.get('idx')}-frame-{inner.get('idx')}: {type(e).__name__}"
                        continue
            except Exception as e:
                ultimo_erro = f"iframe-{iframe_info.get('idx')}-erro: {type(e).__name__}"
            finally:
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass

        sleep(0.20)

    raise TimeoutException(f"Contexto do livro não encontrado. Ultimo erro: {ultimo_erro}")


def esperar_controles_guia(timeout=20):
    """Aguarda os controles da guia ISS no contexto atual (main/frame/iframe)."""
    def _achar(_):
        dtpagamento = _primeiro_visivel_no_contexto("//input[@id='dtpagamento' or @name='dtpagamento']")
        visualizar = _primeiro_visivel_no_contexto("//button[@id='btniwpppVisualizar' or contains(normalize-space(.),'Visualizar')]")
        cancelar = _primeiro_visivel_no_contexto("//button[@id='btniwpppCancelar' or contains(normalize-space(.),'Cancelar')]")
        if dtpagamento and visualizar:
            return {"dtpagamento": dtpagamento, "visualizar": visualizar, "cancelar": cancelar}
        return False

    return WebDriverWait(driver, timeout).until(_achar)


def entrar_contexto_guia(timeout=20):
    """Entra no contexto correto do modal de guia ISS: iframe externo e, se existir, frame interno."""
    fim = time.time() + timeout
    ultimo_erro = ""

    while time.time() < fim:
        driver.switch_to.default_content()
        iframe_candidatos = _listar_iframes_ifilho_visiveis()
        if not iframe_candidatos:
            sleep(0.20)
            continue

        for iframe_info in iframe_candidatos:
            try:
                outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
                if int(iframe_info["idx"]) >= len(outer_frames):
                    ultimo_erro = f"iframe-guia-desapareceu-{iframe_info.get('idx')}"
                    continue
                driver.switch_to.frame(outer_frames[int(iframe_info["idx"])])

                try:
                    controles = esperar_controles_guia(timeout=2)
                    return {"iframe": iframe_info, "inner_frame": None, "controls": controles}
                except Exception as e:
                    ultimo_erro = f"iframe-guia-{iframe_info.get('idx')}-sem-controles: {type(e).__name__}"

                inner_frames = driver.find_elements(By.TAG_NAME, "frame")
                candidatos = []
                for idx, frame in enumerate(inner_frames):
                    try:
                        candidatos.append({
                            "idx": idx,
                            "id": (frame.get_attribute("id") or "").strip(),
                            "name": (frame.get_attribute("name") or "").strip(),
                        })
                    except Exception:
                        candidatos.append({"idx": idx, "id": "", "name": ""})

                candidatos.sort(key=lambda item: (0 if item["id"] == "inferior" or item["name"] == "inferior" else 1, -item["idx"]))

                for inner in candidatos:
                    try:
                        driver.switch_to.default_content()
                        outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
                        if int(iframe_info["idx"]) >= len(outer_frames):
                            ultimo_erro = f"iframe-guia-desapareceu-{iframe_info.get('idx')}"
                            continue
                        driver.switch_to.frame(outer_frames[int(iframe_info["idx"])])
                        inner_frames = driver.find_elements(By.TAG_NAME, "frame")
                        if int(inner["idx"]) >= len(inner_frames):
                            ultimo_erro = f"frame-guia-desapareceu-{inner.get('idx')}"
                            continue
                        driver.switch_to.frame(inner_frames[int(inner["idx"])])
                        controles = esperar_controles_guia(timeout=2)
                        return {"iframe": iframe_info, "inner_frame": inner, "controls": controles}
                    except Exception as e:
                        ultimo_erro = f"iframe-guia-{iframe_info.get('idx')}-frame-{inner.get('idx')}: {type(e).__name__}"
                        continue
            except Exception as e:
                ultimo_erro = f"iframe-guia-{iframe_info.get('idx')}-erro: {type(e).__name__}"
            finally:
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass

        sleep(0.20)

    raise TimeoutException(f"Contexto da guia ISS não encontrado. Ultimo erro: {ultimo_erro}")


def _parse_data_br(txt: str):
    s = (txt or "").strip()
    try:
        return datetime.strptime(s, "%d/%m/%Y").date()
    except Exception:
        return None


def preparar_data_previsao_pagamento(modal, modulo_up: str, ref_alvo: str, log_manual: str):
    """Garante data de previsão válida (>= hoje) antes de clicar em Visualizar."""
    el = modal.get("dtpagamento")
    if el is None:
        return

    hoje = datetime.now().date()
    hoje_txt = hoje.strftime("%d/%m/%Y")
    atual_txt = (el.get_attribute("value") or "").strip()
    atual_dt = _parse_data_br(atual_txt)

    if atual_dt is not None and atual_dt >= hoje:
        driver.execute_script(
            """
            const el = arguments[0];
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new Event('blur', { bubbles: true }));
            """,
            el,
        )
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_DATA=OK | VALOR={atual_txt}")
        return

    driver.execute_script(
        """
        const el = arguments[0];
        const val = arguments[1];
        el.value = val;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
        """,
        el,
        hoje_txt,
    )

    WebDriverWait(driver, 5).until(lambda d: (el.get_attribute("value") or "").strip() == hoje_txt)
    append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_DATA=OK | VALOR={hoje_txt}")


def _achar_select_modal(controles, nome: str):
    if isinstance(controles, dict) and nome in controles:
        return controles[nome]
    raise RuntimeError(f"Select do modal não encontrado: {nome}")


def _disparar_eventos_select(el):
    driver.execute_script(
        """
        const el = arguments[0];
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
        """,
        el,
    )


def _selecionar_select_robusto(el, valor: str, texto_visivel: str = ""):
    valor = str(valor)
    texto_visivel = str(texto_visivel or valor)

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    WebDriverWait(driver, 10).until(lambda d: el.is_displayed() and el.is_enabled())

    if (el.get_attribute("value") or "").strip() == valor:
        _disparar_eventos_select(el)
        return

    ultimo_erro = None
    for _ in range(3):
        try:
            select = Select(el)
            try:
                select.select_by_value(valor)
            except Exception:
                select.select_by_visible_text(texto_visivel)
            _disparar_eventos_select(el)
            WebDriverWait(driver, 5).until(lambda d: (el.get_attribute("value") or "").strip() == valor)
            return
        except Exception as e:
            ultimo_erro = e
            try:
                driver.execute_script(
                    """
                    const el = arguments[0];
                    const valor = arguments[1];
                    el.value = valor;
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.dispatchEvent(new Event('blur', { bubbles: true }));
                    """,
                    el,
                    valor,
                )
                WebDriverWait(driver, 3).until(lambda d: (el.get_attribute("value") or "").strip() == valor)
                return
            except Exception as js_error:
                ultimo_erro = js_error
        sleep(DOWNLOAD_POLL_INTERVAL_SECONDS)

    raise RuntimeError(f"Falha ao selecionar valor '{valor}' no select: {ultimo_erro}")


def selecionar_competencia_modal_livro(modal, ano_alvo: int, mes_alvo: int, modulo_up: str, ref_alvo: str, log_manual: str):
    """Seleciona exercício e mês no modal do livro, validando os valores antes do Visualizar."""
    select_exercicio_el = _achar_select_modal(modal, "exercicio")
    valor_ano = str(ano_alvo)
    valor_exercicio_atual = (select_exercicio_el.get_attribute("value") or "").strip()
    if valor_exercicio_atual != valor_ano:
        _selecionar_select_robusto(select_exercicio_el, valor_ano, valor_ano)
    else:
        _disparar_eventos_select(select_exercicio_el)
    append_txt(log_manual, f"{modulo_up} | {ref_alvo} | EXERCICIO_LIVRO=OK | VALOR={select_exercicio_el.get_attribute('value')}")

    sleep(0.20)

    select_mes_el = _achar_select_modal(modal, "mes")
    valor_mes = str(int(mes_alvo))
    _selecionar_select_robusto(select_mes_el, valor_mes, valor_mes)
    append_txt(log_manual, f"{modulo_up} | {ref_alvo} | MES_LIVRO=OK | VALOR={select_mes_el.get_attribute('value')}")


def fechar_modal_livro_se_aberto(timeout=8) -> bool:
    """Fecha a janela/modal do livro (iframe _iFilho) se ainda estiver aberta."""
    def _buscar_candidatos():
        candidatos = []
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        for idx, fr in enumerate(frames):
            try:
                fid = (fr.get_attribute("id") or "").strip()
                fname = (fr.get_attribute("name") or "").strip()
                if "_iFilho" in fid or "_iFilho" in fname:
                    candidatos.append({"idx": idx, "id": fid, "name": fname})
            except Exception:
                continue
        return candidatos

    def _botoes_fechar_visiveis():
        xp = (
            "//button[contains(normalize-space(.),'Fechar') "
            "or contains(@onclick,'cancelar,executarMetodo') "
            "or contains(@class,'btn-default') and contains(.,'Fechar')]"
        )
        return driver.find_elements(By.XPATH, xp)

    try:
        driver.switch_to.default_content()
    except Exception:
        return False

    fim = time.time() + timeout
    fechou = False

    while time.time() < fim:
        candidatos = _buscar_candidatos()
        if not candidatos:
            return fechou

        tentou_algum = False
        for cand in candidatos:
            try:
                driver.switch_to.default_content()
                outer = driver.find_elements(By.TAG_NAME, "iframe")
                idx_outer = int(cand["idx"])
                if idx_outer >= len(outer):
                    continue
                driver.switch_to.frame(outer[idx_outer])

                botoes = _botoes_fechar_visiveis()
                if not botoes:
                    inner_frames = driver.find_elements(By.TAG_NAME, "frame")
                    for idx_inner in range(len(inner_frames)):
                        try:
                            driver.switch_to.default_content()
                            outer = driver.find_elements(By.TAG_NAME, "iframe")
                            if idx_outer >= len(outer):
                                break
                            driver.switch_to.frame(outer[idx_outer])
                            inner_frames = driver.find_elements(By.TAG_NAME, "frame")
                            if idx_inner >= len(inner_frames):
                                continue
                            driver.switch_to.frame(inner_frames[idx_inner])
                            botoes = _botoes_fechar_visiveis()
                            if botoes:
                                break
                        except Exception:
                            continue

                if not botoes:
                    continue

                btn = None
                for b in botoes:
                    try:
                        if b.is_displayed():
                            btn = b
                            break
                    except Exception:
                        continue
                if btn is None:
                    btn = botoes[0]

                tentou_algum = True
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                except Exception:
                    pass
                click_robusto(btn)
                fechou = True
                break
            except Exception:
                continue

        try:
            driver.switch_to.default_content()
        except Exception:
            pass

        if fechou:
            try:
                WebDriverWait(driver, 4).until(
                    lambda d: len([
                        f for f in d.find_elements(By.TAG_NAME, "iframe")
                        if (
                            ("_iFilho" in ((f.get_attribute("id") or "").strip()))
                            or ("_iFilho" in ((f.get_attribute("name") or "").strip()))
                        ) and f.is_displayed()
                    ]) == 0
                )
            except Exception:
                pass
            return True

        if not tentou_algum:
            sleep(0.2)

    return fechou


def abrir_modal_livro_e_visualizar(modulo: str, termo_titulo: str, ano_alvo: int, mes_alvo: int, ref_alvo: str) -> str:
    """Abre o modal do livro do módulo informado, seleciona a competência alvo e clica em Visualizar."""
    log_manual = caminho_log_manual(ref_alvo, ano_alvo, mes_alvo)
    modulo_up = (modulo or "").upper()
    handles_base = []
    try:
        driver.switch_to.default_content()
        handles_base = list(driver.window_handles)
        fechar_alerta_mensagens_prefeitura_se_aberto(timeout=3, log_path=log_manual, modulo=modulo_up, ref_alvo=ref_alvo)

        botao_livro = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.XPATH,
                f"//button[@type='button' and contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'livro') and contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'{termo_titulo.lower()}')]"
            ))
        )
        click_robusto(botao_livro)
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | BOTAO_LIVRO=OK | clicado")

        contexto = None
        for tentativa_contexto in range(2):
            try:
                contexto = entrar_contexto_livro(timeout=20 if tentativa_contexto == 0 else 10)
                break
            except Exception as exc_contexto:
                fechou_alerta = fechar_alerta_mensagens_prefeitura_se_aberto(
                    timeout=4,
                    log_path=log_manual,
                    modulo=modulo_up,
                    ref_alvo=ref_alvo,
                )
                if tentativa_contexto == 0 and fechou_alerta:
                    append_txt(log_manual, f"{modulo_up} | {ref_alvo} | ALERTA_MENSAGENS=RETRY_CONTEXTO_LIVRO")
                    continue
                salvar_evidencia_html(
                    f"LIVRO_{modulo_up}",
                    "CONTEXTO_LIVRO_NAO_ENCONTRADO",
                    ano_alvo=ano_alvo,
                    mes_alvo=mes_alvo,
                    ref_alvo=ref_alvo,
                    log_path=log_manual,
                    extra={
                        "modulo": modulo_up,
                        "termo_titulo": termo_titulo,
                        "erro": f"{type(exc_contexto).__name__}: {str(exc_contexto)[:300]}",
                        "tentativa_contexto": tentativa_contexto + 1,
                    },
                )
                raise
        iframe_info = contexto.get("iframe") or {}
        inner_info = contexto.get("inner_frame")
        append_txt(
            log_manual,
            f"{modulo_up} | {ref_alvo} | IFRAME_LIVRO=OK | IDX={iframe_info.get('idx','')} | ID={iframe_info.get('id','')} | NAME={iframe_info.get('name','')}"
        )
        if inner_info:
            append_txt(
                log_manual,
                f"{modulo_up} | {ref_alvo} | FRAME_LIVRO=OK | IDX={inner_info.get('idx','')} | ID={inner_info.get('id','')} | NAME={inner_info.get('name','')}"
            )

        # Reentra no mesmo contexto encontrado para manter os controles estáveis.
        driver.switch_to.default_content()
        outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
        driver.switch_to.frame(outer_frames[int(iframe_info["idx"])])
        if inner_info:
            inner_frames = driver.find_elements(By.TAG_NAME, "frame")
            driver.switch_to.frame(inner_frames[int(inner_info["idx"])])

        modal = esperar_controles_livro(timeout=5)
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | MODAL_LIVRO=OK | aberto")

        selecionar_competencia_modal_livro(modal, ano_alvo, mes_alvo, modulo_up, ref_alvo, log_manual)
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | COMPETENCIA_LIVRO=OK | EXERCICIO={ano_alvo} | MES={mes_alvo:02d}")

        limpar_tmp_livros_pdf()
        pdfs_antes = [a for a in os.listdir(TEMP_DOWNLOAD_DIR) if a.lower().endswith(".pdf")]
        mm_aaaa = f"{mes_alvo:02d}-{ano_alvo}"
        nome_fallback_pdf = (
            LIVRO_TOMADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)
            if modulo_up == "TOMADOS"
            else LIVRO_PRESTADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)
        )
        visualizar_btn = modal["visualizar"]
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", visualizar_btn)
        WebDriverWait(driver, 5).until(lambda d: visualizar_btn.is_displayed() and visualizar_btn.is_enabled())

        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | VISUALIZAR=OK | EXERCICIO={ano_alvo} | MES={mes_alvo:02d}")

        pdf_baixado = ""
        try:
            pdf_baixado = baixar_arquivo_via_submit_form(
                visualizar_btn,
                timeout=120,
                nome_fallback=nome_fallback_pdf,
                tipo_esperado="pdf",
                debug_dump_tag=f"submit_form_pdf_livro_{modulo_up.lower()}",
                debug_info={"modulo": modulo_up, "tipo": "livro"},
                debug_ref_alvo=ref_alvo,
            )
            append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_FETCH=OK | via submit_form")
        except Exception as exc_submit:
            append_txt(
                log_manual,
                f"{modulo_up} | {ref_alvo} | PDF_SUBMIT_FETCH_FAIL | {type(exc_submit).__name__}: {str(exc_submit)[:140]}",
            )

        if not pdf_baixado:
            try:
                modal = esperar_controles_livro(timeout=5)
                visualizar_btn = modal["visualizar"]
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", visualizar_btn)
            except Exception:
                pass
            try:
                popup_info = {}
                pdf_baixado, popup_info = baixar_pdf_via_popup_ou_logs(
                    visualizar_btn,
                    handles_base,
                    timeout=120,
                    nome_fallback=nome_fallback_pdf,
                    aceitar_alerta=True,
                )
                pdf_url = popup_info.get("url") or ""
                origem_popup = popup_info.get("source") or "popup_url"
                if pdf_url:
                    append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_POPUP_URL=OK | {pdf_url}")
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_FETCH=OK | via {origem_popup}")
            except Exception as exc_popup:
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_POPUP_FAIL | {type(exc_popup).__name__}: {str(exc_popup)[:140]}")

            if not pdf_baixado:
                try:
                    pdf_baixado = aguardar_pdf_novo(TEMP_DOWNLOAD_DIR, timeout=45, pdfs_antes=pdfs_antes)
                    append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_FETCH=OK | via arquivo _tmp")
                except Exception:
                    pass

        if not pdf_baixado:
            try:
                fluxo_pdf = capturar_fluxo_pdf_via_logs(timeout=20)
                pdf_baixado, origem_rede = materializar_pdf_de_fluxo_rede(
                    fluxo_pdf,
                    nome_fallback=nome_fallback_pdf,
                    timeout=120,
                )
                if pdf_baixado:
                    append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_FETCH=OK | via {origem_rede}")
            except Exception as e:
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_NET_FETCH_FAIL | {type(e).__name__}: {str(e)[:140]}")

        # Se já baixou pelo arquivo direto, NÃO fazer varredura de URLs/janelas.
        # Isso evita "piscar" de abas e acelera o encerramento do fluxo.
        if not pdf_baixado:
            urls_pdf = coletar_urls_pdf_pos_visualizar(handles_base, timeout=12)
            if urls_pdf:
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_URLS=OK | QTD={len(urls_pdf)}")
                urls_prioritarias = [u for u in urls_pdf if (".pdf" in u.lower() or "/resultados/" in u.lower())]
                urls_tentativa = (urls_prioritarias or urls_pdf)[:8]
                for url_cand in urls_tentativa:
                    try:
                        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_URL=TRY | {url_cand}")
                        pdf_baixado = baixar_pdf_por_url_com_cookies(url_cand, timeout=120)
                        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_FETCH=OK | via URL+cookies")
                        break
                    except Exception as e:
                        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_URL=FAIL | {type(e).__name__}: {str(e)[:90]}")
            else:
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_URLS=VAZIO")

        if not pdf_baixado:
            if LIVRO_PDF_FALLBACK_CHROME:
                pdf_baixado = aguardar_pdf_novo(TEMP_DOWNLOAD_DIR, timeout=120, pdfs_antes=pdfs_antes)
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_FETCH=OK | via download Chrome")
            else:
                raise RuntimeError("Nao foi possivel baixar PDF via URL+cookies e fallback Chrome esta desativado")

        destino_pdf = salvar_pdf_livro(pdf_baixado, modulo_up, ano_alvo, mes_alvo)
        append_txt(
            log_manual,
            f"{modulo_up} | {ref_alvo} | PDF_LIVRO=OK | ARQ={os.path.basename(destino_pdf)} | PASTA={os.path.dirname(destino_pdf)}"
        )
        return "OK"
    except Exception as e:
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | ERRO_MODAL_LIVRO | {type(e).__name__}: {str(e)[:180]}")
        raise
    finally:
        try:
            fechar_janelas_extras(handles_base)
        except Exception:
            pass
        try:
            if fechar_modal_livro_se_aberto(timeout=8):
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | MODAL_LIVRO=FECHADO")
        except Exception:
            pass
        try:
            driver.switch_to.default_content()
        except Exception:
            pass

def salvar_xml_tomados(xml_tmp_path: str, ano_alvo: int, mes_alvo: int) -> str:
    competencia_dir = pasta_servico("TOMADOS", ano_alvo=ano_alvo, mes_alvo=mes_alvo)
    os.makedirs(competencia_dir, exist_ok=True)

    mm_aaaa = f"{mes_alvo:02d}-{ano_alvo}"
    nome_final = TOMADOS_XML_BASENAME.format(mm_aaaa=mm_aaaa)
    nome_final = prefixar_nome_arquivo_cadastro(nome_final)
    destino_final = os.path.join(competencia_dir, nome_final)

    if os.path.exists(destino_final):
        base, ext = os.path.splitext(nome_final)
        k = 1
        while True:
            cand = os.path.join(competencia_dir, f"{base}_{k}{ext}")
            if not os.path.exists(cand):
                destino_final = cand
                break
            k += 1

    mover_com_retry(xml_tmp_path, destino_final)
    espelhar_arquivo_em_saida_geral(destino_final, "XML_TOMADOS", ano_alvo, mes_alvo, "SERVICOS_TOMADOS")
    return destino_final


def salvar_pdf_tomados_nota(pdf_tmp_path: str, info: dict, ano_alvo: int, mes_alvo: int) -> str:
    info = enriquecer_info_com_cadastro(info)
    competencia_dir = pasta_servico("TOMADOS", ano_alvo=ano_alvo, mes_alvo=mes_alvo)
    os.makedirs(competencia_dir, exist_ok=True)

    nf = re.sub(r"\D", "", info.get("nf") or "").strip() or "NF"
    parsed = parse_data_emissao_site(info.get("data_emissao", ""))
    if parsed:
        yyyy, mm, dd, _, _ = parsed
        data_nome = f"{dd}-{mm}-{yyyy}"
    else:
        data_nome = f"{mes_alvo:02d}-{ano_alvo}"

    nome_final = TOMADOS_NFSE_PDF_BASENAME.format(nf=nf, data=data_nome)
    nome_final = prefixar_nome_arquivo_cadastro(nome_final, cadastro=info)
    destino_final = os.path.join(competencia_dir, nome_final)

    if os.path.exists(destino_final):
        base, ext = os.path.splitext(nome_final)
        k = 1
        while True:
            cand = os.path.join(competencia_dir, f"{base}_{k}{ext}")
            if not os.path.exists(cand):
                destino_final = cand
                break
            k += 1

    mover_com_retry(pdf_tmp_path, destino_final)
    return destino_final


def salvar_pdf_livro(pdf_tmp_path: str, modulo_up: str, ano_alvo: int, mes_alvo: int) -> str:
    competencia_dir = pasta_servico(modulo_up, ano_alvo=ano_alvo, mes_alvo=mes_alvo)
    os.makedirs(competencia_dir, exist_ok=True)

    mm_aaaa = f"{mes_alvo:02d}-{ano_alvo}"
    if (modulo_up or "").upper() == "TOMADOS":
        nome_final = LIVRO_TOMADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)
    else:
        nome_final = LIVRO_PRESTADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)
    nome_final = prefixar_nome_arquivo_cadastro(nome_final)

    destino_final = os.path.join(competencia_dir, nome_final)

    if os.path.exists(destino_final):
        base, ext = os.path.splitext(nome_final)
        k = 1
        while True:
            cand = os.path.join(competencia_dir, f"{base}_{k}{ext}")
            if not os.path.exists(cand):
                destino_final = cand
                break
            k += 1

    mover_com_retry(pdf_tmp_path, destino_final)
    return destino_final


def salvar_pdf_guia(pdf_tmp_path: str, modulo_up: str, ano_alvo: int, mes_alvo: int) -> str:
    competencia_dir = pasta_servico(modulo_up, ano_alvo=ano_alvo, mes_alvo=mes_alvo)
    os.makedirs(competencia_dir, exist_ok=True)

    mm_aaaa = f"{mes_alvo:02d}-{ano_alvo}"
    if (modulo_up or "").upper() == "TOMADOS":
        nome_final = GUIA_ISS_TOMADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)
    else:
        nome_final = GUIA_ISS_PRESTADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)
    nome_final = prefixar_nome_arquivo_cadastro(nome_final)

    destino_final = os.path.join(competencia_dir, nome_final)

    if os.path.exists(destino_final):
        base, ext = os.path.splitext(nome_final)
        k = 1
        while True:
            cand = os.path.join(competencia_dir, f"{base}_{k}{ext}")
            if not os.path.exists(cand):
                destino_final = cand
                break
            k += 1

    mover_com_retry(pdf_tmp_path, destino_final)
    espelhar_arquivo_em_saida_geral(destino_final, "ISS", ano_alvo, mes_alvo)
    return destino_final


def baixar_guia_iss_se_existir(modulo_up: str, ref_alvo: str, col_ref: str, ano_alvo: int, mes_alvo: int) -> str:
    """Se existir ícone de guia ISS na linha da referência, abre modal e baixa o PDF."""
    log_manual = caminho_log_manual(ref_alvo, ano_alvo, mes_alvo)
    handles_base = []
    try:
        driver.switch_to.default_content()
        handles_base = list(driver.window_handles)
        icone_guia = encontrar_icone_guia_por_referencia(ref_alvo, col_ref)
        if icone_guia is None:
            qtd_print_total = 0
            qtd_print_ref = 0
            try:
                qtd_print_total = len(driver.find_elements(By.XPATH, "//span[contains(@class,'fa-print')]"))
            except Exception:
                pass
            try:
                qtd_print_ref = len(driver.find_elements(By.XPATH, f"//span[contains(@class,'fa-print') and contains(@title,'{ref_alvo}')]"))
            except Exception:
                pass
            append_txt(
                log_manual,
                f"{modulo_up} | {ref_alvo} | GUIA_ISS=NAO_DISPONIVEL | PRINT_TOTAL={qtd_print_total} | PRINT_REF={qtd_print_ref}",
            )
            return "NAO_DISPONIVEL"

        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", icone_guia)
        except Exception:
            pass
        click_robusto(icone_guia)
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_ICONE=OK | clicado")

        contexto = entrar_contexto_guia(timeout=20)
        iframe_info = contexto.get("iframe") or {}
        inner_info = contexto.get("inner_frame")
        append_txt(
            log_manual,
            f"{modulo_up} | {ref_alvo} | GUIA_IFRAME=OK | IDX={iframe_info.get('idx','')} | ID={iframe_info.get('id','')} | NAME={iframe_info.get('name','')}"
        )
        if inner_info:
            append_txt(
                log_manual,
                f"{modulo_up} | {ref_alvo} | GUIA_FRAME=OK | IDX={inner_info.get('idx','')} | ID={inner_info.get('id','')} | NAME={inner_info.get('name','')}"
            )

        driver.switch_to.default_content()
        outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
        driver.switch_to.frame(outer_frames[int(iframe_info["idx"])])
        if inner_info:
            inner_frames = driver.find_elements(By.TAG_NAME, "frame")
            driver.switch_to.frame(inner_frames[int(inner_info["idx"])])

        modal = esperar_controles_guia(timeout=5)
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_MODAL=OK | aberto")
        preparar_data_previsao_pagamento(modal, modulo_up, ref_alvo, log_manual)

        limpar_tmp_livros_pdf()
        pdfs_antes = [a for a in os.listdir(TEMP_DOWNLOAD_DIR) if a.lower().endswith(".pdf")]
        mm_aaaa = f"{mes_alvo:02d}-{ano_alvo}"
        nome_fallback_pdf = (
            GUIA_ISS_TOMADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)
            if modulo_up == "TOMADOS"
            else GUIA_ISS_PRESTADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)
        )

        visualizar_btn = modal["visualizar"]
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", visualizar_btn)
        WebDriverWait(driver, 5).until(lambda d: visualizar_btn.is_displayed() and visualizar_btn.is_enabled())
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_VISUALIZAR=OK")

        pdf_baixado = ""
        try:
            pdf_baixado = baixar_arquivo_via_submit_form(
                visualizar_btn,
                timeout=120,
                nome_fallback=nome_fallback_pdf,
                tipo_esperado="pdf",
                debug_dump_tag=f"submit_form_pdf_guia_{modulo_up.lower()}",
                debug_info={"modulo": modulo_up, "tipo": "guia"},
                debug_ref_alvo=ref_alvo,
            )
            append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_FETCH=OK | via submit_form")
        except Exception as exc_submit:
            append_txt(
                log_manual,
                f"{modulo_up} | {ref_alvo} | GUIA_PDF_SUBMIT_FAIL | {type(exc_submit).__name__}: {str(exc_submit)[:140]}",
            )

        if not pdf_baixado:
            try:
                modal = esperar_controles_guia(timeout=5)
                visualizar_btn = modal["visualizar"]
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", visualizar_btn)
            except Exception:
                pass
            try:
                popup_info = {}
                pdf_baixado, popup_info = baixar_pdf_via_popup_ou_logs(
                    visualizar_btn,
                    handles_base,
                    timeout=120,
                    nome_fallback=nome_fallback_pdf,
                    aceitar_alerta=True,
                )
                pdf_url = popup_info.get("url") or ""
                origem_popup = popup_info.get("source") or "popup_url"
                if pdf_url:
                    append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_POPUP_URL=OK | {pdf_url}")
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_FETCH=OK | via {origem_popup}")
            except Exception as exc_popup:
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_POPUP_FAIL | {type(exc_popup).__name__}: {str(exc_popup)[:140]}")

            if not pdf_baixado:
                try:
                    pdf_baixado = aguardar_pdf_novo(TEMP_DOWNLOAD_DIR, timeout=45, pdfs_antes=pdfs_antes)
                    append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_FETCH=OK | via arquivo _tmp")
                except Exception:
                    pass

        if not pdf_baixado:
            try:
                fluxo_pdf = capturar_fluxo_pdf_via_logs(timeout=20)
                pdf_baixado, origem_rede = materializar_pdf_de_fluxo_rede(
                    fluxo_pdf,
                    nome_fallback=nome_fallback_pdf,
                    timeout=120,
                )
                if pdf_baixado:
                    append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_FETCH=OK | via {origem_rede}")
            except Exception as e:
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_NET_FAIL | {type(e).__name__}: {str(e)[:140]}")

        if not pdf_baixado:
            urls_pdf = coletar_urls_pdf_pos_visualizar(handles_base, timeout=12)
            if urls_pdf:
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_URLS=OK | QTD={len(urls_pdf)}")
                urls_prioritarias = [u for u in urls_pdf if (".pdf" in u.lower() or "/resultados/" in u.lower())]
                urls_tentativa = (urls_prioritarias or urls_pdf)[:8]
                for url_cand in urls_tentativa:
                    try:
                        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_URL=TRY | {url_cand}")
                        pdf_baixado = baixar_pdf_por_url_com_cookies(url_cand, timeout=120)
                        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_FETCH=OK | via URL+cookies")
                        break
                    except Exception as e:
                        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_URL=FAIL | {type(e).__name__}: {str(e)[:90]}")
            else:
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_URLS=VAZIO")

        if not pdf_baixado:
            if LIVRO_PDF_FALLBACK_CHROME:
                pdf_baixado = aguardar_pdf_novo(TEMP_DOWNLOAD_DIR, timeout=120, pdfs_antes=pdfs_antes)
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_FETCH=OK | via download Chrome")
            else:
                raise RuntimeError("Nao foi possivel baixar PDF da guia via URL+cookies e fallback Chrome esta desativado")

        destino_pdf = salvar_pdf_guia(pdf_baixado, modulo_up, ano_alvo, mes_alvo)
        append_txt(
            log_manual,
            f"{modulo_up} | {ref_alvo} | GUIA_PDF=OK | ARQ={os.path.basename(destino_pdf)} | PASTA={os.path.dirname(destino_pdf)}"
        )
        return "OK"
    except Exception as e:
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_ISS=ERRO | {type(e).__name__}: {str(e)[:180]}")
        return "ERRO"
    finally:
        try:
            if handles_base:
                fechar_janelas_extras(handles_base)
        except Exception:
            pass
        try:
            if fechar_modal_livro_se_aberto(timeout=8):
                append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_MODAL=FECHADO")
        except Exception:
            pass
        try:
            driver.switch_to.default_content()
        except Exception:
            pass


def executar_etapa_servicos_tomados(ano_alvo: int, mes_alvo: int) -> str:
    """Etapa adicional: Declaração Fiscal -> Serviços Tomados -> exportar XML do mês alvo.

    Retorna: "OK" | "SUCESSO_SEM_TOMADOS" | "ERRO"
    Se TOMADOS_OBRIGATORIO=1 e ocorrer ERRO, levanta RuntimeError(MSG_TOMADOS_FALHA).
    """
    ref_alvo = f"{mes_alvo:02d}/{ano_alvo}"
    log_manual = caminho_log_manual(ref_alvo, ano_alvo, mes_alvo)

    try:
        salvar_log_tomados("INICIO", ref_alvo, mensagem="Iniciando etapa Serviços Tomados")
        limpar_status_tomados_pdf(referencia=ref_alvo)
    except Exception:
        pass

    try:
        if not _fluxo_multicadastro_tardio_local_ativo():
            clicar_inicio_para_dashboard(timeout=40)
            abrir_declaracao_fiscal(timeout=40)
        elif dashboard_aberto():
            abrir_declaracao_fiscal(timeout=40)

        garantir_declaracao_fiscal_contexto(
            timeout=40,
            log_path=log_manual,
            contexto_evidencia="DECLARACAO_FISCAL_TOMADOS",
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            ref_alvo=ref_alvo,
        )
        abrir_servicos_tomados(timeout=40)
        fechar_alerta_mensagens_prefeitura_se_aberto(timeout=5, log_path=log_manual, modulo="TOMADOS", ref_alvo=ref_alvo)

        # Aguarda a grid existir (pode estar vazia; não exige linhas)
        WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.ID, "_gridTable"))
        )

        # Se a grid estiver vazia (NENHUM REGISTRO...), já encerramos como sucesso sem tomados.
        if tomados_grid_sem_registros(timeout=25):
            salvar_print_evidencia(
                "SERVICOS_TOMADOS",
                "SEM_IDENTIFICACAO_OU_SEM_COMPETENCIA",
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
                log_path=log_manual,
            )
            salvar_log_tomados("SUCESSO_SEM_TOMADOS", ref_alvo, mensagem="Nenhum registro para apresentação")
            print("[TOMADOS] Nenhum registro para apresentação. SUCESSO_SEM_TOMADOS.")
            return "SUCESSO_SEM_TOMADOS"

        # Agora que sabemos que há linhas, mapeamos a coluna "Referência"
        col_ref = esperar_grid_declaracao_fiscal(timeout=50)

        garantir_contexto_referencia_grid(
            ref_alvo,
            col_ref,
            timeout=15,
            log_path=log_manual,
            modulo="TOMADOS",
        )

        btn_laranja = encontrar_botao_laranja_por_referencia(ref_alvo, col_ref=col_ref)
        if not btn_laranja:
            salvar_print_evidencia(
                "SERVICOS_TOMADOS",
                "COMPETENCIA_NAO_ENCONTRADA",
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
                log_path=log_manual,
            )
            salvar_log_tomados("SUCESSO_SEM_TOMADOS", ref_alvo, mensagem="Referência não encontrada na Declaração Fiscal")
            print(f"[TOMADOS] Referência {ref_alvo} não encontrada. SUCESSO_SEM_TOMADOS.")
            return "SUCESSO_SEM_TOMADOS"

        click_robusto(btn_laranja)
        fechar_alerta_mensagens_prefeitura_se_aberto(timeout=5, log_path=log_manual, modulo="TOMADOS", ref_alvo=ref_alvo)

        # Checkpoint crítico: referência no topo precisa bater
        # Checkpoint crítico: referência no topo precisa bater (normalizado)
        ref_topo = normalizar_referencia_competencia(obter_referencia_topo_declaracao(timeout=50))
        if ref_topo != normalizar_referencia_competencia(ref_alvo):
            salvar_evidencia_html(
                "SERVICOS_TOMADOS",
                "REFERENCIA_TOPO_DIVERGENTE",
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
                log_path=log_manual,
                extra={
                    "referencia_esperada": ref_alvo,
                    "referencia_topo": ref_topo,
                },
            )
            raise RuntimeError(f"Referência no topo não confere. Esperado {ref_alvo}, veio {ref_topo}")

        # Checkpoint: controles essenciais
        # Checkpoint: controles essenciais (NÃO exigir botão clicável antes de selecionar)
        try:
            chk_all = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "_gridListaAll")))
        except TimeoutException:
            chk_all = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//input[@type='checkbox' and contains(@onclick,'marcarTodos') and contains(@onclick,'gridLista')]"
                ))
            )

        # O botão de exportação pode ficar desabilitado até haver seleção; aguarde presença primeiro
        try:
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "_imagebutton4")))
        except TimeoutException:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//button[contains(@title,'arquivo XML') and contains(@title,'Exportar')]"
                ))
            )

        limpar_tmp_servicos_tomados()
        xmls_antes = [a for a in os.listdir(TEMP_DOWNLOAD_DIR) if a.lower().endswith(".xml")]

        # Marcar todos e validar
        if not chk_all.is_selected():
            click_robusto(chk_all)
        WebDriverWait(driver, 20).until(lambda d: chk_all.is_selected())

        # Baixar XML (primário: submit+fetch na sessão, sem depender do download do Chrome)
        try:
            btn_download = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "_imagebutton4")))
        except TimeoutException:
            btn_download = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[contains(@title,'arquivo XML') and contains(@title,'Exportar')]"
                ))
            )
        try:
            xml_baixado = baixar_arquivo_via_submit_form(
                btn_download,
                timeout=120,
                nome_fallback=f"SERVICOS_TOMADOS_{mes_alvo:02d}-{ano_alvo}.xml",
            )
        except Exception as e_fetch:
            append_txt(
                log_manual,
                f"TOMADOS | {ref_alvo} | XML_FETCH_FALLBACK_CHROME | {type(e_fetch).__name__}: {str(e_fetch)[:120]}",
            )
            click_robusto(btn_download)
            xml_baixado = aguardar_xml_novo(TEMP_DOWNLOAD_DIR, timeout=120, xmls_antes=xmls_antes)

        destino_final = salvar_xml_tomados(xml_baixado, ano_alvo, mes_alvo)

        salvar_log_tomados("OK", ref_alvo, arquivo_xml=destino_final, mensagem="XML de Serviços Tomados baixado e organizado")
        print(f"[TOMADOS] OK -> {destino_final}")

        total_pdfs = baixar_pdfs_servicos_tomados(ano_alvo, mes_alvo, ref_alvo)
        limpar_status_tomados_pdf(referencia=ref_alvo)
        print(f"[TOMADOS] PDFs OK -> {total_pdfs}")
        return "OK"

    except Exception as e:
        msg = str(e)
        try:
            salvar_log_tomados("ERRO", ref_alvo, mensagem=msg[:200])
        except Exception:
            pass
        print(f"[TOMADOS] ERRO: {msg}")
        if MSG_TOMADOS_PDF_REVISAO_MANUAL in msg:
            raise RuntimeError(msg)
        if TOMADOS_OBRIGATORIO:
            raise RuntimeError(f"{MSG_TOMADOS_FALHA}: {msg}")
        return "ERRO"


def executar_etapa_fiscal_por_cadastro(
    cadastro: dict,
    ano_alvo: int,
    mes_alvo: int,
    *,
    executar_tomados_xml: bool,
    executar_cadeado_tomados_flag: bool,
    executar_cadeado_prestados_flag: bool,
) -> dict:
    cadastro = _sincronizar_contexto_cadastro_relacionado(cadastro)
    etapas = {}
    falhas = []
    ref_alvo = f"{mes_alvo:02d}/{ano_alvo}"
    log_manual = caminho_log_manual(ref_alvo, ano_alvo, mes_alvo)

    def _retornar_para_declaracao(etapa: str, contexto_evidencia: str) -> bool:
        try:
            retornar_para_declaracao_fiscal(
                timeout=40,
                log_path=log_manual,
                contexto_evidencia=contexto_evidencia,
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
            )
            return True
        except Exception as exc:
            falhas.append(f"{etapa}_RETORNO_DECLARACAO={type(exc).__name__}: {str(exc)[:140]}")
            return False

    interromper_fluxo = False
    _definir_fluxo_multicadastro_tardio_local(True)
    try:
        if executar_tomados_xml:
            try:
                status_tomados = executar_etapa_servicos_tomados(ano_alvo, mes_alvo)
                etapas["tomados_xml"] = status_tomados
                if status_tomados == "ERRO":
                    falhas.append("TOMADOS_XML=ERRO")
            except Exception as exc:
                etapas["tomados_xml"] = "ERRO"
                falhas.append(f"TOMADOS_XML={type(exc).__name__}: {str(exc)[:140]}")
            interromper_fluxo = not _retornar_para_declaracao("TOMADOS_XML", "DECLARACAO_FISCAL_TOMADOS")

        if executar_cadeado_tomados_flag and not interromper_fluxo:
            try:
                status_cadeado_tomados = executar_cadeado_tomados(ano_alvo, mes_alvo)
                etapas["cadeado_tomados"] = status_cadeado_tomados
                if status_cadeado_tomados == "ERRO":
                    falhas.append("CADEADO_TOMADOS=ERRO")
            except Exception as exc:
                etapas["cadeado_tomados"] = "ERRO"
                falhas.append(f"CADEADO_TOMADOS={type(exc).__name__}: {str(exc)[:140]}")
            interromper_fluxo = not _retornar_para_declaracao("CADEADO_TOMADOS", "DECLARACAO_FISCAL_TOMADOS")

        if executar_cadeado_prestados_flag and not interromper_fluxo:
            try:
                status_cadeado_prestados = executar_cadeado_prestados(ano_alvo, mes_alvo)
                etapas["cadeado_prestados"] = status_cadeado_prestados
                if status_cadeado_prestados == "ERRO":
                    falhas.append("CADEADO_PRESTADOS=ERRO")
            except Exception as exc:
                etapas["cadeado_prestados"] = "ERRO"
                falhas.append(f"CADEADO_PRESTADOS={type(exc).__name__}: {str(exc)[:140]}")
            _retornar_para_declaracao("CADEADO_PRESTADOS", "DECLARACAO_FISCAL_PRESTADOS")
    finally:
        _definir_fluxo_multicadastro_tardio_local(False)

    return {
        "cadastro": dict(contexto_cadastro_atual() or cadastro),
        "status": "FALHA" if falhas else "SUCESSO",
        "motivo": " ; ".join(falhas[:4]) if falhas else "OK",
        "etapas": etapas,
    }


def executar_fluxo_fiscal_multicadastro_tardio_se_necessario(
    ano_alvo: int,
    mes_alvo: int,
    executar_tomados_xml: bool,
    executar_cadeado_tomados_flag: bool,
    executar_cadeado_prestados_flag: bool,
    executar_pausa_final: bool = False,
) -> dict | None:
    if not any([executar_tomados_xml, executar_cadeado_tomados_flag, executar_cadeado_prestados_flag]):
        return None

    ref_alvo = f"{mes_alvo:02d}/{ano_alvo}"
    log_manual = caminho_log_manual(ref_alvo, ano_alvo, mes_alvo)

    clicar_inicio_para_dashboard(timeout=40)
    abrir_declaracao_fiscal(timeout=40)
    garantir_declaracao_fiscal_contexto(
        timeout=40,
        log_path=log_manual,
        contexto_evidencia="DECLARACAO_FISCAL_TARDIO",
        ano_alvo=ano_alvo,
        mes_alvo=mes_alvo,
        ref_alvo=ref_alvo,
    )

    if not multicadastro_tardio_ativo():
        return None

    estado = estado_multicadastro_tardio()
    cadastros = [dict(item) for item in (estado.get("cadastros") or [])]
    if not cadastros:
        return None

    try:
        indice_inicial = int(estado.get("indice_atual") or 0)
    except Exception:
        indice_inicial = 0
    indice_inicial = max(0, min(indice_inicial, len(cadastros) - 1))

    resultados = []
    append_txt(log_manual, f"MULTI_CADASTRO_TARDIO | {ref_alvo} | INICIO | TOTAL={len(cadastros)} | INICIAL={indice_inicial + 1}")

    for idx in range(indice_inicial, len(cadastros)):
        cadastro = dict(cadastros[idx])
        try:
            cadastro = ativar_cadastro_multicadastro_tardio(idx)
        except Exception as exc:
            item = {
                "cadastro": dict(cadastro),
                "status": "FALHA",
                "motivo": f"Falha ao ativar cadastro tardio: {type(exc).__name__}: {str(exc)[:160]}",
                "etapas": {},
            }
            resultados.append(item)
            registrar_resultado_multicadastro_tardio(item)
            continue

        if idx > indice_inicial:
            try:
                retornar_para_cadastros_relacionados(timeout=40)
                cadastro = selecionar_cadastro_relacionado(cadastro, timeout=40)
            except Exception as exc:
                item = {
                    "cadastro": dict(cadastro),
                    "status": "FALHA",
                    "motivo": f"Falha ao selecionar cadastro tardio: {type(exc).__name__}: {str(exc)[:160]}",
                    "etapas": {},
                }
                resultados.append(item)
                registrar_resultado_multicadastro_tardio(item)
                continue

        prefixo = _prefixo_cadastro_base(cadastro) or f"CAD_{int(cadastro.get('ordem') or 0):02d}"
        append_txt(log_manual, f"MULTI_CADASTRO_TARDIO | {ref_alvo} | CADASTRO_INICIO | {prefixo}")

        item = executar_etapa_fiscal_por_cadastro(
            cadastro,
            ano_alvo,
            mes_alvo,
            executar_tomados_xml=executar_tomados_xml,
            executar_cadeado_tomados_flag=executar_cadeado_tomados_flag,
            executar_cadeado_prestados_flag=executar_cadeado_prestados_flag,
        )
        resultados.append(item)
        registrar_resultado_multicadastro_tardio(item)
        append_txt(
            log_manual,
            f"MULTI_CADASTRO_TARDIO | {ref_alvo} | CADASTRO_FIM | {prefixo} | STATUS={item['status']} | MOTIVO={item['motivo']}",
        )

    if executar_pausa_final:
        pausa_final_livros_manual(f"Competência alvo: {ref_alvo}")

    status_final, motivo_final = agregar_resultados_multicadastro(resultados)
    append_txt(log_manual, f"MULTI_CADASTRO_TARDIO | {ref_alvo} | RESUMO | STATUS={status_final} | MOTIVO={motivo_final}")
    return {
        "status": status_final,
        "motivo": motivo_final,
        "resultados": resultados,
    }


def executar_cadeado_tomados(ano_alvo: int, mes_alvo: int) -> str:
    """Confirma o fechamento e abre o modal do livro em Serviços Tomados."""
    ref_alvo = f"{mes_alvo:02d}/{ano_alvo}"
    log_manual = caminho_log_manual(ref_alvo, ano_alvo, mes_alvo)
    try:
        append_txt(log_manual, f"TOMADOS | {ref_alvo} | INICIO | Cadeado + Modal Livro")
        if not _fluxo_multicadastro_tardio_local_ativo():
            clicar_inicio_para_dashboard(timeout=40)
            abrir_declaracao_fiscal(timeout=40)
        elif dashboard_aberto():
            abrir_declaracao_fiscal(timeout=40)
        garantir_declaracao_fiscal_contexto(
            timeout=40,
            log_path=log_manual,
            contexto_evidencia="DECLARACAO_FISCAL_TOMADOS",
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            ref_alvo=ref_alvo,
        )
        abrir_servicos_tomados(timeout=40)
        fechar_alerta_mensagens_prefeitura_se_aberto(timeout=5, log_path=log_manual, modulo="TOMADOS", ref_alvo=ref_alvo)
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@id,'gridTable')]")))

        # Se não houver linhas, não há o que fechar
        if tomados_grid_sem_registros(timeout=25):
            salvar_print_evidencia(
                "SERVICOS_TOMADOS",
                "SEM_IDENTIFICACAO_OU_SEM_COMPETENCIA",
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
                log_path=log_manual,
            )
            append_txt(log_manual, f"TOMADOS | {ref_alvo} | SEM_MOVIMENTO (sem linhas)")
            return "SEM_MOVIMENTO"

        col_ref = esperar_grid_declaracao_fiscal(timeout=50)
        tr = encontrar_tr_por_referencia_grid(ref_alvo, col_ref)
        if not tr:
            salvar_print_evidencia(
                "SERVICOS_TOMADOS",
                "COMPETENCIA_NAO_ENCONTRADA",
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
                log_path=log_manual,
            )
            append_txt(log_manual, f"TOMADOS | {ref_alvo} | SEM_MOVIMENTO (referência não encontrada)")
            return "SEM_MOVIMENTO"

        garantir_contexto_referencia_grid(
            ref_alvo,
            col_ref,
            timeout=15,
            log_path=log_manual,
            modulo="TOMADOS",
        )

        clicou = clicar_cadeado_na_linha(ref_alvo, col_ref)
        if clicou:
            append_txt(log_manual, f"TOMADOS | {ref_alvo} | CADEADO=OK")
        else:
            append_txt(log_manual, f"TOMADOS | {ref_alvo} | SEM_CADEADO | Cadeado não disponível (provável já fechado)")

        if EXECUTAR_LIVROS:
            abrir_modal_livro_e_visualizar("TOMADOS", "tomados", ano_alvo, mes_alvo, ref_alvo)
        else:
            append_txt(log_manual, f"TOMADOS | {ref_alvo} | LIVRO=SKIP_PERFIL")

        if EXECUTAR_ISS:
            baixar_guia_iss_se_existir("TOMADOS", ref_alvo, col_ref, ano_alvo, mes_alvo)
        else:
            append_txt(log_manual, f"TOMADOS | {ref_alvo} | GUIA_ISS=SKIP_PERFIL")

        if PAUSAR_ENTRE_MODULOS_CADEADO:
            pausa_manual(f"TOMADOS {ref_alvo}")
        return "OK"
    except Exception as e:
        append_txt(
            log_manual,
            f"TOMADOS | {ref_alvo} | ERRO_CADEADO | {type(e).__name__}: {str(e)[:180]}"
        )
        if PAUSAR_ENTRE_MODULOS_CADEADO:
            pausa_manual(f"TOMADOS {ref_alvo} (erro no cadeado)")
        return "ERRO"


def executar_cadeado_prestados(ano_alvo: int, mes_alvo: int) -> str:
    """Confirma o fechamento e abre o modal do livro em Serviços Prestados."""
    ref_alvo = f"{mes_alvo:02d}/{ano_alvo}"
    log_manual = caminho_log_manual(ref_alvo, ano_alvo, mes_alvo)
    try:
        if PRESTADOS_SEM_MODULO:
            append_txt(log_manual, f"PRESTADOS | {ref_alvo} | SKIP_SEM_MODULO_NOTA_FISCAL")
            return "SKIP_SEM_MODULO_NOTA_FISCAL"

        append_txt(log_manual, f"PRESTADOS | {ref_alvo} | INICIO | Cadeado + Modal Livro")
        try:
            driver.switch_to.default_content()
            if fechar_modal_livro_se_aberto(timeout=3):
                append_txt(log_manual, f"PRESTADOS | {ref_alvo} | MODAL_RESIDUAL=FECHADO")
        except Exception:
            pass
        if not _fluxo_multicadastro_tardio_local_ativo():
            clicar_inicio_para_dashboard(timeout=40)
            abrir_declaracao_fiscal(timeout=40)
        elif dashboard_aberto():
            abrir_declaracao_fiscal(timeout=40)
        garantir_declaracao_fiscal_contexto(
            timeout=40,
            log_path=log_manual,
            contexto_evidencia="DECLARACAO_FISCAL_PRESTADOS",
            ano_alvo=ano_alvo,
            mes_alvo=mes_alvo,
            ref_alvo=ref_alvo,
        )
        if not declaracao_fiscal_tem_servicos_prestados():
            salvar_print_evidencia(
                "DECLARACAO_FISCAL_PRESTADOS",
                "SERVICOS_PRESTADOS_NAO_IDENTIFICADO",
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
                log_path=log_manual,
            )
            append_txt(log_manual, f"PRESTADOS | {ref_alvo} | SKIP_SEM_TILE_DECLARACAO_FISCAL")
            return "SKIP_SEM_TILE_DECLARACAO_FISCAL"
        abrir_servicos_prestados(timeout=40)
        fechar_alerta_mensagens_prefeitura_se_aberto(timeout=5, log_path=log_manual, modulo="PRESTADOS", ref_alvo=ref_alvo)
        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@id,'gridTable')]")))

        # Reusa detector (funciona igual na prática, mas é tolerante)
        if tomados_grid_sem_registros(timeout=25):
            salvar_print_evidencia(
                "DECLARACAO_FISCAL_PRESTADOS",
                "SEM_COMPETENCIA_PRESTADOS",
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
                log_path=log_manual,
            )
            append_txt(log_manual, f"PRESTADOS | {ref_alvo} | SEM_MOVIMENTO (sem linhas)")
            return "SEM_MOVIMENTO"

        col_ref = esperar_grid_declaracao_fiscal(timeout=50)
        tr = encontrar_tr_por_referencia_grid(ref_alvo, col_ref)
        if not tr:
            salvar_print_evidencia(
                "DECLARACAO_FISCAL_PRESTADOS",
                "COMPETENCIA_NAO_ENCONTRADA_PRESTADOS",
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
                log_path=log_manual,
            )
            append_txt(log_manual, f"PRESTADOS | {ref_alvo} | SEM_MOVIMENTO (referência não encontrada)")
            return "SEM_MOVIMENTO"

        garantir_contexto_referencia_grid(
            ref_alvo,
            col_ref,
            timeout=15,
            log_path=log_manual,
            modulo="PRESTADOS",
        )

        clicou = clicar_cadeado_na_linha(ref_alvo, col_ref)
        if clicou:
            append_txt(log_manual, f"PRESTADOS | {ref_alvo} | CADEADO=OK")
        else:
            append_txt(log_manual, f"PRESTADOS | {ref_alvo} | SEM_CADEADO | Cadeado não disponível (provável já fechado)")

        if EXECUTAR_LIVROS:
            abrir_modal_livro_e_visualizar("PRESTADOS", "prestados", ano_alvo, mes_alvo, ref_alvo)
        else:
            append_txt(log_manual, f"PRESTADOS | {ref_alvo} | LIVRO=SKIP_PERFIL")

        if EXECUTAR_ISS:
            baixar_guia_iss_se_existir("PRESTADOS", ref_alvo, col_ref, ano_alvo, mes_alvo)
        else:
            append_txt(log_manual, f"PRESTADOS | {ref_alvo} | GUIA_ISS=SKIP_PERFIL")

        if PAUSAR_ENTRE_MODULOS_CADEADO:
            pausa_manual(f"PRESTADOS {ref_alvo}")
        return "OK"
    except Exception as e:
        append_txt(
            log_manual,
            f"PRESTADOS | {ref_alvo} | ERRO_CADEADO | {type(e).__name__}: {str(e)[:180]}"
        )
        if PAUSAR_ENTRE_MODULOS_CADEADO:
            pausa_manual(f"PRESTADOS {ref_alvo} (erro no cadeado)")
        return "ERRO"


def _obter_checkbox_prestados_por_indice(i: int, checkboxes_cache=None):
    checkboxes = list(checkboxes_cache or [])
    if not checkboxes:
        checkboxes = obter_checkboxes_lista(timeout=0)
    if not checkboxes:
        checkboxes = obter_checkboxes_lista(timeout=min(WAIT_LISTA_TIMEOUT, 5))
    if i >= len(checkboxes):
        return None
    return checkboxes[i]


def _obter_botao_exportar_xml_prestados():
    try:
        btn = driver.find_element(By.ID, "_imagebutton12")
        if btn.is_displayed() and btn.is_enabled():
            return btn
    except Exception:
        pass
    return wait.until(EC.element_to_be_clickable((By.ID, "_imagebutton12")))


def _obter_modal_exportacao_xml_prestados():
    try:
        modal = driver.find_element(By.ID, "modalboxexportarnotas")
        if modal.is_displayed():
            return modal
    except Exception:
        pass
    return wait.until(EC.visibility_of_element_located((By.ID, "modalboxexportarnotas")))


def _obter_menu_mais_acoes_prestados():
    try:
        btn = driver.find_element(By.ID, "_navigatorListabt12")
        if btn.is_displayed() and btn.is_enabled():
            return btn
    except Exception:
        pass
    return wait.until(EC.element_to_be_clickable((By.ID, "_navigatorListabt12")))


def _obter_acao_exportar_xml_nacional_prestados():
    xpath = (
        "//a[@gridrelacionado='gridLista' "
        "and contains(normalize-space(.), 'Exportar arquivo XML Nacional')]"
    )
    try:
        return driver.find_element(By.XPATH, xpath)
    except Exception:
        return wait.until(EC.presence_of_element_located((By.XPATH, xpath)))


def _baixar_xml_prestados_via_menu_exportar(info: dict, i: int) -> tuple[str, str]:
    nome_fallback = f"{(info.get('nf') or 'NF')}.xml"
    acao_exportar = _obter_acao_exportar_xml_nacional_prestados()

    try:
        xml_baixado = baixar_arquivo_via_submit_form(
            acao_exportar,
            timeout=120,
            nome_fallback=nome_fallback,
        )
        return xml_baixado, "menu_submit"
    except Exception as erro_submit:
        xmls_antes = [a for a in os.listdir(TEMP_DOWNLOAD_DIR) if a.lower().endswith(".xml")]
        erro_click = None

        try:
            click_robusto(_obter_menu_mais_acoes_prestados())
            acao_exportar = _obter_acao_exportar_xml_nacional_prestados()
            click_robusto(acao_exportar)

            if fechar_aviso_se_existir(timeout=1):
                checkbox = _obter_checkbox_prestados_por_indice(i)
                if checkbox is None:
                    raise RuntimeError("Checkbox da nota nao encontrado apos fechar aviso de selecao.")
                marcar_somente_checkbox(checkbox, full_reset=True)
                click_robusto(_obter_menu_mais_acoes_prestados())
                acao_exportar = _obter_acao_exportar_xml_nacional_prestados()
                click_robusto(acao_exportar)

            xml_baixado = aguardar_xml_novo(TEMP_DOWNLOAD_DIR, timeout=60, xmls_antes=xmls_antes)
            return xml_baixado, "menu_click"
        except Exception as exc_click:
            erro_click = exc_click

        raise RuntimeError(
            "Fluxo XML Nacional via menu falhou. "
            f"submit={type(erro_submit).__name__}: {str(erro_submit)[:120]} | "
            f"click={type(erro_click).__name__}: {str(erro_click)[:120]}"
        )


def _baixar_xml_prestados_via_modal_legado(info: dict, i: int) -> tuple[str, str]:
    nome_fallback = f"{(info.get('nf') or 'NF')}.xml"
    xmls_antes = [a for a in os.listdir(TEMP_DOWNLOAD_DIR) if a.lower().endswith(".xml")]

    botao_xml = _obter_botao_exportar_xml_prestados()
    click_robusto(botao_xml)

    if fechar_aviso_se_existir(timeout=1):
        checkbox = _obter_checkbox_prestados_por_indice(i)
        if checkbox is None:
            raise RuntimeError("Checkbox da nota nao encontrado apos fechar aviso de selecao.")
        marcar_somente_checkbox(checkbox, full_reset=True)
        botao_xml = _obter_botao_exportar_xml_prestados()
        click_robusto(botao_xml)

    modal = _obter_modal_exportacao_xml_prestados()
    select = Select(modal.find_element(By.TAG_NAME, "select"))
    try:
        selected_text = (select.first_selected_option.text or "").strip().lower()
    except Exception:
        selected_text = ""
    if "nf nacional" not in selected_text:
        select.select_by_visible_text("NF Nacional")
        try:
            WebDriverWait(
                driver,
                max(1, int(math.ceil(SELECT_OPTION_SETTLE_TIMEOUT))),
                poll_frequency=GRID_POLL_INTERVAL_SECONDS,
            ).until(
                lambda d: "nf nacional" in (select.first_selected_option.text or "").strip().lower()
            )
        except Exception:
            pass

    visualizar_btn = modal.find_element(By.XPATH, ".//button[contains(.,'Visualizar')]")
    try:
        xml_baixado = baixar_arquivo_via_submit_form(
            visualizar_btn,
            timeout=120,
            nome_fallback=nome_fallback,
        )
        return xml_baixado, "modal_submit"
    except Exception:
        click_robusto(visualizar_btn)
        xml_baixado = aguardar_xml_novo(TEMP_DOWNLOAD_DIR, timeout=80, xmls_antes=xmls_antes)
        return xml_baixado, "modal_click"


def _processar_download_xml_prestados(checkbox, info: dict, i: int, tentativa: int, inicio_perf: float) -> str:
    key = chave_unica(info)

    if key in chaves_ok:
        print(f"[{i+1}] SKIP -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
        perf_log("prestados.nota_skip_cache", inicio_perf, f"indice={i+1}")
        return "SKIP_CACHE"

    if nota_cancelada(info):
        msg = f"Nota cancelada (situacao: {info.get('situacao', '')})"
        print(f"[{i+1}] SKIP CANCELADA -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
        try:
            salvar_log("", info, status="SKIP_CANCELADA", mensagem=msg[:180])
        except Exception:
            pass
        perf_log("prestados.nota_skip_cancelada", inicio_perf, f"indice={i+1}")
        return "SKIP_CANCELADA"

    print(f"[{i+1}] Tentativa {tentativa}/{MAX_RETRIES_POR_NOTA} -> NF={info.get('nf')} RPS={info.get('rps')} DATA={info.get('data_emissao')}")

    try:
        xml_baixado, origem_download = _baixar_xml_prestados_via_menu_exportar(info, i)
    except Exception as erro_menu:
        print(
            f"[{i+1}] Aviso: fluxo via menu XML Nacional falhou, aplicando fallback legado. "
            f"{type(erro_menu).__name__}: {str(erro_menu)[:180]}"
        )
        xml_baixado, origem_download = _baixar_xml_prestados_via_modal_legado(info, i)

    destino_final = organizar_xml_por_pasta(xml_baixado, info)
    logp = salvar_log(
        destino_final,
        info,
        status="OK",
        mensagem=f"OK (download confirmado por arquivo via {origem_download})",
    )
    chaves_ok.add(key)

    print(f"OK -> {destino_final}")
    print(f"Log -> {logp}")

    fechar_modal_exportacao()
    perf_log(
        "prestados.nota_ok",
        inicio_perf,
        f"indice={i+1} nf={info.get('nf') or ''} tentativa={tentativa} origem={origem_download}",
    )
    return "OK"


def processar_nota_por_indice(i, ano_alvo, mes_alvo, checkboxes_cache=None):
    global PARAR_PROCESSAMENTO, ENCONTROU_MES_ALVO, CONT_FORA_APOS_ALVO, CONT_FORA_ANTES_ALVO, SEM_COMPETENCIA_NA_EMPRESA

    tentativa = 0
    while tentativa < MAX_RETRIES_POR_NOTA:
        tentativa += 1
        info = {}
        inicio_perf = time.perf_counter()

        try:
            if is_502_page():
                recover_from_502()

            checkbox = _obter_checkbox_prestados_por_indice(i, checkboxes_cache if tentativa == 1 else None)
            if checkbox is None:
                perf_log("prestados.nota_skip_eof", inicio_perf, f"indice={i+1}")
                return True

            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", checkbox)
            if GRID_SCROLL_SETTLE_SECONDS > 0:
                sleep(GRID_SCROLL_SETTLE_SECONDS)
            marcar_somente_checkbox(checkbox)

            info = extrair_info_linha(checkbox)

            # Avalia competência primeiro para encerrar cedo em ordem decrescente.
            comp = comparar_competencia_nota(info, ano_alvo, mes_alvo)
            if comp is None:
                msg = f"Data de emissao invalida/ausente: {info.get('data_emissao', '')}"
                print(f"[{i+1}] SKIP DATA INVALIDA -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                try:
                    salvar_log("", info, status="SKIP_DATA_INVALIDA", mensagem=msg[:180])
                except Exception:
                    pass
                perf_log("prestados.nota_skip_data_invalida", inicio_perf, f"indice={i+1}")
                return True

            if comp == 1:
                CONT_FORA_ANTES_ALVO = 0
                msg = f"Nota mais nova que mês alvo {mes_alvo:02d}/{ano_alvo}"
                print(f"[{i+1}] SKIP COMPETENCIA (MAIS NOVA) -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                try:
                    salvar_log("", info, status="SKIP_FORA_COMPETENCIA", mensagem=msg[:180])
                except Exception:
                    pass
                perf_log("prestados.nota_skip_competencia_nova", inicio_perf, f"indice={i+1}")
                return True

            if comp == -1:
                msg = f"Nota mais antiga que mês alvo {mes_alvo:02d}/{ano_alvo}"
                print(f"[{i+1}] SKIP COMPETENCIA (MAIS ANTIGA) -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                try:
                    salvar_log("", info, status="SKIP_FORA_COMPETENCIA", mensagem=msg[:180])
                except Exception:
                    pass

                # Lista costuma estar em ordem decrescente. Ao encontrar a primeira nota mais antiga:
                # - se já encontramos notas do mês alvo, podemos encerrar a varredura (sucesso, sem marcar sem-competência)
                # - se ainda NÃO encontramos mês alvo, então a empresa não tem notas na competência alvo
                if ENCONTROU_MES_ALVO:
                    print("Primeira nota mais antiga que a competência alvo encontrada após achar o mês alvo. Encerrando varredura da empresa.")
                    PARAR_PROCESSAMENTO = True
                else:
                    print("Primeira nota mais antiga que a competência alvo encontrada antes de achar o mês alvo. Encerrando empresa como sem competência.")
                    SEM_COMPETENCIA_NA_EMPRESA = True
                    PARAR_PROCESSAMENTO = True
                perf_log("prestados.nota_skip_competencia_antiga", inicio_perf, f"indice={i+1}")
                return True

            # comp == 0 (mês alvo)
            ENCONTROU_MES_ALVO = True
            CONT_FORA_APOS_ALVO = 0
            CONT_FORA_ANTES_ALVO = 0

            _processar_download_xml_prestados(checkbox, info, i, tentativa, inicio_perf)
            return True

        except TimeoutException as e:
            # Timeout geralmente = não abriu modal / não baixou a tempo
            msg = f"Timeout: {str(e)}"
            print(f"Erro nota [{i+1}] tentativa {tentativa}/{MAX_RETRIES_POR_NOTA}: {msg}")

            try:
                salvar_log("", info, status="ERRO", mensagem=msg[:180])
            except Exception:
                pass

            if is_502_page():
                recover_from_502()

            fechar_modal_exportacao()
            sleep(min(2 * tentativa, 6))
            perf_log("prestados.nota_timeout", inicio_perf, f"indice={i+1} tentativa={tentativa}")

        except (StaleElementReferenceException, WebDriverException) as e:
            msg = str(e)
            print(f"Erro nota [{i+1}] tentativa {tentativa}/{MAX_RETRIES_POR_NOTA}: {msg}")

            try:
                salvar_log("", info, status="ERRO", mensagem=msg[:180])
            except Exception:
                pass

            if is_502_page():
                recover_from_502()

            fechar_modal_exportacao()
            sleep(min(2 * tentativa, 6))
            perf_log("prestados.nota_webdriver", inicio_perf, f"indice={i+1} tentativa={tentativa}")

        except RuntimeError:
            raise

        except Exception as e:
            msg = str(e)
            print(f"Erro inesperado nota [{i+1}] tentativa {tentativa}/{MAX_RETRIES_POR_NOTA}: {msg}")

            try:
                salvar_log("", info, status="ERRO", mensagem=("Inesperado: " + msg)[:180])
            except Exception:
                pass

            if is_502_page():
                recover_from_502()

            fechar_modal_exportacao()
            sleep(min(2 * tentativa, 6))
            perf_log("prestados.nota_erro", inicio_perf, f"indice={i+1} tentativa={tentativa}")

    print(f"Falhou apos {MAX_RETRIES_POR_NOTA} tentativas. Pulando nota {i+1}.")
    return False


def processar_nota_por_indice_apuracao_completa(i, ano_alvo, checkboxes_cache=None):
    tentativa = 0
    while tentativa < MAX_RETRIES_POR_NOTA:
        tentativa += 1
        info = {}
        inicio_perf = time.perf_counter()

        try:
            if is_502_page():
                recover_from_502()

            checkbox = _obter_checkbox_prestados_por_indice(i, checkboxes_cache if tentativa == 1 else None)
            if checkbox is None:
                perf_log("prestados.apuracao_completa_skip_eof", inicio_perf, f"indice={i+1}")
                return "EOF", None

            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", checkbox)
            if GRID_SCROLL_SETTLE_SECONDS > 0:
                sleep(GRID_SCROLL_SETTLE_SECONDS)
            marcar_somente_checkbox(checkbox)

            info = extrair_info_linha(checkbox)
            competencia = competencia_nota(info)
            comp = comparar_competencia_nota_intervalo(info, ano_alvo, 1, ano_alvo, 12)
            if comp is None:
                msg = f"Data de emissao invalida/ausente: {info.get('data_emissao', '')}"
                print(f"[{i+1}] SKIP DATA INVALIDA -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                try:
                    salvar_log("", info, status="SKIP_DATA_INVALIDA", mensagem=msg[:180])
                except Exception:
                    pass
                perf_log("prestados.apuracao_completa_skip_data_invalida", inicio_perf, f"indice={i+1}")
                return "SKIP", None

            if comp == 1:
                msg = f"Nota mais nova que a faixa alvo 01/{ano_alvo}..12/{ano_alvo}"
                print(f"[{i+1}] SKIP FAIXA (MAIS NOVA) -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                try:
                    salvar_log("", info, status="SKIP_FORA_COMPETENCIA", mensagem=msg[:180])
                except Exception:
                    pass
                perf_log("prestados.apuracao_completa_skip_competencia_nova", inicio_perf, f"indice={i+1}")
                return "SKIP", None

            if comp == -1:
                msg = f"Nota mais antiga que a faixa alvo 01/{ano_alvo}..12/{ano_alvo}"
                print(f"[{i+1}] SKIP FAIXA (MAIS ANTIGA) -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                try:
                    salvar_log("", info, status="SKIP_FORA_COMPETENCIA", mensagem=msg[:180])
                except Exception:
                    pass
                perf_log("prestados.apuracao_completa_stop", inicio_perf, f"indice={i+1}")
                return "STOP", None

            _processar_download_xml_prestados(checkbox, info, i, tentativa, inicio_perf)
            return "IN_RANGE", competencia

        except TimeoutException as e:
            msg = f"Timeout: {str(e)}"
            print(f"Erro nota [{i+1}] tentativa {tentativa}/{MAX_RETRIES_POR_NOTA}: {msg}")
            try:
                salvar_log("", info, status="ERRO", mensagem=msg[:180])
            except Exception:
                pass
            if is_502_page():
                recover_from_502()
            fechar_modal_exportacao()
            sleep(min(2 * tentativa, 6))
            perf_log("prestados.apuracao_completa_timeout", inicio_perf, f"indice={i+1} tentativa={tentativa}")

        except (StaleElementReferenceException, WebDriverException) as e:
            msg = str(e)
            print(f"Erro nota [{i+1}] tentativa {tentativa}/{MAX_RETRIES_POR_NOTA}: {msg}")
            try:
                salvar_log("", info, status="ERRO", mensagem=msg[:180])
            except Exception:
                pass
            if is_502_page():
                recover_from_502()
            fechar_modal_exportacao()
            sleep(min(2 * tentativa, 6))
            perf_log("prestados.apuracao_completa_webdriver", inicio_perf, f"indice={i+1} tentativa={tentativa}")

        except RuntimeError:
            raise

        except Exception as e:
            msg = str(e)
            print(f"Erro inesperado nota [{i+1}] tentativa {tentativa}/{MAX_RETRIES_POR_NOTA}: {msg}")
            try:
                salvar_log("", info, status="ERRO", mensagem=("Inesperado: " + msg)[:180])
            except Exception:
                pass
            if is_502_page():
                recover_from_502()
            fechar_modal_exportacao()
            sleep(min(2 * tentativa, 6))
            perf_log("prestados.apuracao_completa_erro", inicio_perf, f"indice={i+1} tentativa={tentativa}")

    print(f"Falhou apos {MAX_RETRIES_POR_NOTA} tentativas. Pulando nota {i+1}.")
    return "ERRO", None


def descrever_perfil_execucao() -> str:
    modulos = []
    saidas = []

    if EXECUTAR_PRESTADOS:
        modulos.append("Prestados")
    if EXECUTAR_TOMADOS:
        modulos.append("Tomados")
    if EXECUTAR_XML:
        saidas.append("XML")
    if EXECUTAR_LIVROS:
        saidas.append("Livros")
    if EXECUTAR_ISS:
        saidas.append("ISS")
    if PAUSA_MANUAL_FINAL:
        saidas.append("PausaFinal")
    if MODO_AUTOMATICO:
        saidas.append("ModoAutomatico")
    if APURAR_COMPLETO:
        saidas.append("ApurarCompleto")

    modulos_txt = ", ".join(modulos) if modulos else "nenhum"
    saidas_txt = ", ".join(saidas) if saidas else "nenhuma"
    return f"Modulos=[{modulos_txt}] | Saidas=[{saidas_txt}]"


def executar_etapa_servicos_prestados(ano_alvo: int, mes_alvo: int) -> str:
    global PARAR_PROCESSAMENTO, ENCONTROU_MES_ALVO, CONT_FORA_APOS_ALVO, CONT_FORA_ANTES_ALVO, SEM_COMPETENCIA_NA_EMPRESA, PRESTADOS_SEM_MODULO

    sem_competencia_inicial = False

    if PRESTADOS_SEM_MODULO:
        print("Prestados sem módulo de Nota Fiscal: etapa de Prestados será ignorada e o fluxo seguirá para Tomados.")
        try:
            garantir_log_com_header(log_path_empresa())
            append_txt(
                log_path_empresa(),
                "STATUS=SKIP_SEM_NOTA_FISCAL_PRESTADOS | MSG=Modulo Nota Fiscal indisponivel; etapa Prestados ignorada e Tomados processado",
            )
        except Exception:
            pass
    else:
        try:
            status_lista = esperar_lista_ou_sem_checkbox(timeout=20)
            if status_lista in {"sem_checkbox", "sem_checkbox_com_data"}:
                data_ref = primeira_data_emissao_visivel_sem_checkbox()
                comp = comparar_competencia_nota({"data_emissao": data_ref}, ano_alvo, mes_alvo) if data_ref else None

                if comp == -1:
                    print(
                        f"Sem checkbox e data inicial antiga ({data_ref}) para alvo {mes_alvo:02d}/{ano_alvo}. "
                        "Encerrando como sem competência."
                    )
                else:
                    print("Lista carregada sem checkbox; encerrando como sem competência.")

                sem_competencia_inicial = True
                print("Lista NFSe carregada sem checkbox; marcando empresa como sem competência para Prestados (etapa Tomados ainda será executada).")
        except TimeoutException:
            if STRICT_LISTA_INICIAL:
                raise RuntimeError(MSG_CAPTCHA_TIMEOUT)
            print("Aviso: lista inicial nao carregou em 20s; seguindo com tentativas por item.")

    ENCONTROU_MES_ALVO = False
    CONT_FORA_APOS_ALVO = 0
    CONT_FORA_ANTES_ALVO = 0
    SEM_COMPETENCIA_NA_EMPRESA = sem_competencia_inicial
    PARAR_PROCESSAMENTO = PRESTADOS_SEM_MODULO
    print(f"Mes alvo de download: {mes_alvo:02d}/{ano_alvo}")
    print(f"Heuristica de parada: {LIMITE_HEURISTICA_FORA_ALVO} notas antigas consecutivas (antes ou apos mês alvo).")

    if (not PRESTADOS_SEM_MODULO) and (not sem_competencia_inicial):
        try:
            checks0 = obter_checkboxes_lista(timeout=min(WAIT_LISTA_TIMEOUT, 5))
            if checks0:
                infos = [extrair_info_linha(checks0[0])]
                if len(checks0) >= 2:
                    infos.append(extrair_info_linha(checks0[1]))

                comps = [comparar_competencia_nota(info, ano_alvo, mes_alvo) for info in infos]
                if comps and all(c == -1 for c in comps):
                    SEM_COMPETENCIA_NA_EMPRESA = True
                    PARAR_PROCESSAMENTO = True
                    datas = ", ".join([info.get("data_emissao", "") for info in infos])
                    print(
                        f"Fast sem competência (Prestados): primeiras datas [{datas}] "
                        f"já são mais antigas que {mes_alvo:02d}/{ano_alvo}. Pulando varredura."
                    )
        except Exception as e:
            print(f"Aviso: fast-path sem competência não aplicado: {e}")

    if (not PRESTADOS_SEM_MODULO) and (not sem_competencia_inicial) and (not PARAR_PROCESSAMENTO) and len(obter_checkboxes_lista(timeout=min(WAIT_LISTA_TIMEOUT, 5))) > 0:
        definir_page_size(100)
    else:
        print("Page size não ajustado (lista sem checkbox ou sem notas na gridLista).")

    pagina = 1
    while True:
        if PARAR_PROCESSAMENTO:
            if PRESTADOS_SEM_MODULO:
                print("Prestados ignorado por ausência do módulo Nota Fiscal.")
                return "PRESTADOS_SEM_MODULO"
            else:
                print("Prestados encerrado pelo fast-path (sem competência).")
                return "SUCESSO_SEM_COMPETENCIA"

        checkboxes_pagina = obter_checkboxes_lista(timeout=min(WAIT_LISTA_TIMEOUT, 5) if pagina == 1 else 0)
        total = len(checkboxes_pagina)
        print(f"Pagina {pagina}: {total} notas encontradas.")

        if total == 0 and not ENCONTROU_MES_ALVO:
            SEM_COMPETENCIA_NA_EMPRESA = True
            print("Lista carregada sem checkboxes; encerrando empresa como sem competência.")
            return "SUCESSO_SEM_COMPETENCIA"

        i = 0
        while i < total:
            processar_nota_por_indice(i, ano_alvo, mes_alvo, checkboxes_cache=checkboxes_pagina)
            if PARAR_PROCESSAMENTO:
                break
            i += 1

        if PARAR_PROCESSAMENTO:
            print("Encerrando varredura por heuristica de competência.")
            return "SUCESSO_SEM_COMPETENCIA"

        if not ir_para_proxima_pagina():
            break
        pagina += 1

    if SEM_COMPETENCIA_NA_EMPRESA:
        return "SUCESSO_SEM_COMPETENCIA"
    if PRESTADOS_SEM_MODULO:
        return "PRESTADOS_SEM_MODULO"
    return "OK"


def executar_fechamento_fiscal(ano_alvo: int, mes_alvo: int, executar_tomados: bool, executar_prestados: bool) -> None:
    if not (EXECUTAR_LIVROS or EXECUTAR_ISS or PAUSA_MANUAL_FINAL):
        return

    print(f"Fechamento fiscal selecionado. {descrever_perfil_execucao()}")

    resultado_tardio = executar_fluxo_fiscal_multicadastro_tardio_se_necessario(
        ano_alvo,
        mes_alvo,
        executar_tomados_xml=False,
        executar_cadeado_tomados_flag=executar_tomados,
        executar_cadeado_prestados_flag=executar_prestados and not PRESTADOS_SEM_MODULO,
        executar_pausa_final=PAUSA_MANUAL_FINAL,
    )
    if resultado_tardio is not None:
        if resultado_tardio["status"] == "REVISAO_MANUAL":
            raise RuntimeError(f"{MSG_MULTI_CADASTRO_REVISAO_MANUAL}: {resultado_tardio['motivo']}")
        return

    if executar_tomados:
        try:
            executar_cadeado_tomados(ano_alvo, mes_alvo)
        except Exception as e:
            try:
                append_txt(caminho_log_manual(f"{mes_alvo:02d}/{ano_alvo}", ano_alvo, mes_alvo), f"TOMADOS | ERRO_CADEADO | {str(e)[:180]}")
            except Exception:
                pass
    else:
        print("Fechamento de Tomados não selecionado.")

    if executar_prestados:
        if PRESTADOS_SEM_MODULO:
            print("Prestados sem módulo Nota Fiscal: fechamento de Prestados será ignorado.")
        else:
            try:
                executar_cadeado_prestados(ano_alvo, mes_alvo)
            except Exception as e:
                try:
                    append_txt(caminho_log_manual(f"{mes_alvo:02d}/{ano_alvo}", ano_alvo, mes_alvo), f"PRESTADOS | ERRO_CADEADO | {str(e)[:180]}")
                except Exception:
                    pass
    else:
        print("Fechamento de Prestados não selecionado.")

    pausa_final_livros_manual(f"Competência alvo: {mes_alvo:02d}/{ano_alvo}")


def registrar_apuracao_completa(ref_alvo: str, etapa: str, status: str, mensagem: str = "") -> None:
    path = caminho_log_apuracao_completa(ref_alvo)
    parts = [
        f"REFERENCIA={ref_alvo}",
        f"ETAPA={etapa}",
        f"STATUS={status}",
    ]
    parts.extend(_partes_contexto_cadastro())
    if mensagem:
        parts.append(f"MSG={(mensagem or '').replace('|', '/')}")
    append_txt(path, " | ".join(parts))


def reabrir_lista_nota_fiscal() -> None:
    clicar_inicio_para_dashboard(timeout=40)
    navegar_para_lista_nota_fiscal()


def executar_etapa_servicos_prestados_apuracao_completa(ano_alvo: int) -> set[tuple[int, int]]:
    meses_encontrados: set[tuple[int, int]] = set()

    if PRESTADOS_SEM_MODULO:
        print("Prestados sem modulo de Nota Fiscal: apuracao completa de Prestados sera ignorada.")
        return meses_encontrados

    try:
        status_lista = esperar_lista_ou_sem_checkbox(timeout=20)
        if status_lista in {"sem_checkbox", "sem_checkbox_com_data"}:
            data_ref = primeira_data_emissao_visivel_sem_checkbox()
            comp = comparar_competencia_nota_intervalo({"data_emissao": data_ref}, ano_alvo, 1, ano_alvo, 12) if data_ref else None
            if comp == -1:
                print(
                    f"Sem checkbox e data inicial antiga ({data_ref}) para faixa 01/{ano_alvo}..12/{ano_alvo}. "
                    "Encerrando como sem competencia."
                )
                return meses_encontrados
            print("Lista carregada sem checkbox; apuracao completa de Prestados nao encontrou notas selecionaveis.")
            return meses_encontrados
    except TimeoutException:
        print("Aviso: lista inicial nao carregou em 20s; seguindo com tentativas por item na apuracao completa.")

    if len(obter_checkboxes_lista(timeout=min(WAIT_LISTA_TIMEOUT, 5))) > 0:
        definir_page_size(100)
    else:
        print("Page size nao ajustado na apuracao completa (lista sem checkboxes).")

    pagina = 1
    while True:
        checkboxes_pagina = obter_checkboxes_lista(timeout=min(WAIT_LISTA_TIMEOUT, 5) if pagina == 1 else 0)
        total = len(checkboxes_pagina)
        print(f"Pagina {pagina}: {total} notas encontradas na apuracao completa de Prestados {ano_alvo}.")

        if total == 0:
            break

        parar_faixa = False
        i = 0
        while i < total:
            status_item, competencia = processar_nota_por_indice_apuracao_completa(
                i,
                ano_alvo,
                checkboxes_cache=checkboxes_pagina,
            )
            if competencia is not None:
                meses_encontrados.add(competencia)
            if status_item == "STOP":
                parar_faixa = True
                break
            i += 1

        if parar_faixa:
            print("Prestados apuracao completa: faixa anual encerrada ao encontrar a primeira nota anterior a 01 do ano alvo.")
            break

        if not ir_para_proxima_pagina():
            break
        pagina += 1

    return meses_encontrados


def executar_fluxo_apuracao_completa(apuracao_ref: str) -> None:
    competencias = runtime_competencias_alvo(apuracao_ref, apurar_completo=True)
    competencias_desc = sorted(competencias, key=lambda item: (item[0], item[1]), reverse=True)
    falhas: list[str] = []
    prestados_ok = 0
    tomados_ok = 0
    competencias_prestados = set()

    print(
        "APURAR_COMPLETO=1 -> "
        f"{len(competencias)} competencias no ano alvo {competencias[0][0]} "
        f"(01/{competencias[0][0]} ate {competencias[-1][1]:02d}/{competencias[-1][0]})."
    )
    print(f"PERFIL_EXECUCAO_ATIVO={'1' if PERFIL_EXECUCAO_ATIVO else '0'} -> {descrever_perfil_execucao()}")

    if EXECUTAR_PRESTADOS and not PRESTADOS_SEM_MODULO:
        try:
            competencias_prestados = executar_etapa_servicos_prestados_apuracao_completa(competencias[0][0])
            prestados_ok = len(competencias_prestados)
        except Exception as exc:
            msg = f"{type(exc).__name__}: {str(exc)[:180]}"
            registrar_apuracao_completa(f"ANO/{competencias[0][0]}", "PRESTADOS", "ERRO", msg)
            falhas.append(f"ANO/{competencias[0][0]} PRESTADOS -> {msg}")

    for idx, (ano_alvo, mes_alvo) in enumerate(competencias_desc, start=1):
        ref_alvo = f"{mes_alvo:02d}/{ano_alvo}"
        print("")
        print(f"[APURACAO_COMPLETA] Competencia {idx}/{len(competencias)} -> {ref_alvo}")
        registrar_apuracao_completa(ref_alvo, "COMPETENCIA", "INICIO", f"indice={idx}/{len(competencias)}")

        if EXECUTAR_PRESTADOS:
            if PRESTADOS_SEM_MODULO:
                registrar_apuracao_completa(ref_alvo, "PRESTADOS", "SKIP_SEM_MODULO")
            elif not any(f"PRESTADOS ->" in falha for falha in falhas):
                status_prestados = "OK" if (ano_alvo, mes_alvo) in competencias_prestados else "SUCESSO_SEM_COMPETENCIA"
                registrar_apuracao_completa(ref_alvo, "PRESTADOS", status_prestados)

        if EXECUTAR_TOMADOS:
            try:
                status_tomados = executar_etapa_servicos_tomados(ano_alvo, mes_alvo)
                registrar_apuracao_completa(ref_alvo, "TOMADOS", status_tomados)
                if status_tomados == "OK":
                    tomados_ok += 1
            except Exception as exc:
                msg = f"{type(exc).__name__}: {str(exc)[:180]}"
                registrar_apuracao_completa(ref_alvo, "TOMADOS", "ERRO", msg)
                falhas.append(f"{ref_alvo} TOMADOS -> {msg}")

    print(
        "[APURACAO_COMPLETA] Resumo -> "
        f"PrestadosOK={prestados_ok} | TomadosOK={tomados_ok} | Falhas={len(falhas)}"
    )

    if falhas:
        resumo = " ; ".join(falhas[:4])
        raise RuntimeError(
            f"{MSG_APURACAO_COMPLETA_REVISAO_MANUAL}: "
            f"{len(falhas)} competencia(s) com falha. {resumo}"
        )

    if EXECUTAR_PRESTADOS and not PRESTADOS_SEM_MODULO and prestados_ok == 0 and tomados_ok == 0:
        raise RuntimeError(MSG_SEM_COMPETENCIA)

    print("Processo finalizado.")
    sleep(2)


def executar_fluxo_personalizado(ano_alvo: int, mes_alvo: int) -> None:
    global PARAR_PROCESSAMENTO, ENCONTROU_MES_ALVO, CONT_FORA_APOS_ALVO, CONT_FORA_ANTES_ALVO, SEM_COMPETENCIA_NA_EMPRESA

    print(f"PERFIL_EXECUCAO_ATIVO=1 -> {descrever_perfil_execucao()}")

    if APURAR_COMPLETO:
        executar_fluxo_apuracao_completa(APURACAO_REFERENCIA or datetime.now().strftime("%m/%Y"))
        return

    if not (EXECUTAR_PRESTADOS or EXECUTAR_TOMADOS):
        print("Perfil customizado sem módulos selecionados. Encerrando sem executar.")
        return

    if not (EXECUTAR_XML or EXECUTAR_LIVROS):
        print("Perfil customizado sem XML nem Livros selecionados. Encerrando sem executar.")
        return

    ENCONTROU_MES_ALVO = False
    CONT_FORA_APOS_ALVO = 0
    CONT_FORA_ANTES_ALVO = 0
    SEM_COMPETENCIA_NA_EMPRESA = False
    PARAR_PROCESSAMENTO = False

    if EXECUTAR_PRESTADOS and EXECUTAR_XML:
        executar_etapa_servicos_prestados(ano_alvo, mes_alvo)
    elif EXECUTAR_PRESTADOS:
        print("Prestados selecionado apenas para fechamento fiscal. Etapa XML de Prestados será ignorada.")
    else:
        print("Prestados não selecionado no perfil customizado.")

    resultado_tardio = executar_fluxo_fiscal_multicadastro_tardio_se_necessario(
        ano_alvo,
        mes_alvo,
        executar_tomados_xml=EXECUTAR_TOMADOS and EXECUTAR_XML,
        executar_cadeado_tomados_flag=EXECUTAR_TOMADOS and (EXECUTAR_LIVROS or EXECUTAR_ISS or PAUSA_MANUAL_FINAL),
        executar_cadeado_prestados_flag=EXECUTAR_PRESTADOS and (EXECUTAR_LIVROS or EXECUTAR_ISS or PAUSA_MANUAL_FINAL) and not PRESTADOS_SEM_MODULO,
        executar_pausa_final=PAUSA_MANUAL_FINAL and (EXECUTAR_LIVROS or EXECUTAR_ISS or EXECUTAR_TOMADOS or EXECUTAR_PRESTADOS),
    )
    if resultado_tardio is not None:
        if resultado_tardio["status"] == "REVISAO_MANUAL":
            raise RuntimeError(f"{MSG_MULTI_CADASTRO_REVISAO_MANUAL}: {resultado_tardio['motivo']}")
        print("Processo finalizado.")
        sleep(2)
        return

    if EXECUTAR_TOMADOS and EXECUTAR_XML:
        executar_etapa_servicos_tomados(ano_alvo, mes_alvo)
    elif EXECUTAR_TOMADOS:
        print("Tomados selecionado apenas para fechamento fiscal. Etapa XML de Tomados será ignorada.")
    else:
        print("Tomados não selecionado no perfil customizado.")

    executar_fechamento_fiscal(
        ano_alvo,
        mes_alvo,
        executar_tomados=EXECUTAR_TOMADOS,
        executar_prestados=EXECUTAR_PRESTADOS,
    )

    if EXECUTAR_PRESTADOS and EXECUTAR_XML and SEM_COMPETENCIA_NA_EMPRESA:
        raise RuntimeError(MSG_SEM_COMPETENCIA)

    print("Processo finalizado.")
    sleep(2)


def precisa_lista_nota_fiscal_inicial() -> bool:
    if MODO_APENAS_CADEADO and not PERFIL_EXECUCAO_ATIVO:
        return False
    if PERFIL_EXECUCAO_ATIVO:
        return EXECUTAR_PRESTADOS and EXECUTAR_XML
    if APURAR_COMPLETO:
        return EXECUTAR_PRESTADOS and EXECUTAR_XML
    return True


def preparar_execucao_atual(reset_prestados_sem_modulo: bool = False) -> None:
    global PARAR_PROCESSAMENTO, ENCONTROU_MES_ALVO, CONT_FORA_APOS_ALVO, CONT_FORA_ANTES_ALVO
    global SEM_COMPETENCIA_NA_EMPRESA, ULTIMO_CHECKBOX_MARCADO, chaves_ok, GRID_HEADER_CACHE, PRESTADOS_SEM_MODULO

    resetar_multicadastro_tardio()
    PARAR_PROCESSAMENTO = False
    ENCONTROU_MES_ALVO = False
    CONT_FORA_APOS_ALVO = 0
    CONT_FORA_ANTES_ALVO = 0
    SEM_COMPETENCIA_NA_EMPRESA = False
    ULTIMO_CHECKBOX_MARCADO = None
    GRID_HEADER_CACHE = {"ts": 0.0, "map": {}}
    if reset_prestados_sem_modulo:
        PRESTADOS_SEM_MODULO = False
    chaves_ok = carregar_chaves_ok(log_path_empresa())
    print(f"Chaves OK carregadas do log: {len(chaves_ok)}")


def executar_fluxo_empresa_atual(apuracao_ref: str, ano_alvo: int, mes_alvo: int) -> str:
    if precisa_lista_nota_fiscal_inicial() and not PRESTADOS_SEM_MODULO:
        garantir_lista_nota_fiscal_contexto(timeout=40)

    if MODO_APENAS_CADEADO and not PERFIL_EXECUCAO_ATIVO:
        print("MODO_APENAS_CADEADO=1 -> Somente fechar período (cadeado).")
        if PAUSAR_ENTRE_MODULOS_CADEADO:
            print("PAUSAR_ENTRE_MODULOS_CADEADO=1 -> haverá pausa entre Tomados e Prestados.")
        else:
            print("PAUSAR_ENTRE_MODULOS_CADEADO=0 -> Tomados e Prestados serão executados em sequência, com pausa apenas no final.")
        executar_fechamento_fiscal(
            ano_alvo,
            mes_alvo,
            executar_tomados=True,
            executar_prestados=True,
        )
        print("MODO_APENAS_CADEADO finalizado. Encerrando execução.")
        return "SUCESSO"

    if APURAR_COMPLETO and not PERFIL_EXECUCAO_ATIVO:
        executar_fluxo_apuracao_completa(apuracao_ref)
        return "SUCESSO"

    if PERFIL_EXECUCAO_ATIVO:
        executar_fluxo_personalizado(ano_alvo, mes_alvo)
        return "SUCESSO"

    executar_etapa_servicos_prestados(ano_alvo, mes_alvo)

    resultado_tardio = executar_fluxo_fiscal_multicadastro_tardio_se_necessario(
        ano_alvo,
        mes_alvo,
        executar_tomados_xml=True,
        executar_cadeado_tomados_flag=ACRESCENTAR_CADEADO_E_PAUSA_FINAL,
        executar_cadeado_prestados_flag=ACRESCENTAR_CADEADO_E_PAUSA_FINAL and not PRESTADOS_SEM_MODULO,
        executar_pausa_final=ACRESCENTAR_CADEADO_E_PAUSA_FINAL and PAUSA_MANUAL_FINAL,
    )
    if resultado_tardio is not None:
        if resultado_tardio["status"] == "REVISAO_MANUAL":
            raise RuntimeError(f"{MSG_MULTI_CADASTRO_REVISAO_MANUAL}: {resultado_tardio['motivo']}")
        print("Processo finalizado.")
        sleep(2)
        return "SUCESSO"

    executar_etapa_servicos_tomados(ano_alvo, mes_alvo)

    if ACRESCENTAR_CADEADO_E_PAUSA_FINAL:
        if PRESTADOS_SEM_MODULO:
            print("ACRESCENTAR_CADEADO_E_PAUSA_FINAL=1 -> clicando cadeado vermelho somente em Tomados (Prestados indisponível) e pausando ao final.")
        else:
            print("ACRESCENTAR_CADEADO_E_PAUSA_FINAL=1 -> clicando cadeado vermelho em Tomados e Prestados (competência alvo) e pausando ao final.")
        executar_fechamento_fiscal(
            ano_alvo,
            mes_alvo,
            executar_tomados=True,
            executar_prestados=True,
        )

    if SEM_COMPETENCIA_NA_EMPRESA:
        raise RuntimeError(MSG_SEM_COMPETENCIA)

    print("Processo finalizado.")
    sleep(2)
    return "SUCESSO"


def _resultado_execucao_cadastro(exc: Exception | None) -> dict:
    if exc is None:
        return {"status": "SUCESSO", "motivo": "OK"}

    msg = str(exc)
    if MSG_SEM_COMPETENCIA in msg:
        return {"status": "SUCESSO_SEM_COMPETENCIA", "motivo": MSG_SEM_COMPETENCIA}
    if MSG_SEM_SERVICOS in msg:
        return {"status": "SUCESSO_SEM_SERVICOS", "motivo": MSG_SEM_SERVICOS}
    return {"status": "FALHA", "motivo": msg}


def executar_fluxo_cadastro(cadastro: dict, apuracao_ref: str, ano_alvo: int, mes_alvo: int) -> dict:
    ativar_contexto_cadastro(cadastro)
    preparar_execucao_atual(reset_prestados_sem_modulo=True)

    prefixo = _prefixo_cadastro_base(cadastro)
    print("")
    print(f"[MULTI_CADASTRO] Iniciando {prefixo or 'CADASTRO'} -> {cadastro.get('identificacao') or cadastro.get('ccm') or cadastro.get('ordem')}")

    try:
        executar_fluxo_empresa_atual(apuracao_ref, ano_alvo, mes_alvo)
        resultado = _resultado_execucao_cadastro(None)
    except Exception as exc:
        resultado = _resultado_execucao_cadastro(exc)
        if resultado["status"] == "FALHA":
            print(f"[MULTI_CADASTRO] Falha no cadastro {prefixo or cadastro.get('ordem')}: {resultado['motivo'][:180]}")
        else:
            print(f"[MULTI_CADASTRO] Cadastro {prefixo or cadastro.get('ordem')} concluído com status {resultado['status']}.")

    return {
        "cadastro": dict(cadastro),
        "status": resultado["status"],
        "motivo": resultado["motivo"],
    }


def agregar_resultados_multicadastro(resultados: list[dict]) -> tuple[str, str]:
    falhas = [r for r in resultados if r.get("status") == "FALHA"]
    if falhas:
        detalhes = []
        for item in falhas[:4]:
            cadastro = item.get("cadastro") or {}
            prefixo = _prefixo_cadastro_base(cadastro) or f"CAD_{int(cadastro.get('ordem') or 0):02d}"
            detalhes.append(f"{prefixo}: {str(item.get('motivo') or '')[:120]}")
        return "REVISAO_MANUAL", " ; ".join(detalhes)

    statuses = {str(item.get("status") or "") for item in resultados}
    if statuses == {"SUCESSO_SEM_COMPETENCIA"}:
        return "SUCESSO_SEM_COMPETENCIA", MSG_SEM_COMPETENCIA
    if statuses == {"SUCESSO_SEM_SERVICOS"}:
        return "SUCESSO_SEM_SERVICOS", MSG_SEM_SERVICOS
    return "SUCESSO", "OK"


def executar_fluxo_multiplos_cadastros(apuracao_ref: str, ano_alvo: int, mes_alvo: int) -> None:
    if not tela_cadastros_relacionados_aberta():
        raise RuntimeError(MSG_EMPRESA_MULTIPLA)

    try:
        cadastros = coletar_cadastros_relacionados()
        if not cadastros:
            salvar_print_evidencia("CADASTROS_RELACIONADOS", "GRID_NAO_IDENTIFICADA", ano_alvo=ano_alvo, mes_alvo=mes_alvo)
            raise RuntimeError(MSG_EMPRESA_MULTIPLA)

        print(f"[MULTI_CADASTRO] Cadastros identificados: {len(cadastros)}")
        resultados = []

        for idx, cadastro in enumerate(cadastros):
            cadastro = _sincronizar_contexto_cadastro_relacionado(cadastro)
            if idx > 0:
                try:
                    retornar_para_cadastros_relacionados(timeout=40)
                except Exception as exc:
                    resultados.append(
                        {
                            "cadastro": dict(cadastro),
                            "status": "FALHA",
                            "motivo": f"Falha ao retornar para Cadastros Relacionados: {type(exc).__name__}: {str(exc)[:160]}",
                        }
                    )
                    continue

            try:
                cadastro = selecionar_cadastro_relacionado(cadastro, timeout=40)
                resultados.append(executar_fluxo_cadastro(cadastro, apuracao_ref, ano_alvo, mes_alvo))
            except Exception as exc:
                resultados.append(
                    {
                        "cadastro": dict(cadastro),
                        "status": "FALHA",
                        "motivo": f"{type(exc).__name__}: {str(exc)[:180]}",
                    }
                )

        status_final, motivo_final = agregar_resultados_multicadastro(resultados)
        print(f"[MULTI_CADASTRO] Resumo -> total={len(resultados)} | status={status_final} | motivo={motivo_final}")

        if status_final == "REVISAO_MANUAL":
            raise RuntimeError(f"{MSG_MULTI_CADASTRO_REVISAO_MANUAL}: {motivo_final}")
        if status_final == "SUCESSO_SEM_COMPETENCIA":
            raise RuntimeError(MSG_SEM_COMPETENCIA)
        if status_final == "SUCESSO_SEM_SERVICOS":
            raise RuntimeError(MSG_SEM_SERVICOS)
    finally:
        resetar_contexto_cadastro()

# =====================
# MAIN
# =====================
def main():
    apuracao_ref = APURACAO_REFERENCIA or datetime.now().strftime("%m/%Y")
    validate_apuracao_referencia(apuracao_ref)

    inicializar_chrome()
    aguardar_navegacao_manual_inicial()
    status_login = preencher_login_prefeitura_se_habilitado()
    resolver_pasta_empresa_atual()

    ano_alvo, mes_alvo = calcular_mes_alvo(apuracao_ref)
    if status_login == "cadastros_relacionados" or tela_cadastros_relacionados_aberta():
        executar_fluxo_multiplos_cadastros(apuracao_ref, ano_alvo, mes_alvo)
        return

    preparar_execucao_atual(reset_prestados_sem_modulo=False)
    executar_fluxo_empresa_atual(apuracao_ref, ano_alvo, mes_alvo)


def _run_main_entrypoint():
    try:
        main()
    except RuntimeError as e:
        msg = str(e)
        print(msg)

        # Erros controlados (não são "forenses")
        exit_code = map_runtime_error_to_exit_code(msg)
        if exit_code is not None:
            sys.exit(exit_code)

        # Inesperado -> kit forense + re-raise
        salvar_kit_forense("RuntimeError inesperado", e)
        raise
    except Exception as e:
        # Qualquer exceção não mapeada -> kit forense
        print(f"Erro inesperado: {type(e).__name__}: {e}")
        salvar_kit_forense("Exceção não tratada", e)
        raise
    finally:
        if driver is not None:
            driver.quit()
        if TEMP_PROFILE_DIR:
            shutil.rmtree(TEMP_PROFILE_DIR, ignore_errors=True)


if __name__ == "__main__":
    _run_main_entrypoint()
