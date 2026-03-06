from selenium import webdriver
from selenium.webdriver.chrome.options import Options
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
import math
import ssl
import unicodedata
from datetime import datetime
import xml.etree.ElementTree as ET
import shutil
from urllib import request as urllib_request
from urllib import parse as urllib_parse

# =====================
# CONFIGURACAO
# =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

TEMP_DOWNLOAD_DIR = os.path.join(DOWNLOAD_DIR, "_tmp")
os.makedirs(TEMP_DOWNLOAD_DIR, exist_ok=True)


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



def append_txt(path: str, line: str):
    """Append a human-readable line to a .txt log (one event per line)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    safe = (line or "").replace("\n", " ").replace("\r", " ").strip()
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {safe}\n")


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
MAX_RETRIES_POR_NOTA = 3
WAIT_LISTA_TIMEOUT = 60
LOG_FILENAME = "log_downloads_nfse.txt"
TOMADOS_LOG_FILENAME = "log_tomados.txt"
TOMADOS_XML_BASENAME = "SERVICOS_TOMADOS_{mm_aaaa}.xml"  # ex: SERVICOS_TOMADOS_01-2026.xml
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
LIMITE_HEURISTICA_FORA_ALVO = int(os.environ.get("LIMITE_HEURISTICA_FORA_ALVO", "2"))
STRICT_LISTA_INICIAL = os.environ.get("STRICT_LISTA_INICIAL", "0").strip() == "1"
MSG_CAPTCHA_TIMEOUT = "CAPTCHA_NAO_RESOLVIDO_NO_TEMPO"
MSG_CAPTCHA_INCORRETO = "CAPTCHA_INCORRETO"
MSG_SEM_COMPETENCIA = "SUCESSO_SEM_COMPETENCIA"
MSG_SEM_SERVICOS = "SUCESSO_SEM_SERVICOS"
MSG_CREDENCIAL_INVALIDA = "CREDENCIAL_INVALIDA"
EXIT_CODE_CAPTCHA_TIMEOUT = 30
EXIT_CODE_SEM_COMPETENCIA = 40
EXIT_CODE_SEM_SERVICOS = 41
EXIT_CODE_CREDENCIAL_INVALIDA = 50
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
options = Options()
options.add_argument("--start-maximized")

# Perfil exclusivo com fallback cross-platform.
DEFAULT_PROFILE_DIR = (
    r"C:\ChromeRobotProfile" if os.name == "nt"
    else os.path.join(tempfile.gettempdir(), "ChromeRobotProfile")
)

TEMP_PROFILE_DIR = None
if FORCAR_SESSAO_LIMPA_LOGIN:
    TEMP_PROFILE_DIR = tempfile.mkdtemp(prefix="ChromeRobotProfile_")
    CHROME_PROFILE_DIR = TEMP_PROFILE_DIR
else:
    CHROME_PROFILE_DIR = os.environ.get("CHROME_PROFILE_DIR", DEFAULT_PROFILE_DIR)

options.add_argument(f"--user-data-dir={CHROME_PROFILE_DIR}")
options.add_argument("--ignore-certificate-errors")
options.add_argument("--ignore-ssl-errors")
options.add_argument("--allow-running-insecure-content")
options.add_argument("--disable-features=InsecureDownloadWarnings")
options.add_argument("--safebrowsing-disable-download-protection")
options.add_argument("--no-sandbox")

prefs = {
    "download.default_directory": TEMP_DOWNLOAD_DIR,  # importante
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False,
    "safebrowsing.disable_download_protection": True,
    "plugins.always_open_pdf_externally": True,
    "profile.default_content_setting_values.automatic_downloads": 1,
}
options.add_experimental_option("prefs", prefs)
options.add_experimental_option("excludeSwitches", ["enable-automation"])

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
        # log simples no raiz da empresa
        append_txt(os.path.join(pasta_empresa(), "log_fechamento_manual.txt"), msg)
    except Exception:
        pass

    if PAUSA_MANUAL_SEGUNDOS <= 0:
        try:
            aguardar_enter_robusto("[MANUAL] Pressione ENTER para continuar...", tag="MANUAL")
        except Exception:
            sleep(5)
    else:
        print(f"[MANUAL] Aguardando {PAUSA_MANUAL_SEGUNDOS}s para ações manuais...")
        sleep(PAUSA_MANUAL_SEGUNDOS)


def pausa_final_livros_manual(contexto: str = ""):
    """Pausa final para baixar os livros manualmente. Encerra apenas ao pressionar ENTER."""
    if not ACRESCENTAR_CADEADO_E_PAUSA_FINAL:
        return
    msg = f"[MANUAL-FINAL] {contexto}".strip() if contexto else "[MANUAL-FINAL] Baixe os livros manualmente (Tomados/Prestados)."
    print(msg)
    try:
        append_txt(os.path.join(pasta_empresa(), "log_fechamento_manual.txt"), msg)
    except Exception:
        pass
    try:
        aguardar_enter_robusto("[MANUAL-FINAL] Pressione ENTER para encerrar o main.py...", tag="MANUAL-FINAL")
    except Exception:
        sleep(5)


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
    try:
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 30)
        print(f"Chrome iniciado. Perfil: {CHROME_PROFILE_DIR}")
    except Exception as e:
        raise RuntimeError(f"{MSG_CHROME_INIT_FALHA}: {e}")



# =====================
# KIT FORENSE (DEBUG)
# =====================
def _agora_ts():
    # 2026-02-25_14-03-22
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def salvar_kit_forense(motivo, exc=None):
    """Salva evidências (screenshot, HTML, URL, logs) em downloads/<empresa>/_debug/<ts>/.
    Retorna o caminho do diretório criado (ou None se não foi possível).
    """
    try:
        empresa_pasta = globals().get("EMPRESA_PASTA") or os.environ.get("EMPRESA_PASTA_FORCADA") or "SEM_EMPRESA"
        ts = _agora_ts()
        debug_dir = os.path.join(DOWNLOAD_DIR, empresa_pasta, "_debug", ts)
        os.makedirs(debug_dir, exist_ok=True)

        # Metadados básicos
        with open(os.path.join(debug_dir, "info.txt"), "w", encoding="utf-8") as f:
            f.write(f"timestamp={ts}\n")
            f.write(f"motivo={motivo}\n")
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


def _slug_evidencia(txt: str) -> str:
    base = _norm_ascii(txt or "")
    base = re.sub(r"[^a-z0-9]+", "_", base).strip("_")
    return base or "evidencia"


def salvar_print_evidencia(contexto: str, motivo: str, ano_alvo: int = 0, mes_alvo: int = 0, ref_alvo: str = "", log_path: str = "") -> str:
    """Salva screenshot completo com carimbo visível de data/hora para fins de evidência."""
    ts = datetime.now()
    ts_visivel = ts.strftime("%d/%m/%Y %H:%M:%S")
    ts_arquivo = ts.strftime("%Y-%m-%d_%H-%M-%S")

    base_dir = os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA)
    if ano_alvo and mes_alvo:
        base_dir = os.path.join(base_dir, f"{mes_alvo:02d}.{ano_alvo}")
    evid_dir = os.path.join(base_dir, "_evidencias")
    os.makedirs(evid_dir, exist_ok=True)

    partes_nome = [_slug_evidencia(contexto), _slug_evidencia(ref_alvo or motivo), ts_arquivo]
    destino = os.path.join(evid_dir, "_".join([p for p in partes_nome if p]) + ".png")

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
        return

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

    print("CNPJ e senha preenchidos. Resolva o captcha e clique em 'Entrar' manualmente.")
    print(f"Aguardando dashboard por até {login_wait_seconds}s...")

    fim = time.time() + login_wait_seconds
    while time.time() < fim:
        if credencial_invalida_na_tela():
            raise RuntimeError(MSG_CREDENCIAL_INVALIDA)
        if captcha_incorreto_na_tela():
            raise RuntimeError(MSG_CAPTCHA_INCORRETO)
        if sem_modulo_nota_fiscal_no_dashboard():
            PRESTADOS_SEM_MODULO = True
            print("Módulo Nota Fiscal não disponível no dashboard. Prestados será ignorado e o fluxo seguirá para Tomados.")
            break
        if driver.find_elements(By.ID, LOGIN_CARD_DASHBOARD):
            break
        sleep(0.4)
    else:
        raise RuntimeError(MSG_CAPTCHA_TIMEOUT)

    if PRESTADOS_SEM_MODULO:
        return

    navegar_para_lista_nota_fiscal()


if not AUTO_LOGIN_PREFEITURA:
    print("Entre manualmente em: Nota Fiscal - Lista Nota Fiscais")
    print(f"Você tem {login_wait_seconds} segundos para iniciar manualmente.")
    sleep(login_wait_seconds)

# =====================
# HELPERS â€“ CLIQUE / LISTA / 502
# =====================
def click_robusto(el):
    try:
        el.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", el)


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
    WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located((By.NAME, "gridListaCheck"))
    )

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

    return WebDriverWait(driver, timeout).until(cond)

def lista_ativa():
    try:
        return len(driver.find_elements(By.NAME, "gridListaCheck")) > 0
    except Exception:
        return False

def assinatura_lista():
    """Assinatura leve da grid para detectar refresh/paginação."""
    try:
        checks = driver.find_elements(By.NAME, "gridListaCheck")
        values = [c.get_attribute("value") or "" for c in checks[:5]]
        return f"{len(checks)}|{'|'.join(values)}"
    except Exception:
        return ""

def esperar_troca_de_grid(assinatura_anterior, timeout=WAIT_LISTA_TIMEOUT):
    def mudou(_):
        return assinatura_lista() != assinatura_anterior
    WebDriverWait(driver, timeout).until(mudou)

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
    txt = (txt or "").replace("\xa0", " ").strip().lower()
    txt = txt.replace("Âº", "o")
    txt = re.sub(r"\s+", " ", txt)
    return txt


def mapa_colunas_grid_lista(force=False):
    """
    Mapeia colunas pelo THEAD para evitar leitura deslocada (empresa/tema diferente).
    O `columnorder` do TH começa em 0 após a coluna de checkbox,
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

            if ("data emissao" in titulo) and ("rps" not in titulo):
                col_map["data_emissao"] = idx_td
            elif ("situacao" in titulo) or ("situação" in titulo):
                col_map["situacao"] = idx_td
            elif " chave de validacao/acesso" in f" {titulo}" or "chave de validacao" in titulo:
                col_map["chave"] = idx_td
            elif "rps" in titulo and "data" not in titulo and "serie" not in titulo and "série" not in titulo:
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

    return info

def nota_cancelada(info: dict) -> bool:
    situacao = (info.get("situacao") or "").strip().lower()
    return "cancelad" in situacao

def chave_unica(info: dict) -> str:
    ch = (info.get("chave") or "").strip()
    if ch:
        return ch
    iid = (info.get("id_interno") or "").strip()
    if iid:
        return f"ID:{iid}"
    nf = (info.get("nf") or "").strip()
    rps = (info.get("rps") or "").strip()
    dt = (info.get("data_emissao") or "").strip()
    return f"NF:{nf}|RPS:{rps}|DT:{dt}"

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
    parsed = parse_data_emissao_site(info.get("data_emissao", ""))
    if not parsed:
        return None

    ano = int(parsed[0])
    mes = int(parsed[1])
    comp_nota = (ano, mes)
    comp_alvo = (ano_alvo, mes_alvo)

    if comp_nota == comp_alvo:
        return 0
    if comp_nota > comp_alvo:
        return 1
    return -1

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
    fim = time.time() + timeout
    xmls_antes = set(xmls_antes or [])

    while time.time() < fim:
        arquivos = list(os.scandir(pasta))

        if any(e.name.lower().endswith(".crdownload") for e in arquivos):
            sleep(0.4)
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
            return os.path.join(pasta, novo_mais_recente)

        sleep(0.4)

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


def baixar_arquivo_via_submit_form(button, timeout=120, nome_fallback="arquivo.xml") -> str:
    """
    Executa o submit relacionado ao botão dentro da sessão autenticada e
    baixa o payload via fetch (sem usar o gerenciador de downloads do Chrome).
    """
    timeout_ms = max(5000, int(timeout * 1000))
    script = r"""
const btn = arguments[0];
const timeoutMs = arguments[1];
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

    await new Promise((r) => setTimeout(r, 80));

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

    result = driver.execute_async_script(script, button, timeout_ms)
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

    if "xml" in content_type and not nome.lower().endswith(".xml"):
        nome = f"{os.path.splitext(nome)[0]}.xml"
    elif not os.path.splitext(nome)[1]:
        nome = f"{nome}.xml"

    # Validação leve para evitar salvar HTML de erro como XML.
    head = payload[:200].lstrip()
    if not (
        head.startswith(b"<?xml")
        or head.startswith(b"<")
        or b"<nfe" in payload[:2048].lower()
        or b"<nfse" in payload[:2048].lower()
    ):
        raise RuntimeError("Resposta do submit não parece XML válido")

    return _salvar_bytes_tmp(payload, nome)


def aguardar_pdf_novo(pasta, timeout=120, pdfs_antes=None):
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
                return destino
            except Exception:
                continue

        sleep(0.4)

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


def baixar_pdf_por_url_com_cookies(pdf_url: str, timeout=120) -> str:
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
    }
    if cookie_header:
        headers["Cookie"] = cookie_header

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

# =====================
# EMPRESA â€“ PASTA CANONICA (evita H2_IMOBILIARIA vs IMOBILIARIA_H2_LTDA)
# =====================
SUFIXOS_LEGAIS = {
    "LTDA", "LTDA.", "ME", "EPP", "EIRELI",
    "S/A", "SA", "S.A", "S.A.",
}
_SIGLA_ALFANUM_PATTERN = re.compile(r"^(?=.*[A-Z])(?=.*\d)[A-Z0-9]{2,10}$")

def normalizar_nome_empresa(nome: str) -> str:
    """
    Gera nome de pasta no padrão "clean" do escritório.

    Regras:
      - Remove anotações em (...) e [...]
      - Remove acentos e caracteres especiais
      - Junta siglas de letras separadas (ex.: "A. B. S." -> "ABS")
      - Remove sufixos legais (LTDA, ME, EPP, EIRELI, SA, etc.)
      - Mantém a ordem dos termos e separa por underscore
    """
    raw = (nome or "").strip()
    if not raw:
        return "EMPRESA_DESCONHECIDA"

    # Remove observações internas comuns na planilha
    raw = re.sub(r"\[[^\]]*\]", " ", raw)   # [ ... ]
    raw = re.sub(r"\([^)]*\)", " ", raw)     # ( ... )

    raw = unicodedata.normalize("NFKD", raw.upper())
    raw = "".join(ch for ch in raw if not unicodedata.combining(ch))

    # Troca separadores por espaço e remove tudo que não for alfanumérico
    raw = re.sub(r"[\/\.\-]", " ", raw)
    raw = re.sub(r"[^A-Z0-9\s]", " ", raw)
    raw = re.sub(r"\s+", " ", raw).strip()

    tokens = [t for t in raw.split(" ") if t]
    if not tokens:
        return "EMPRESA_DESCONHECIDA"

    # Junta sequências de letras isoladas (A B S -> ABS)
    merged = []
    buf = []

    def flush_buf():
        nonlocal buf
        if buf:
            merged.append("".join(buf))
            buf = []

    for t in tokens:
        if len(t) == 1 and t.isalpha():
            buf.append(t)
        else:
            flush_buf()
            merged.append(t)
    flush_buf()

    # Remove sufixos legais (após juntar siglas)
    cleaned = [t for t in merged if t not in SUFIXOS_LEGAIS]

    joined = "_".join(cleaned) if cleaned else "EMPRESA_DESCONHECIDA"
    joined = re.sub(r"_+", "_", joined).strip("_")

    # Evita nomes gigantes (path do Windows pode estourar facilmente)
    max_len = int(os.environ.get("EMPRESA_PASTA_MAXLEN", "80"))
    if len(joined) > max_len:
        joined = joined[:max_len].rstrip("_")

    return joined or "EMPRESA_DESCONHECIDA"

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

if EMPRESA_PASTA_FORCADA.strip():
    EMPRESA_PASTA = normalizar_nome_empresa(EMPRESA_PASTA_FORCADA)
else:
    EMPRESA_PASTA = normalizar_nome_empresa(detectar_nome_empresa_da_tela())

print(f"Pasta da empresa: {EMPRESA_PASTA}")

# =====================
# LOG (1 por empresa) + IDP
# =====================
def pasta_empresa():
    return os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA)


def log_path_empresa():
    return os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA, LOG_FILENAME)

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
    path_ = log_path_empresa()
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

def organizar_xml_por_pasta(caminho_xml, info):
    """
    Move para:
      downloads/EMPRESA_PASTA/MM.AAAA/NFS_<NF>_<DD-MM-YYYY>.xml
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

    nf = (info.get("nf") or "").strip() or "SEM_NUMERO"
    destino_dir = os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA, competencia)
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
    return None

def ir_para_proxima_pagina():
    global ULTIMO_CHECKBOX_MARCADO
    assinatura_anterior = assinatura_lista()
    pag_antes = pagina_atual()

    botoes = driver.find_elements(
        By.XPATH,
        "//span[contains(@onclick,'mudarPagina,gridLista') and normalize-space(.)='Â»']"
    )
    if not botoes:
        return False

    botao_next = botoes[0]
    click_robusto(botao_next)

    try:
        esperar_troca_de_grid(assinatura_anterior, timeout=WAIT_LISTA_TIMEOUT)
        ULTIMO_CHECKBOX_MARCADO = None
        return True
    except TimeoutException:
        pag_depois = pagina_atual()
        if pag_antes is not None and pag_depois is not None and pag_depois > pag_antes:
            ULTIMO_CHECKBOX_MARCADO = None
            return True
        return False

# =====================
# PROCESSAR UMA NOTA (sem falso 502)
# =====================
# =====================
# SERVIÃ‡OS TOMADOS (DECLARAÃ‡ÃƒO FISCAL)
# =====================
def tomados_log_path_empresa():
    return os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA, TOMADOS_LOG_FILENAME)

def garantir_log_tomados_com_header(path_):
    """Agora garante um .txt simples para Tomados (uma linha por evento)."""
    os.makedirs(os.path.dirname(path_), exist_ok=True)
    if not os.path.exists(path_):
        with open(path_, "w", encoding="utf-8") as f:
            f.write("# Log Serviços Tomados (uma linha por evento)\n")

def salvar_log_tomados(status: str, referencia: str, arquivo_xml: str = "", mensagem: str = ""):
    path_ = tomados_log_path_empresa()
    garantir_log_tomados_com_header(path_)

    arquivo_base = os.path.basename(arquivo_xml) if arquivo_xml else ""
    pasta_xml = os.path.dirname(arquivo_xml) if arquivo_xml else ""

    parts = [
        f"STATUS={status}",
        f"REFERENCIA={referencia}",
    ]
    if arquivo_base:
        parts.append(f"ARQ={arquivo_base}")
    if pasta_xml:
        parts.append(f"PASTA={pasta_xml}")
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

def clicar_inicio_para_dashboard(timeout=30):
    """Volta segura para o dashboard usando o breadcrumb 'Início'."""
    try:
        if (
            len(driver.find_elements(By.ID, "imgdeclaracaofiscal")) > 0
            or len(driver.find_elements(By.ID, "divtxtdeclaracaofiscal")) > 0
        ):
            return
    except Exception:
        pass

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

    link_inicio = WebDriverWait(driver, timeout).until(_achar_link_inicio)
    click_robusto(link_inicio)

    # Checkpoint: dashboard precisa ter Declaração Fiscal
    WebDriverWait(driver, timeout).until(
        lambda d: (
            len(d.find_elements(By.ID, "imgdeclaracaofiscal")) > 0
            or len(d.find_elements(By.ID, "divtxtdeclaracaofiscal")) > 0
        )
    )

def abrir_declaracao_fiscal(timeout=30):
    # Preferência: IMG (mais "único"), fallback: DIV
    if driver.find_elements(By.ID, "imgdeclaracaofiscal"):
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "imgdeclaracaofiscal")))
        click_robusto(el)
        return
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "divtxtdeclaracaofiscal")))
    click_robusto(el)

def abrir_servicos_tomados(timeout=30):
    if driver.find_elements(By.ID, "imgtomadiss"):
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "imgtomadiss")))
        click_robusto(el)
        return
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "divtxttomadiss")))
    click_robusto(el)


def abrir_servicos_prestados(timeout=30):
    # tile Serviços Prestados
    if driver.find_elements(By.ID, "imgprestiss"):
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "imgprestiss")))
        click_robusto(el)
        return
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "divtxtprestiss")))
    click_robusto(el)


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


def entrar_contexto_livro(timeout=20):
    """Entra no contexto correto do livro: iframe externo e, se existir, frame interno 'inferior'."""
    def _achar_iframe():
        for idx, frame in enumerate(driver.find_elements(By.TAG_NAME, "iframe")):
            try:
                fid = (frame.get_attribute("id") or "").strip()
                fname = (frame.get_attribute("name") or "").strip()
                if not frame.is_displayed():
                    continue
                if fid == "_iFilho1" or fname == "_iFilho1" or "_iFilho" in fid or "_iFilho" in fname:
                    return {
                        "idx": idx,
                        "id": fid,
                        "name": fname,
                    }
            except Exception:
                continue
        return None

    fim = time.time() + timeout
    ultimo_erro = ""

    while time.time() < fim:
        driver.switch_to.default_content()
        iframe_info = _achar_iframe()
        if not iframe_info:
            sleep(0.20)
            continue

        try:
            outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
            driver.switch_to.frame(outer_frames[int(iframe_info["idx"])])

            try:
                controles = esperar_controles_livro(timeout=2)
                return {"iframe": iframe_info, "inner_frame": None, "controls": controles}
            except Exception as e:
                ultimo_erro = f"iframe-sem-controles: {type(e).__name__}"

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

            candidatos.sort(key=lambda item: 0 if item["id"] == "inferior" or item["name"] == "inferior" else 1)

            for inner in candidatos:
                try:
                    driver.switch_to.default_content()
                    outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
                    driver.switch_to.frame(outer_frames[int(iframe_info["idx"])])
                    inner_frames = driver.find_elements(By.TAG_NAME, "frame")
                    driver.switch_to.frame(inner_frames[int(inner["idx"])])
                    controles = esperar_controles_livro(timeout=2)
                    return {"iframe": iframe_info, "inner_frame": inner, "controls": controles}
                except Exception as e:
                    ultimo_erro = f"frame-interno-{inner.get('idx')}: {type(e).__name__}"
                    continue
        except Exception as e:
            ultimo_erro = f"iframe-erro: {type(e).__name__}"
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
    def _achar_iframe():
        for idx, frame in enumerate(driver.find_elements(By.TAG_NAME, "iframe")):
            try:
                fid = (frame.get_attribute("id") or "").strip()
                fname = (frame.get_attribute("name") or "").strip()
                if not frame.is_displayed():
                    continue
                if "_iFilho" in fid or "_iFilho" in fname:
                    return {
                        "idx": idx,
                        "id": fid,
                        "name": fname,
                    }
            except Exception:
                continue
        return None

    fim = time.time() + timeout
    ultimo_erro = ""

    while time.time() < fim:
        driver.switch_to.default_content()
        iframe_info = _achar_iframe()
        if not iframe_info:
            sleep(0.20)
            continue

        try:
            outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
            driver.switch_to.frame(outer_frames[int(iframe_info["idx"])])

            try:
                controles = esperar_controles_guia(timeout=2)
                return {"iframe": iframe_info, "inner_frame": None, "controls": controles}
            except Exception as e:
                ultimo_erro = f"iframe-sem-controles-guia: {type(e).__name__}"

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

            candidatos.sort(key=lambda item: 0 if item["id"] == "inferior" or item["name"] == "inferior" else 1)

            for inner in candidatos:
                try:
                    driver.switch_to.default_content()
                    outer_frames = driver.find_elements(By.TAG_NAME, "iframe")
                    driver.switch_to.frame(outer_frames[int(iframe_info["idx"])])
                    inner_frames = driver.find_elements(By.TAG_NAME, "frame")
                    driver.switch_to.frame(inner_frames[int(inner["idx"])])
                    controles = esperar_controles_guia(timeout=2)
                    return {"iframe": iframe_info, "inner_frame": inner, "controls": controles}
                except Exception as e:
                    ultimo_erro = f"frame-interno-guia-{inner.get('idx')}: {type(e).__name__}"
                    continue
        except Exception as e:
            ultimo_erro = f"iframe-guia-erro: {type(e).__name__}"
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
                sleep(0.25)

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
    log_manual = os.path.join(pasta_empresa(), "log_fechamento_manual.txt")
    modulo_up = (modulo or "").upper()
    handles_base = []
    try:
        driver.switch_to.default_content()
        handles_base = list(driver.window_handles)

        botao_livro = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((
                By.XPATH,
                f"//button[@type='button' and contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'livro') and contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'{termo_titulo.lower()}')]"
            ))
        )
        click_robusto(botao_livro)
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | BOTAO_LIVRO=OK | clicado")

        contexto = entrar_contexto_livro(timeout=20)
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
        visualizar_btn = modal["visualizar"]
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", visualizar_btn)
        WebDriverWait(driver, 5).until(lambda d: visualizar_btn.is_displayed() and visualizar_btn.is_enabled())
        click_robusto(visualizar_btn)

        try:
            WebDriverWait(driver, 2).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
        except Exception:
            pass

        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | VISUALIZAR=OK | EXERCICIO={ano_alvo} | MES={mes_alvo:02d}")

        pdf_baixado = ""
        try:
            pdf_baixado = aguardar_pdf_novo(TEMP_DOWNLOAD_DIR, timeout=45, pdfs_antes=pdfs_antes)
            append_txt(log_manual, f"{modulo_up} | {ref_alvo} | PDF_FETCH=OK | via arquivo _tmp")
        except Exception:
            pass

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
    competencia_dir = os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA, f"{mes_alvo:02d}.{ano_alvo}")
    os.makedirs(competencia_dir, exist_ok=True)

    mm_aaaa = f"{mes_alvo:02d}-{ano_alvo}"
    nome_final = TOMADOS_XML_BASENAME.format(mm_aaaa=mm_aaaa)
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
    return destino_final


def salvar_pdf_livro(pdf_tmp_path: str, modulo_up: str, ano_alvo: int, mes_alvo: int) -> str:
    competencia_dir = os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA, f"{mes_alvo:02d}.{ano_alvo}")
    os.makedirs(competencia_dir, exist_ok=True)

    mm_aaaa = f"{mes_alvo:02d}-{ano_alvo}"
    if (modulo_up or "").upper() == "TOMADOS":
        nome_final = LIVRO_TOMADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)
    else:
        nome_final = LIVRO_PRESTADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)

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
    competencia_dir = os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA, f"{mes_alvo:02d}.{ano_alvo}")
    os.makedirs(competencia_dir, exist_ok=True)

    mm_aaaa = f"{mes_alvo:02d}-{ano_alvo}"
    if (modulo_up or "").upper() == "TOMADOS":
        nome_final = GUIA_ISS_TOMADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)
    else:
        nome_final = GUIA_ISS_PRESTADOS_PDF_BASENAME.format(mm_aaaa=mm_aaaa)

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


def baixar_guia_iss_se_existir(modulo_up: str, ref_alvo: str, col_ref: str, ano_alvo: int, mes_alvo: int) -> str:
    """Se existir ícone de guia ISS na linha da referência, abre modal e baixa o PDF."""
    log_manual = os.path.join(pasta_empresa(), "log_fechamento_manual.txt")
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

        visualizar_btn = modal["visualizar"]
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", visualizar_btn)
        WebDriverWait(driver, 5).until(lambda d: visualizar_btn.is_displayed() and visualizar_btn.is_enabled())
        click_robusto(visualizar_btn)
        append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_VISUALIZAR=OK")

        try:
            WebDriverWait(driver, 2).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
        except Exception:
            pass

        pdf_baixado = ""
        try:
            pdf_baixado = aguardar_pdf_novo(TEMP_DOWNLOAD_DIR, timeout=45, pdfs_antes=pdfs_antes)
            append_txt(log_manual, f"{modulo_up} | {ref_alvo} | GUIA_PDF_FETCH=OK | via arquivo _tmp")
        except Exception:
            pass

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
    log_manual = os.path.join(pasta_empresa(), "log_fechamento_manual.txt")

    try:
        salvar_log_tomados("INICIO", ref_alvo, mensagem="Iniciando etapa Serviços Tomados")
    except Exception:
        pass

    def _obter_ref_topo(timeout=30) -> str:
        # Preferência: label com id fixo
        try:
            el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, "_label30")))
            return (el.text or "").strip()
        except Exception:
            pass
        # Fallback: encontra o label após o texto 'Referência'
        try:
            el = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, "//label[normalize-space()='Referência']/following::label[1]"))
            )
            return (el.text or "").strip()
        except Exception:
            return ""

    try:
        clicar_inicio_para_dashboard(timeout=40)

        abrir_declaracao_fiscal(timeout=40)
        if not declaracao_fiscal_tem_servicos_prestados():
            salvar_print_evidencia(
                "DECLARACAO_FISCAL_PRESTADOS",
                "SERVICOS_PRESTADOS_NAO_IDENTIFICADO",
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
                log_path=log_manual,
            )
        abrir_servicos_tomados(timeout=40)

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

        # Checkpoint crítico: referência no topo precisa bater
        # Checkpoint crítico: referência no topo precisa bater (normalizado)
        def _norm_ref(s: str) -> str:
            s = (s or "").replace("\u00a0", " ").strip()
            s = re.sub(r"\s+", "", s)
            m = re.match(r"^(\d{1,2})/(\d{4})$", s)
            if m:
                return f"{int(m.group(1)):02d}/{m.group(2)}"
            return s

        ref_topo = _norm_ref(_obter_ref_topo(timeout=50))
        if ref_topo != _norm_ref(ref_alvo):
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
                os.path.join(pasta_empresa(), "log_fechamento_manual.txt"),
                f"TOMADOS | {ref_alvo} | XML_FETCH_FALLBACK_CHROME | {type(e_fetch).__name__}: {str(e_fetch)[:120]}",
            )
            click_robusto(btn_download)
            xml_baixado = aguardar_xml_novo(TEMP_DOWNLOAD_DIR, timeout=120, xmls_antes=xmls_antes)

        destino_final = salvar_xml_tomados(xml_baixado, ano_alvo, mes_alvo)

        salvar_log_tomados("OK", ref_alvo, arquivo_xml=destino_final, mensagem="XML de Serviços Tomados baixado e organizado")
        print(f"[TOMADOS] OK -> {destino_final}")
        return "OK"

    except Exception as e:
        msg = str(e)
        try:
            salvar_log_tomados("ERRO", ref_alvo, mensagem=msg[:200])
        except Exception:
            pass
        print(f"[TOMADOS] ERRO: {msg}")
        if TOMADOS_OBRIGATORIO:
            raise RuntimeError(f"{MSG_TOMADOS_FALHA}: {msg}")
        return "ERRO"


def executar_cadeado_tomados(ano_alvo: int, mes_alvo: int) -> str:
    """Confirma o fechamento e abre o modal do livro em Serviços Tomados."""
    ref_alvo = f"{mes_alvo:02d}/{ano_alvo}"
    log_manual = os.path.join(pasta_empresa(), "log_fechamento_manual.txt")
    try:
        append_txt(log_manual, f"TOMADOS | {ref_alvo} | INICIO | Cadeado + Modal Livro")
        clicar_inicio_para_dashboard(timeout=40)
        abrir_declaracao_fiscal(timeout=40)
        if not declaracao_fiscal_tem_servicos_prestados():
            salvar_print_evidencia(
                "DECLARACAO_FISCAL_PRESTADOS",
                "SERVICOS_PRESTADOS_NAO_IDENTIFICADO",
                ano_alvo=ano_alvo,
                mes_alvo=mes_alvo,
                ref_alvo=ref_alvo,
                log_path=log_manual,
            )
        abrir_servicos_tomados(timeout=40)
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

        clicou = clicar_cadeado_na_linha(ref_alvo, col_ref)
        if clicou:
            append_txt(os.path.join(pasta_empresa(), "log_fechamento_manual.txt"), f"TOMADOS | {ref_alvo} | CADEADO=OK")
        else:
            append_txt(os.path.join(pasta_empresa(), "log_fechamento_manual.txt"), f"TOMADOS | {ref_alvo} | SEM_CADEADO | Cadeado não disponível (provável já fechado)")

        abrir_modal_livro_e_visualizar("TOMADOS", "tomados", ano_alvo, mes_alvo, ref_alvo)
        baixar_guia_iss_se_existir("TOMADOS", ref_alvo, col_ref, ano_alvo, mes_alvo)
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
    log_manual = os.path.join(pasta_empresa(), "log_fechamento_manual.txt")
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
        clicar_inicio_para_dashboard(timeout=40)
        abrir_declaracao_fiscal(timeout=40)
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

        clicou = clicar_cadeado_na_linha(ref_alvo, col_ref)
        if clicou:
            append_txt(log_manual, f"PRESTADOS | {ref_alvo} | CADEADO=OK")
        else:
            append_txt(log_manual, f"PRESTADOS | {ref_alvo} | SEM_CADEADO | Cadeado não disponível (provável já fechado)")

        abrir_modal_livro_e_visualizar("PRESTADOS", "prestados", ano_alvo, mes_alvo, ref_alvo)
        baixar_guia_iss_se_existir("PRESTADOS", ref_alvo, col_ref, ano_alvo, mes_alvo)
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
def processar_nota_por_indice(i, ano_alvo, mes_alvo):
    global PARAR_PROCESSAMENTO, ENCONTROU_MES_ALVO, CONT_FORA_APOS_ALVO, CONT_FORA_ANTES_ALVO, SEM_COMPETENCIA_NA_EMPRESA

    tentativa = 0
    while tentativa < MAX_RETRIES_POR_NOTA:
        tentativa += 1
        info = {}

        try:
            if is_502_page():
                recover_from_502()

            esperar_lista(timeout=WAIT_LISTA_TIMEOUT)

            checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
            if i >= len(checkboxes):
                return True

            checkbox = checkboxes[i]
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", checkbox)
            sleep(0.10)
            marcar_somente_checkbox(checkbox)

            info = extrair_info_linha(checkbox)
            key = chave_unica(info)

            # Avalia competência primeiro para encerrar cedo em ordem decrescente.
            comp = comparar_competencia_nota(info, ano_alvo, mes_alvo)
            if comp is None:
                msg = f"Data de emissao invalida/ausente: {info.get('data_emissao', '')}"
                print(f"[{i+1}] SKIP DATA INVALIDA -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                try:
                    salvar_log("", info, status="SKIP_DATA_INVALIDA", mensagem=msg[:180])
                except Exception:
                    pass
                return True

            if comp == 1:
                CONT_FORA_ANTES_ALVO = 0
                msg = f"Nota mais nova que mês alvo {mes_alvo:02d}/{ano_alvo}"
                print(f"[{i+1}] SKIP COMPETENCIA (MAIS NOVA) -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                try:
                    salvar_log("", info, status="SKIP_FORA_COMPETENCIA", mensagem=msg[:180])
                except Exception:
                    pass
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
                return True

            # comp == 0 (mês alvo)
            ENCONTROU_MES_ALVO = True
            CONT_FORA_APOS_ALVO = 0
            CONT_FORA_ANTES_ALVO = 0

            # PREMIUM: SKIP por log
            if key in chaves_ok:
                print(f"[{i+1}] SKIP -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                return True

            # Regra de negocio: não baixar notas canceladas.
            if nota_cancelada(info):
                msg = f"Nota cancelada (situacao: {info.get('situacao', '')})"
                print(f"[{i+1}] SKIP CANCELADA -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                try:
                    salvar_log("", info, status="SKIP_CANCELADA", mensagem=msg[:180])
                except Exception:
                    pass
                return True

            print(f"[{i+1}] Tentativa {tentativa}/{MAX_RETRIES_POR_NOTA} -> NF={info.get('nf')} RPS={info.get('rps')} DATA={info.get('data_emissao')}")

            xmls_antes = [a for a in os.listdir(TEMP_DOWNLOAD_DIR) if a.lower().endswith(".xml")]

            # botão exportar XML (ID fixo)
            botao_xml = wait.until(EC.element_to_be_clickable((By.ID, "_imagebutton12")))
            click_robusto(botao_xml)

            # aviso de "só 1 por vez"
            if fechar_aviso_se_existir(timeout=2):
                checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
                checkbox = checkboxes[i]
                marcar_somente_checkbox(checkbox, full_reset=True)
                botao_xml = wait.until(EC.element_to_be_clickable((By.ID, "_imagebutton12")))
                click_robusto(botao_xml)

            # modal exportação
            modal = wait.until(EC.visibility_of_element_located((By.ID, "modalboxexportarnotas")))
            select = Select(modal.find_element(By.TAG_NAME, "select"))
            select.select_by_visible_text("NF Nacional")
            sleep(0.15)

            visualizar_btn = modal.find_element(By.XPATH, ".//button[contains(.,'Visualizar')]")
            try:
                xml_baixado = baixar_arquivo_via_submit_form(
                    visualizar_btn,
                    timeout=120,
                    nome_fallback=f"{(info.get('nf') or 'NF')}.xml",
                )
            except Exception:
                click_robusto(visualizar_btn)
                # >>> NÃO checar 502 aqui. O critério é: BAIXOU XML NOVO.
                xml_baixado = aguardar_xml_novo(TEMP_DOWNLOAD_DIR, timeout=80, xmls_antes=xmls_antes)

            destino_final = organizar_xml_por_pasta(xml_baixado, info)

            # log OK + cache (isso evita repetir na mesma execução)
            logp = salvar_log(destino_final, info, status="OK", mensagem="OK (download confirmado por arquivo)")
            chaves_ok.add(key)

            print(f"OK -> {destino_final}")
            print(f"Log -> {logp}")

            # fechar modal (não pode quebrar o fluxo se falhar)
            fechar_modal_exportacao()

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

    print(f"Falhou apos {MAX_RETRIES_POR_NOTA} tentativas. Pulando nota {i+1}.")
    return False

# =====================
# MAIN
# =====================
def main():
    global PARAR_PROCESSAMENTO, ENCONTROU_MES_ALVO, CONT_FORA_APOS_ALVO, CONT_FORA_ANTES_ALVO, SEM_COMPETENCIA_NA_EMPRESA, ULTIMO_CHECKBOX_MARCADO, chaves_ok, PRESTADOS_SEM_MODULO

    inicializar_chrome()
    preencher_login_prefeitura_se_habilitado()
    # Carrega cache de chaves OK do log (idempotência)
    chaves_ok = carregar_chaves_ok(log_path_empresa())
    print(f"Chaves OK carregadas do log: {len(chaves_ok)}")

    ULTIMO_CHECKBOX_MARCADO = None

    ano_alvo, mes_alvo = calcular_mes_alvo(APURACAO_REFERENCIA)

    # =====================
    # MODO APENAS CADEADO (FECHAMENTO) + PAUSA MANUAL
    # =====================
    if MODO_APENAS_CADEADO:
        print("MODO_APENAS_CADEADO=1 -> Somente fechar período (cadeado).")
        if PAUSAR_ENTRE_MODULOS_CADEADO:
            print("PAUSAR_ENTRE_MODULOS_CADEADO=1 -> haverá pausa entre Tomados e Prestados.")
        else:
            print("PAUSAR_ENTRE_MODULOS_CADEADO=0 -> Tomados e Prestados serão executados em sequência, com pausa apenas no final.")
        executar_cadeado_tomados(ano_alvo, mes_alvo)
        if PRESTADOS_SEM_MODULO:
            print("Prestados sem módulo Nota Fiscal: etapa final de Prestados será ignorada.")
        else:
            executar_cadeado_prestados(ano_alvo, mes_alvo)
        pausa_final_livros_manual(f"Competência alvo: {mes_alvo:02d}/{ano_alvo}")
        print("MODO_APENAS_CADEADO finalizado. Encerrando execução.")
        return


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


    # FAST-PATH (Prestados): se as primeiras notas visíveis já forem mais antigas que o mês alvo,
    # assumimos a lista ordenada por data desc e encerramos Prestados como sem competência,
    # evitando waits/paginação desnecessários.
    if (not PRESTADOS_SEM_MODULO) and (not sem_competencia_inicial):
        try:
            checks0 = driver.find_elements(By.NAME, "gridListaCheck")
            if checks0:
                # Confirma com 1-2 linhas para evitar falso positivo se a ordenação estiver diferente.
                infos = [extrair_info_linha(checks0[0])]
                if len(checks0) >= 2:
                    infos.append(extrair_info_linha(checks0[1]))

                comps = [comparar_competencia_nota(info, ano_alvo, mes_alvo) for info in infos]
                # considera "mais antiga" somente quando todas as linhas avaliadas forem -1
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


    # Ajusta page size apenas se a lista NFSe possui checkboxes (evita timeout quando lista vier sem checkbox)
    if (not PRESTADOS_SEM_MODULO) and (not sem_competencia_inicial) and (not PARAR_PROCESSAMENTO) and len(driver.find_elements(By.NAME, "gridListaCheck")) > 0:
        definir_page_size(100)
    else:
        print("Page size não ajustado (lista sem checkbox ou sem notas na gridLista).")

    pagina = 1
    while True:
        if PARAR_PROCESSAMENTO:
            if PRESTADOS_SEM_MODULO:
                print("Prestados ignorado por ausência do módulo Nota Fiscal.")
            else:
                print("Prestados encerrado pelo fast-path (sem competência).")
            break

        total = len(driver.find_elements(By.NAME, "gridListaCheck"))
        print(f"Pagina {pagina}: {total} notas encontradas.")

        if total == 0 and not ENCONTROU_MES_ALVO:
            SEM_COMPETENCIA_NA_EMPRESA = True
            print("Lista carregada sem checkboxes; encerrando empresa como sem competência.")
            break

        i = 0
        while i < total:
            processar_nota_por_indice(i, ano_alvo, mes_alvo)
            if PARAR_PROCESSAMENTO:
                break
            i += 1

        if PARAR_PROCESSAMENTO:
            print("Encerrando varredura por heuristica de competência.")
            break

        if not ir_para_proxima_pagina():
            break
        pagina += 1

    # Etapa 2: Serviços Tomados (Declaração Fiscal)
    executar_etapa_servicos_tomados(ano_alvo, mes_alvo)

    # =====================
    # ACRESCENTO: CADEADO (Tomados + Prestados) + PAUSA FINAL (manual para baixar livros)
    # =====================
    if ACRESCENTAR_CADEADO_E_PAUSA_FINAL:
        if PRESTADOS_SEM_MODULO:
            print("ACRESCENTAR_CADEADO_E_PAUSA_FINAL=1 -> clicando cadeado vermelho somente em Tomados (Prestados indisponível) e pausando ao final.")
        else:
            print("ACRESCENTAR_CADEADO_E_PAUSA_FINAL=1 -> clicando cadeado vermelho em Tomados e Prestados (competência alvo) e pausando ao final.")
        try:
            executar_cadeado_tomados(ano_alvo, mes_alvo)
        except Exception as e:
            try:
                append_txt(os.path.join(pasta_empresa(), "log_fechamento_manual.txt"), f"TOMADOS | ERRO_CADEADO | {str(e)[:180]}")
            except Exception:
                pass
        if not PRESTADOS_SEM_MODULO:
            try:
                executar_cadeado_prestados(ano_alvo, mes_alvo)
            except Exception as e:
                try:
                    append_txt(os.path.join(pasta_empresa(), "log_fechamento_manual.txt"), f"PRESTADOS | ERRO_CADEADO | {str(e)[:180]}")
                except Exception:
                    pass

        # Pausa única no final para você baixar os livros manualmente
        pausa_final_livros_manual(f"Competência alvo: {mes_alvo:02d}/{ano_alvo}")
    if SEM_COMPETENCIA_NA_EMPRESA:
        raise RuntimeError(MSG_SEM_COMPETENCIA)

    print("Processo finalizado.")
    sleep(2)


try:
    main()
except RuntimeError as e:
    msg = str(e)
    print(msg)

    # Erros controlados (não são "forenses")
    if MSG_CAPTCHA_TIMEOUT in msg:
        sys.exit(EXIT_CODE_CAPTCHA_TIMEOUT)
    if MSG_CAPTCHA_INCORRETO in msg:
        sys.exit(EXIT_CODE_CAPTCHA_TIMEOUT)
    if MSG_SEM_COMPETENCIA in msg:
        sys.exit(EXIT_CODE_SEM_COMPETENCIA)
    if MSG_SEM_SERVICOS in msg:
        sys.exit(EXIT_CODE_SEM_SERVICOS)
    if MSG_CREDENCIAL_INVALIDA in msg:
        sys.exit(EXIT_CODE_CREDENCIAL_INVALIDA)
    if MSG_TOMADOS_FALHA in msg:
        sys.exit(EXIT_CODE_TOMADOS_FALHA)
    if MSG_CHROME_INIT_FALHA in msg:
        sys.exit(EXIT_CODE_CHROME_INIT_FALHA)

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
