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
import unicodedata
from datetime import datetime
import xml.etree.ElementTree as ET
import shutil

# =====================
# CONFIGURACAO
# =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

TEMP_DOWNLOAD_DIR = os.path.join(DOWNLOAD_DIR, "_tmp")
os.makedirs(TEMP_DOWNLOAD_DIR, exist_ok=True)



def append_txt(path: str, line: str):
    """Append a human-readable line to a .txt log (one event per line)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    safe = (line or "").replace("\n", " ").replace("\r", " ").strip()
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {safe}\n")
MAX_RETRIES_POR_NOTA = 3
WAIT_LISTA_TIMEOUT = 60
LOG_FILENAME = "log_downloads_nfse.txt"
TOMADOS_LOG_FILENAME = "log_tomados.txt"
PRESTADOS_LOG_FILENAME = "log_prestados.txt"
TOMADOS_XML_BASENAME = "SERVICOS_TOMADOS_{mm_aaaa}.xml"  # ex: SERVICOS_TOMADOS_01-2026.xml

# Se quiser travar a empresa SEM depender do texto do site, descomente:
# EMPRESA_PASTA_FORCADA = "H2_IMOBILIARIA"
EMPRESA_PASTA_FORCADA = os.environ.get("EMPRESA_PASTA_FORCADA", "")

def _default_apuracao_referencia() -> str:
    """Retorna MM/AAAA com base no mês anterior à data atual."""
    agora = datetime.now()
    ano = agora.year
    mes = agora.month - 1
    if mes == 0:
        mes = 12
        ano -= 1
    return f"{mes:02d}/{ano}"


# Competência alvo: por padrão, mês anterior ao mês atual (apuração).
# Pode sobrescrever com APURACAO_REFERENCIA=MM/AAAA (ex.: 03/2026 -> alvo 02/2026).
APURACAO_REFERENCIA = os.environ.get("APURACAO_REFERENCIA", _default_apuracao_referencia()).strip()

PARAR_PROCESSAMENTO = False
ENCONTROU_MES_ALVO = False
CONT_FORA_APOS_ALVO = 0
CONT_FORA_ANTES_ALVO = 0
SEM_COMPETENCIA_NA_EMPRESA = False
chaves_ok = set()  # cache das CHAVE_UNICA com STATUS=OK (carregado do log)
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
EXIT_CODE_EMPRESA_MULTIPLA = 61
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

prefs = {
    "download.default_directory": TEMP_DOWNLOAD_DIR,  # importante
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "plugins.always_open_pdf_externally": True,
}
options.add_experimental_option("prefs", prefs)

driver = None
wait = None
login_wait_seconds = int(os.environ.get("LOGIN_WAIT_SECONDS", "120"))


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



# =====================
# DETECTOR - EMPRESA MULTIPLA (Cadastros Relacionados)
# =====================
def tela_cadastros_relacionados() -> bool:
    try:
        if driver.find_elements(By.ID, "__grid1Table") and driver.find_elements(By.NAME, "_grid1Selected"):
            return True

        src = _norm_txt(driver.page_source or "")
        return ("cadastros relacionados" in src) and ("selecao do contribuinte" in src)
    except Exception:
        return False

def falhar_se_empresa_multipla(contexto: str = ""):
    if not tela_cadastros_relacionados():
        return
    try:
        salvar_kit_forense(f"{MSG_EMPRESA_MULTIPLA} {contexto}".strip(), None)
    except Exception:
        pass
    raise RuntimeError(MSG_EMPRESA_MULTIPLA)

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
    if not AUTO_LOGIN_PREFEITURA:
        return

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
        falhar_se_empresa_multipla('pos-login')
        if credencial_invalida_na_tela():
            raise RuntimeError(MSG_CREDENCIAL_INVALIDA)
        if captcha_incorreto_na_tela():
            raise RuntimeError(MSG_CAPTCHA_INCORRETO)
        if sem_modulo_nota_fiscal_no_dashboard():
            raise RuntimeError(MSG_SEM_SERVICOS)
        if driver.find_elements(By.ID, LOGIN_CARD_DASHBOARD):
            break
        sleep(0.4)
    else:
        raise RuntimeError(MSG_CAPTCHA_TIMEOUT)

    navegar_para_lista_nota_fiscal()


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
    falhar_se_empresa_multipla('dashboard-inicial')

    # 1) Dashboard inicial: clicar em Nota Fiscal.
    card_nf = wait.until(EC.element_to_be_clickable((By.ID, LOGIN_CARD_DASHBOARD)))
    click_robusto(card_nf)
    falhar_se_empresa_multipla('apos-card-nota-fiscal')

    # 2) Segundo dashboard: clicar no card Lista Nota Fiscais (id confirmado pelo escritório).
    try:
        card_lista = WebDriverWait(driver, 12).until(EC.element_to_be_clickable((By.ID, LOGIN_CARD_LISTA_NOTAS)))
        click_robusto(card_lista)
        falhar_se_empresa_multipla('apos-card-lista-notas')
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
            falhar_se_empresa_multipla('apos-fallback-lista-notas')
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
    falhar_se_empresa_multipla('esperar-lista')
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

def aguardar_arquivo_novo(pasta, extensao, timeout=80, antes=None):
    """
    Espera aparecer um arquivo novo (por nome OU mtime mais novo) no _tmp.
    extensao: ".pdf" ou ".xml"
    antes: dict {nome: mtime} ou None
    """
    fim = time.time() + timeout
    antes = dict(antes or {})

    while time.time() < fim:
        arquivos = list(os.scandir(pasta))

        if any(e.name.lower().endswith(".crdownload") for e in arquivos):
            sleep(0.4)
            continue

        candidato = None
        candidato_mtime = -1.0

        for entry in arquivos:
            nome = entry.name
            if not nome.lower().endswith(extensao):
                continue
            try:
                st = entry.stat()
                mtime = st.st_mtime
                if st.st_size <= 0:
                    continue
            except Exception:
                continue

            if (nome not in antes) or (mtime > float(antes.get(nome, 0)) + 0.01):
                if mtime > candidato_mtime:
                    candidato = nome
                    candidato_mtime = mtime

        if candidato:
            return os.path.join(pasta, candidato)

        sleep(0.4)

    raise TimeoutError(f"Download nao finalizou (arquivo {extensao} novo nao apareceu)")

def snapshot_mtime(pasta, extensao):
    out = {}
    for e in os.scandir(pasta):
        if e.name.lower().endswith(extensao):
            try:
                out[e.name] = e.stat().st_mtime
            except Exception:
                pass
    return out

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
    nome = (nome or "").strip().upper()
    nome = re.sub(r"[\/\.\-]", " ", nome)
    nome = re.sub(r"[^A-Z0-9\s]", " ", nome)
    tokens = [t for t in re.split(r"\s+", nome) if t]

    tokens = [t for t in tokens if t not in SUFIXOS_LEGAIS and t not in {"S", "A"}]

    siglas = [t for t in tokens if _SIGLA_ALFANUM_PATTERN.match(t)]
    resto = [t for t in tokens if t not in siglas]

    out = siglas + resto
    return "_".join(out) if out else "EMPRESA_DESCONHECIDA"

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


def prestados_log_path_empresa():
    return os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA, PRESTADOS_LOG_FILENAME)

def garantir_log_prestados_com_header(path_):
    os.makedirs(os.path.dirname(path_), exist_ok=True)
    if not os.path.exists(path_):
        with open(path_, "w", encoding="utf-8") as f:
            f.write("# Log Serviços Prestados - Fechamento (uma linha por evento)\n")

def salvar_log_prestados(status: str, referencia: str, arquivo: str = "", mensagem: str = ""):
    path_ = prestados_log_path_empresa()
    garantir_log_prestados_com_header(path_)

    arquivo_base = os.path.basename(arquivo) if arquivo else ""
    pasta = os.path.dirname(arquivo) if arquivo else ""

    parts = [
        f"STATUS={status}",
        f"REFERENCIA={referencia}",
    ]
    if arquivo_base:
        parts.append(f"ARQ={arquivo_base}")
    if pasta:
        parts.append(f"PASTA={pasta}")
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
    falhar_se_empresa_multipla('clicar-inicio')
    """Volta segura para o dashboard usando o breadcrumb 'Início'."""
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
    falhar_se_empresa_multipla('abrir-declaracao-fiscal')
    # Preferência: IMG (mais "único"), fallback: DIV
    if driver.find_elements(By.ID, "imgdeclaracaofiscal"):
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "imgdeclaracaofiscal")))
        click_robusto(el)
        return
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "divtxtdeclaracaofiscal")))
    click_robusto(el)

def abrir_servicos_tomados(timeout=30):
    falhar_se_empresa_multipla('abrir-servicos-tomados')
    if driver.find_elements(By.ID, "imgtomadiss"):
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "imgtomadiss")))
        click_robusto(el)
        return
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "divtxttomadiss")))
    click_robusto(el)

def abrir_servicos_prestados(timeout=30):
    falhar_se_empresa_multipla('abrir-servicos-prestados')
    if driver.find_elements(By.ID, "imgprestiss"):
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "imgprestiss")))
        click_robusto(el)
        return
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "divtxtprestiss")))
    click_robusto(el)

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

def encontrar_tr_por_referencia(ref_alvo: str, col_ref: str):
    celulas_ref = driver.find_elements(By.XPATH, f"//td[contains(@id,',{col_ref}_grid')]")
    for cel in celulas_ref:
        try:
            if (cel.text or "").strip() != ref_alvo:
                continue
            tr = cel.find_element(By.XPATH, "./ancestor::tr[1]")

            cel_ref_linha = tr.find_element(By.XPATH, f".//td[contains(@id,',{col_ref}_grid')]")
            if (cel_ref_linha.text or "").strip() != ref_alvo:
                continue

            return tr
        except Exception:
            continue
    return None

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

def salvar_pdf_tomados(pdf_tmp_path: str, ano_alvo: int, mes_alvo: int, prefixo: str) -> str:
    competencia_dir = os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA, f"{mes_alvo:02d}.{ano_alvo}")
    os.makedirs(competencia_dir, exist_ok=True)

    mm_aaaa = f"{mes_alvo:02d}.{ano_alvo}"
    nome_final = f"{prefixo}_{mm_aaaa}.pdf"
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

def executar_etapa_servicos_tomados(ano_alvo: int, mes_alvo: int) -> str:
    falhar_se_empresa_multipla('etapa-tomados')
    """Etapa adicional: Declaração Fiscal -> Serviços Tomados -> exportar XML do mês alvo.

    Retorna: "OK" | "SUCESSO_SEM_TOMADOS" | "ERRO"
    Se TOMADOS_OBRIGATORIO=1 e ocorrer ERRO, levanta RuntimeError(MSG_TOMADOS_FALHA).
    """
    ref_alvo = f"{mes_alvo:02d}/{ano_alvo}"

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
        falhar_se_empresa_multipla("tomados")

        abrir_declaracao_fiscal(timeout=40)
        falhar_se_empresa_multipla("tomados")
        abrir_servicos_tomados(timeout=40)
        falhar_se_empresa_multipla("tomados")

        # Aguarda a grid existir (pode estar vazia; não exige linhas)
        WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.ID, "_gridTable"))
        )

        # Se a grid estiver vazia (NENHUM REGISTRO...), já encerramos como sucesso sem tomados.
        if tomados_grid_sem_registros(timeout=25):
            salvar_log_tomados("SUCESSO_SEM_TOMADOS", ref_alvo, mensagem="Nenhum registro para apresentação")
            print("[TOMADOS] Nenhum registro para apresentação. SUCESSO_SEM_TOMADOS.")
            return "SUCESSO_SEM_TOMADOS"

        # Agora que sabemos que há linhas, mapeamos a coluna "Referência"
        col_ref = esperar_grid_declaracao_fiscal(timeout=50)

        btn_laranja = encontrar_botao_laranja_por_referencia(ref_alvo, col_ref=col_ref)
        if not btn_laranja:
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

        # Baixar XML (agora sim esperar o botão ficar clicável)
        try:
            btn_download = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "_imagebutton4")))
        except TimeoutException:
            btn_download = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[contains(@title,'arquivo XML') and contains(@title,'Exportar')]"
                ))
            )
        click_robusto(btn_download)

        xml_baixado = aguardar_xml_novo(TEMP_DOWNLOAD_DIR, timeout=120, xmls_antes=xmls_antes)

        destino_final = salvar_xml_tomados(xml_baixado, ano_alvo, mes_alvo)

        salvar_log_tomados("OK", ref_alvo, arquivo_xml=destino_final, mensagem="XML de Serviços Tomados baixado e organizado")
        print(f"[TOMADOS] OK -> {destino_final}")
        return "OK"

    except Exception as e:
        msg = str(e)
        if MSG_EMPRESA_MULTIPLA in msg:
            raise RuntimeError(MSG_EMPRESA_MULTIPLA)
        try:
            salvar_log_tomados("ERRO", ref_alvo, mensagem=msg[:200])
        except Exception:
            pass
        print(f"[TOMADOS] ERRO: {msg}")
        if TOMADOS_OBRIGATORIO:
            raise RuntimeError(f"{MSG_TOMADOS_FALHA}: {msg}")
        return "ERRO"


def executar_fechamento_servicos_tomados(ano_alvo: int, mes_alvo: int) -> str:
    """Fechamento contábil (Tomados): cadeado (se houver na competência), guia ISS (se houver) e PDF do livro.
    Regras:
      - NUNCA pode impedir o download do XML (por isso é chamado após a etapa de XML).
      - Se não houver linha da competência: SEM_MOVIMENTO e não gera PDF.
      - Se não houver cadeado: considerar já fechado e seguir.
      - Qualquer erro aqui: apenas loga e segue (não altera exit code), exceto EMPRESA_MULTIPLA.
    """
    falhar_se_empresa_multipla("fechamento-tomados")
    ref_alvo = f"{mes_alvo:02d}/{ano_alvo}"

    try:
        salvar_log_tomados("FECHAMENTO_INICIO", ref_alvo, mensagem="Iniciando fechamento Tomados (cadeado/guia/livro)")
    except Exception:
        pass

    try:
        clicar_inicio_para_dashboard(timeout=40)
        falhar_se_empresa_multipla("fechamento-tomados")

        abrir_declaracao_fiscal(timeout=40)
        falhar_se_empresa_multipla("fechamento-tomados")
        abrir_servicos_tomados(timeout=40)
        falhar_se_empresa_multipla("fechamento-tomados")

        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.ID, "_gridTable")))

        if tomados_grid_sem_registros(timeout=25):
            salvar_log_tomados("FECHAMENTO_SEM_MOVIMENTO", ref_alvo, mensagem="Nenhum registro para apresentação")
            return "SEM_MOVIMENTO"

        col_ref = esperar_grid_declaracao_fiscal(timeout=50)
        tr = encontrar_tr_por_referencia(ref_alvo, col_ref)
        if not tr:
            salvar_log_tomados("FECHAMENTO_SEM_MOVIMENTO", ref_alvo, mensagem="Referência não encontrada (sem movimento)")
            return "SEM_MOVIMENTO"

        # 1) Cadeado (somente na linha da competência)
        try:
            locks = tr.find_elements(By.XPATH, ".//span[contains(@class,'fa-lock') and contains(@title,'Fechar Movimento')]")
            if locks:
                click_robusto(locks[0])
                # Aguarda um dos sinais de atualização: cadeado sumir OU guia aparecer OU grid recarregar
                WebDriverWait(driver, 40).until(
                    lambda d: (
                        (encontrar_tr_por_referencia(ref_alvo, col_ref) is not None)
                        and (
                            len(encontrar_tr_por_referencia(ref_alvo, col_ref).find_elements(
                                By.XPATH, ".//span[contains(@class,'fa-print') and contains(@title,'Imprimir Guia')]"
                            )) > 0
                            or
                            len(encontrar_tr_por_referencia(ref_alvo, col_ref).find_elements(
                                By.XPATH, ".//span[contains(@class,'fa-lock') and contains(@title,'Fechar Movimento')]"
                            )) == 0
                        )
                    )
                )
                salvar_log_tomados("CADEADO_OK", ref_alvo, mensagem="Movimento fechado (ou já processado) na competência")
            else:
                salvar_log_tomados("SEM_CADEADO", ref_alvo, mensagem="Cadeado não disponível na competência (provavelmente já fechado)")
        except Exception as e:
            salvar_log_tomados("CADEADO_ERRO", ref_alvo, mensagem=str(e)[:200])

        # Re-obter TR após possíveis mudanças no DOM
        tr = encontrar_tr_por_referencia(ref_alvo, col_ref)

        # 2) Guia ISS (pode não existir)
        try:
            prints = tr.find_elements(By.XPATH, ".//span[contains(@class,'fa-print') and contains(@title,'Imprimir Guia')]") if tr else []
            if prints:
                antes_pdf = snapshot_mtime(TEMP_DOWNLOAD_DIR, ".pdf")
                click_robusto(prints[0])
                pdf_tmp = aguardar_arquivo_novo(TEMP_DOWNLOAD_DIR, ".pdf", timeout=120, antes=antes_pdf)
                guia_final = salvar_pdf_tomados(pdf_tmp, ano_alvo, mes_alvo, "GUIA_ISS_TOMADOS")
                salvar_log_tomados("GUIA_ISS_OK", ref_alvo, arquivo_xml=guia_final, mensagem="Guia ISS baixada")
            else:
                salvar_log_tomados("SEM_ISS", ref_alvo, mensagem="Guia ISS não disponível para a competência")
        except Exception as e:
            salvar_log_tomados("GUIA_ISS_ERRO", ref_alvo, mensagem=str(e)[:200])

        # 3) Livro PDF (modal de parâmetros)
        try:
            btn_livro = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "_imagebutton13")))
            click_robusto(btn_livro)

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "exercicio")))
            Select(driver.find_element(By.ID, "exercicio")).select_by_value(str(ano_alvo))
            Select(driver.find_element(By.ID, "mes")).select_by_value(str(mes_alvo))

            antes_pdf = snapshot_mtime(TEMP_DOWNLOAD_DIR, ".pdf")
            click_robusto(driver.find_element(By.ID, "_imagebutton1"))
            pdf_tmp = aguardar_arquivo_novo(TEMP_DOWNLOAD_DIR, ".pdf", timeout=120, antes=antes_pdf)
            livro_final = salvar_pdf_tomados(pdf_tmp, ano_alvo, mes_alvo, "LIVRO_TOMADOS")
            salvar_log_tomados("LIVRO_OK", ref_alvo, arquivo_xml=livro_final, mensagem="Livro de tomados baixado")
        except Exception as e:
            salvar_log_tomados("LIVRO_ERRO", ref_alvo, mensagem=str(e)[:200])

        salvar_log_tomados("FECHAMENTO_OK", ref_alvo, mensagem="Fechamento Tomados finalizado")
        return "OK"

    except Exception as e:
        msg = str(e)
        if MSG_EMPRESA_MULTIPLA in msg:
            raise RuntimeError(MSG_EMPRESA_MULTIPLA)
        try:
            salvar_log_tomados("FECHAMENTO_ERRO", ref_alvo, mensagem=msg[:250])
        except Exception:
            pass
        return "ERRO"



def executar_etapa_servicos_prestados(ano_alvo: int, mes_alvo: int) -> str:
    """Fechamento contábil (Prestados): cadeado (se houver na competência) e PDF do livro.
    Regras:
      - Se não houver linha da competência: SUCESSO_SEM_PRESTADOS e não gera PDF.
      - Se não houver cadeado: considerar já fechado e seguir.
      - Qualquer erro aqui: apenas loga e segue (não altera exit code), exceto EMPRESA_MULTIPLA.
    """
    falhar_se_empresa_multipla("fechamento-prestados")
    ref_alvo = f"{mes_alvo:02d}/{ano_alvo}"

    try:
        salvar_log_prestados("INICIO", ref_alvo, mensagem="Iniciando fechamento Prestados (cadeado/livro)")
    except Exception:
        pass

    try:
        clicar_inicio_para_dashboard(timeout=40)
        falhar_se_empresa_multipla("fechamento-prestados")

        abrir_declaracao_fiscal(timeout=40)
        falhar_se_empresa_multipla("fechamento-prestados")
        abrir_servicos_prestados(timeout=40)
        falhar_se_empresa_multipla("fechamento-prestados")

        WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.ID, "_gridTable")))

        if tomados_grid_sem_registros(timeout=25):
            salvar_log_prestados("SUCESSO_SEM_PRESTADOS", ref_alvo, mensagem="Nenhum registro para apresentação")
            return "SUCESSO_SEM_PRESTADOS"

        col_ref = esperar_grid_declaracao_fiscal(timeout=50)
        tr = encontrar_tr_por_referencia(ref_alvo, col_ref)
        if not tr:
            salvar_log_prestados("SUCESSO_SEM_PRESTADOS", ref_alvo, mensagem="Referência não encontrada (sem movimento)")
            return "SUCESSO_SEM_PRESTADOS"

        # 1) Cadeado (somente na linha da competência)
        try:
            locks = tr.find_elements(By.XPATH, ".//span[contains(@class,'fa-lock') and contains(@title,'Fechar Movimento')]")
            if locks:
                click_robusto(locks[0])
                WebDriverWait(driver, 40).until(
                    lambda d: (
                        (encontrar_tr_por_referencia(ref_alvo, col_ref) is not None)
                        and (
                            len(encontrar_tr_por_referencia(ref_alvo, col_ref).find_elements(
                                By.XPATH, ".//span[contains(@class,'fa-lock') and contains(@title,'Fechar Movimento')]"
                            )) == 0
                        )
                    )
                )
                salvar_log_prestados("CADEADO_OK", ref_alvo, mensagem="Movimento fechado (ou já processado) na competência")
            else:
                salvar_log_prestados("SEM_CADEADO", ref_alvo, mensagem="Cadeado não disponível na competência (provavelmente já fechado)")
        except Exception as e:
            salvar_log_prestados("CADEADO_ERRO", ref_alvo, mensagem=str(e)[:200])

        # 2) Livro PDF (modal de parâmetros)
        try:
            btn_livro = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "_imagebutton13")))
            click_robusto(btn_livro)

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "exercicio")))
            Select(driver.find_element(By.ID, "exercicio")).select_by_value(str(ano_alvo))
            Select(driver.find_element(By.ID, "mes")).select_by_value(str(mes_alvo))

            antes_pdf = snapshot_mtime(TEMP_DOWNLOAD_DIR, ".pdf")
            click_robusto(driver.find_element(By.ID, "_imagebutton1"))
            pdf_tmp = aguardar_arquivo_novo(TEMP_DOWNLOAD_DIR, ".pdf", timeout=120, antes=antes_pdf)
            livro_final = salvar_pdf_tomados(pdf_tmp, ano_alvo, mes_alvo, "LIVRO_PRESTADOS")
            salvar_log_prestados("LIVRO_OK", ref_alvo, arquivo=livro_final, mensagem="Livro de prestados baixado")
        except Exception as e:
            salvar_log_prestados("LIVRO_ERRO", ref_alvo, mensagem=str(e)[:200])

        salvar_log_prestados("FECHAMENTO_OK", ref_alvo, mensagem="Fechamento Prestados finalizado")
        return "OK"

    except Exception as e:
        msg = str(e)
        if MSG_EMPRESA_MULTIPLA in msg:
            raise RuntimeError(MSG_EMPRESA_MULTIPLA)
        try:
            salvar_log_prestados("ERRO", ref_alvo, mensagem=msg[:250])
        except Exception:
            pass
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
    global PARAR_PROCESSAMENTO, ENCONTROU_MES_ALVO, CONT_FORA_APOS_ALVO, CONT_FORA_ANTES_ALVO, SEM_COMPETENCIA_NA_EMPRESA, ULTIMO_CHECKBOX_MARCADO, chaves_ok

    inicializar_chrome()
    # Login: automático (CNPJ/senha) ou manual
    if AUTO_LOGIN_PREFEITURA:
        preencher_login_prefeitura_se_habilitado()
    else:
        driver.get(LOGIN_URL_PREFEITURA)
        print("Entre manualmente em: Nota Fiscal -> Lista Nota Fiscais")
        print(f"Você tem {login_wait_seconds} segundos para iniciar manualmente.")
        fim_manual = time.time() + login_wait_seconds
        while time.time() < fim_manual:
            falhar_se_empresa_multipla('manual-pos-login')
            if lista_ativa() or driver.find_elements(By.ID, 'gridListaPageSize'):
                break
            sleep(0.4)
    # Carrega cache de chaves OK do log (idempotência)
    chaves_ok = carregar_chaves_ok(log_path_empresa())
    print(f"Chaves OK carregadas do log: {len(chaves_ok)}")

    ULTIMO_CHECKBOX_MARCADO = None

    ano_alvo, mes_alvo = calcular_mes_alvo(APURACAO_REFERENCIA)

    sem_competencia_inicial = False

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
    PARAR_PROCESSAMENTO = False
    print(f"Mes alvo de download: {mes_alvo:02d}/{ano_alvo}")
    print(f"Heuristica de parada: {LIMITE_HEURISTICA_FORA_ALVO} notas antigas consecutivas (antes ou apos mês alvo).")


    # FAST-PATH (Prestados): se as primeiras notas visíveis já forem mais antigas que o mês alvo,
    # assumimos a lista ordenada por data desc e encerramos Prestados como sem competência,
    # evitando waits/paginação desnecessários.
    if not sem_competencia_inicial:
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
    if (not sem_competencia_inicial) and (not PARAR_PROCESSAMENTO) and len(driver.find_elements(By.NAME, "gridListaCheck")) > 0:
        definir_page_size(100)
    else:
        print("Page size não ajustado (lista sem checkbox ou sem notas na gridLista).")

    pagina = 1
    while True:
        if PARAR_PROCESSAMENTO:
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
    tomados_status = executar_etapa_servicos_tomados(ano_alvo, mes_alvo)
    if tomados_status == "ERRO":
        print("ALERTA_TOMADOS")

    # Fechamento (não bloqueante): qualquer erro vira log e segue como FECHADO
    try:
        fechamento_tomados_status = executar_fechamento_servicos_tomados(ano_alvo, mes_alvo)
        if fechamento_tomados_status == "ERRO":
            print("ALERTA_FECHAMENTO_TOMADOS")
    except Exception as e:
        if MSG_EMPRESA_MULTIPLA in str(e):
            raise
        print("ALERTA_FECHAMENTO_TOMADOS")
        try:
            salvar_log_tomados("FECHAMENTO_ERRO", f"{mes_alvo:02d}/{ano_alvo}", mensagem=str(e)[:200])
        except Exception:
            pass

    try:
        prestados_status = executar_etapa_servicos_prestados(ano_alvo, mes_alvo)
        if prestados_status == "ERRO":
            print("ALERTA_PRESTADOS")
    except Exception as e:
        if MSG_EMPRESA_MULTIPLA in str(e):
            raise
        print("ALERTA_PRESTADOS")
        try:
            salvar_log_prestados("ERRO", f"{mes_alvo:02d}/{ano_alvo}", mensagem=str(e)[:200])
        except Exception:
            pass
    if SEM_COMPETENCIA_NA_EMPRESA:
        raise RuntimeError(MSG_SEM_COMPETENCIA)

    print("Processo finalizado.")
    sleep(2)


def map_runtime_error_to_exit_code(msg: str):
    msg = str(msg or "")
    if MSG_CAPTCHA_TIMEOUT in msg:
        return EXIT_CODE_CAPTCHA_TIMEOUT
    if MSG_CAPTCHA_INCORRETO in msg:
        return EXIT_CODE_CAPTCHA_TIMEOUT
    if MSG_SEM_COMPETENCIA in msg:
        return EXIT_CODE_SEM_COMPETENCIA
    if MSG_SEM_SERVICOS in msg:
        return EXIT_CODE_SEM_SERVICOS
    if MSG_CREDENCIAL_INVALIDA in msg:
        return EXIT_CODE_CREDENCIAL_INVALIDA
    if MSG_EMPRESA_MULTIPLA in msg:
        return EXIT_CODE_EMPRESA_MULTIPLA
    if MSG_TOMADOS_FALHA in msg:
        return EXIT_CODE_TOMADOS_FALHA
    if MSG_CHROME_INIT_FALHA in msg:
        return EXIT_CODE_CHROME_INIT_FALHA
    return None


if __name__ == "__main__":
    try:
        main()
    except RuntimeError as e:
        msg = str(e)
        print(msg)

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
