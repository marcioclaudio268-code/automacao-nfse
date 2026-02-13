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

MAX_RETRIES_POR_NOTA = 3
WAIT_LISTA_TIMEOUT = 60
LOG_FILENAME = "log_downloads_nfse.csv"

EMPRESA_PASTA_FORCADA = os.environ.get("EMPRESA_PASTA_FORCADA", "")
APURACAO_REFERENCIA = os.environ.get("APURACAO_REFERENCIA", "").strip()

PARAR_PROCESSAMENTO = False
ENCONTROU_MES_ALVO = False
CONT_FORA_APOS_ALVO = 0
CONT_FORA_ANTES_ALVO = 0
CONT_DATA_INVALIDA = 0
SEM_COMPETENCIA_NA_EMPRESA = False
LIMITE_HEURISTICA_FORA_ALVO = int(os.environ.get("LIMITE_HEURISTICA_FORA_ALVO", "1"))
LIMITE_DATA_INVALIDA = int(os.environ.get("LIMITE_DATA_INVALIDA", "10"))
STRICT_LISTA_INICIAL = os.environ.get("STRICT_LISTA_INICIAL", "0").strip() == "1"
MSG_CAPTCHA_TIMEOUT = "CAPTCHA_NAO_RESOLVIDO_NO_TEMPO"
MSG_SEM_COMPETENCIA = "SUCESSO_SEM_COMPETENCIA"
EXIT_CODE_CAPTCHA_TIMEOUT = 30
EXIT_CODE_SEM_COMPETENCIA = 40

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

# =====================
# CHROME – PERFIL EXCLUSIVO
# =====================
options = Options()
options.add_argument("--start-maximized")

DEFAULT_PROFILE_DIR = (
    r"C:\ChromeRobotProfile" if os.name == "nt"
    else os.path.join(tempfile.gettempdir(), "ChromeRobotProfile")
)
CHROME_PROFILE_DIR = os.environ.get("CHROME_PROFILE_DIR", DEFAULT_PROFILE_DIR)
options.add_argument(f"--user-data-dir={CHROME_PROFILE_DIR}")

options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
options.add_experimental_option('useAutomationExtension', False)
options.add_argument("--disable-blink-features=AutomationControlled")

prefs = {
    "download.default_directory": TEMP_DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)

try:
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            });
        """
    })
except Exception:
    pass

wait = WebDriverWait(driver, 30)

print("Chrome iniciado.")
login_wait_seconds = int(os.environ.get("LOGIN_WAIT_SECONDS", "120"))


def verificar_modulo_nfse_existe():
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "imgnotafiscal"))
        )
        print("Modulo Nota Fiscal encontrado (via imgnotafiscal)")
        return True
    except TimeoutException:
        pass
    
    try:
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "divtxtnotafiscal"))
        )
        print("Modulo Nota Fiscal encontrado (via divtxtnotafiscal)")
        return True
    except TimeoutException:
        pass
    
    try:
        elementos_nf = driver.find_elements(By.XPATH, "//*[contains(@id, 'notafiscal')]")
        if len(elementos_nf) > 0:
            print(f"Modulo Nota Fiscal encontrado (via xpath - {len(elementos_nf)} elementos)")
            return True
    except Exception:
        pass
    
    print("Modulo Nota Fiscal NAO encontrado")
    return False


def preencher_login_prefeitura_se_habilitado():
    if not AUTO_LOGIN_PREFEITURA:
        return

    cnpj = re.sub(r"\D", "", os.environ.get("EMPRESA_CNPJ", ""))
    senha = os.environ.get("EMPRESA_SENHA", "")
    if not cnpj or not senha:
        raise RuntimeError("AUTO_LOGIN_PREFEITURA ativo, mas EMPRESA_CNPJ/EMPRESA_SENHA nao informados")

    driver.get(LOGIN_URL_PREFEITURA)
    campo_usuario = wait.until(EC.presence_of_element_located((By.ID, LOGIN_CAMPO_USUARIO)))
    campo_senha = wait.until(EC.presence_of_element_located((By.ID, LOGIN_CAMPO_SENHA)))

    campo_usuario.clear()
    campo_usuario.send_keys(cnpj)
    campo_senha.clear()
    campo_senha.send_keys(senha)

    wait.until(EC.element_to_be_clickable((By.ID, LOGIN_BOTAO_ENTRAR)))

    print("CNPJ e senha preenchidos. Resolva o captcha e clique em 'Entrar' manualmente.")
    print(f"Aguardando dashboard por ate {login_wait_seconds}s...")

    try:
        WebDriverWait(driver, login_wait_seconds).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, "img[id^='img']")) > 0
        )
        print("Dashboard carregado.")
    except TimeoutException:
        raise RuntimeError(MSG_CAPTCHA_TIMEOUT)

    print("Verificando se empresa possui modulo de Nota Fiscal...")
    if not verificar_modulo_nfse_existe():
        print("=" * 80)
        print("MODULO NOTA FISCAL NAO ENCONTRADO")
        print("=" * 80)
        print("Empresa nao possui modulo de NFSe (nao presta servicos).")
        print("Encerrando como SUCESSO SEM COMPETENCIA.")
        print("=" * 80)
        
        try:
            log_p = log_path_empresa()
            garantir_log_com_header(log_p)
            with open(log_p, "a", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow([
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "SEM_MODULO_NFSE",
                    "", "", "", "", "", "", "", "",
                    "Empresa sem modulo Nota Fiscal (nao presta servicos)"
                ])
            print(f"Registro salvo no log: {log_p}")
        except Exception as e:
            print(f"Erro ao salvar log: {e}")
        
        raise RuntimeError(MSG_SEM_COMPETENCIA)

    print("Modulo Nota Fiscal encontrado. Navegando para lista...")
    navegar_para_lista_nota_fiscal()


if not AUTO_LOGIN_PREFEITURA:
    print("Entre manualmente em: Nota Fiscal - Lista Nota Fiscais")
    print(f"Voce tem {login_wait_seconds} segundos para iniciar manualmente.")
    sleep(login_wait_seconds)

# =====================
# HELPERS
# =====================
def click_robusto(el):
    try:
        el.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", el)


def navegar_para_lista_nota_fiscal():
    print("Navegando: Nota Fiscal -> Lista Nota Fiscais...")

    try:
        card_nf = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, LOGIN_CARD_DASHBOARD))
        )
        click_robusto(card_nf)
    except TimeoutException:
        try:
            img_nf = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "imgnotafiscal"))
            )
            click_robusto(img_nf)
        except Exception as e:
            raise RuntimeError(f"Nao foi possivel clicar no card Nota Fiscal: {e}")

    try:
        card_lista = WebDriverWait(driver, 12).until(
            EC.element_to_be_clickable((By.ID, LOGIN_CARD_LISTA_NOTAS))
        )
        click_robusto(card_lista)
        
        status_lista = esperar_lista_ou_sem_checkbox(timeout=WAIT_LISTA_TIMEOUT)
        if status_lista.startswith("sem_checkbox"):
            print("Tela de lista carregada sem checkbox; empresa sera concluida sem competencia.")
        else:
            print("Tela de lista de notas carregada.")
        return
        
    except Exception as e:
        ultimo_erro = e

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
                print("Tela de lista carregada sem checkbox (fallback).")
            else:
                print("Tela de lista de notas carregada (fallback).")
            return
        except Exception as e:
            ultimo_erro = e

    raise RuntimeError(f"NAO_FOI_POSSIVEL_ABRIR_LISTA_NOTAS: {ultimo_erro}")


def esperar_lista(timeout=WAIT_LISTA_TIMEOUT):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located((By.NAME, "gridListaCheck"))
    )

def esperar_lista_ou_sem_checkbox(timeout=WAIT_LISTA_TIMEOUT):
    def cond(_):
        checks = driver.find_elements(By.NAME, "gridListaCheck")
        if len(checks) > 0:
            return "checkboxes"

        radios_linha = driver.find_elements(By.NAME, "gridListaSelected")
        celulas_linha = driver.find_elements(By.CSS_SELECTOR, "td[id*=',-1_gridLista']")
        datas = driver.find_elements(By.CSS_SELECTOR, "td[id*=',11_gridLista']")  # ← CORRIGIDO: 10 -> 11
        
        if len(datas) > 0 and len(checks) == 0:
            return "sem_checkbox_com_data"

        if (len(radios_linha) > 0 or len(celulas_linha) > 0) and len(checks) == 0:
            return "sem_checkbox"

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
    try:
        if lista_ativa():
            return False

        title = (driver.title or "").lower()
        src = (driver.page_source or "").lower()

        if ("bad gateway" in title and "502" in title):
            return True
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

def fechar_aviso_se_existir(timeout=2):
    try:
        modal = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((
                By.XPATH,
                "//*[contains(.,'ATENÇÃO!') and contains(.,'apenas a exportação de uma nota')]"
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

def td_text(tr, col):
    el = tr.find_element(By.CSS_SELECTOR, f"td[id*=',{col}_gridLista']")
    return (el.text or "").strip()

def primeira_data_emissao_visivel_sem_checkbox():
    try:
        el = driver.find_element(By.CSS_SELECTOR, "td[id*=',11_gridLista']")  # ← CORRIGIDO: 10 -> 11
        return (el.text or "").strip()
    except Exception:
        return ""

def extrair_info_linha(checkbox):
    tr = checkbox.find_element(By.XPATH, "./ancestor::tr[1]")
    return {
        "id_interno": checkbox.get_attribute("value") or "",
        "nf": td_text(tr, 6),
        "situacao": td_text(tr, 9),
        "data_emissao": td_text(tr, 11),  # ← CORRIGIDO: 10 -> 11
        "rps": td_text(tr, 12),            # ← AJUSTADO: 11 -> 12 (RPS também mudou!)
        "chave": td_text(tr, 14),
    }

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

def parse_data_emissao_site(data_emissao: str):
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", data_emissao or "")
    if not m:
        return None
    dd, mm, yyyy = m.group(1), m.group(2), m.group(3)
    iso = f"{yyyy}-{mm}-{dd}"
    competencia = f"{mm}.{yyyy}"
    return yyyy, mm, dd, iso, competencia

def calcular_mes_alvo(apuracao_ref: str = ""):
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

def aguardar_xml_novo(pasta, timeout=80, xmls_antes=None):
    fim = time.time() + timeout
    xmls_antes = set(xmls_antes or [])

    while time.time() < fim:
        arquivos = os.listdir(pasta)

        if any(a.endswith(".crdownload") for a in arquivos):
            sleep(0.4)
            continue

        xmls = [a for a in arquivos if a.lower().endswith(".xml")]
        novos = [x for x in xmls if x not in xmls_antes]

        if novos:
            novos.sort(
                key=lambda x: os.path.getmtime(os.path.join(pasta, x)),
                reverse=True
            )
            return os.path.join(pasta, novos[0])

        sleep(0.4)

    raise TimeoutError("Download nao finalizou (xml novo nao apareceu)")

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

def log_path_empresa():
    return os.path.join(DOWNLOAD_DIR, EMPRESA_PASTA, LOG_FILENAME)

def garantir_log_com_header(path_):
    header = [
        "timestamp",
        "status",
        "chave_unica",
        "nf",
        "data_emissao",
        "rps",
        "chave",
        "id_interno",
        "arquivo_xml",
        "pasta_xml",
        "mensagem",
    ]
    if not os.path.exists(path_):
        os.makedirs(os.path.dirname(path_), exist_ok=True)
        with open(path_, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(header)

def carregar_chaves_ok(path_):
    chaves = set()
    if not os.path.exists(path_):
        return chaves
    try:
        with open(path_, "r", newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=";")
            _ = next(reader, None)
            for row in reader:
                if not row or len(row) < 3:
                    continue
                status = (row[1] or "").strip().upper()
                key = (row[2] or "").strip()
                if status == "OK" and key:
                    chaves.add(key)
    except Exception:
        pass
    return chaves

def salvar_log(destino_xml, info, status="OK", mensagem=""):
    path_ = log_path_empresa()
    garantir_log_com_header(path_)

    key = chave_unica(info)
    arquivo_xml = os.path.basename(destino_xml) if destino_xml and os.path.isfile(destino_xml) else ""
    pasta_xml = os.path.dirname(destino_xml) if destino_xml and os.path.isfile(destino_xml) else ""

    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        status,
        key,
        info.get("nf", ""),
        info.get("data_emissao", ""),
        info.get("rps", ""),
        info.get("chave", ""),
        info.get("id_interno", ""),
        arquivo_xml,
        pasta_xml,
        mensagem,
    ]

    with open(path_, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(row)

    return path_

chaves_ok = carregar_chaves_ok(log_path_empresa())
print(f"Chaves OK no log: {len(chaves_ok)}")
print(f"Log: {log_path_empresa()}")

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
        novo_nome = f"NFS_{nf}_{data_iso}.xml"
    else:
        novo_nome = f"NFS_{nf}.xml"

    destino_final = os.path.join(destino_dir, novo_nome)

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
    assinatura_anterior = assinatura_lista()
    pag_antes = pagina_atual()

    botoes = driver.find_elements(
        By.XPATH,
        "//span[contains(@onclick,'mudarPagina,gridLista') and normalize-space(.)='»']"
    )
    if not botoes:
        return False

    botao_next = botoes[0]
    click_robusto(botao_next)

    try:
        esperar_troca_de_grid(assinatura_anterior, timeout=WAIT_LISTA_TIMEOUT)
        return True
    except TimeoutException:
        pag_depois = pagina_atual()
        if pag_antes is not None and pag_depois is not None and pag_depois > pag_antes:
            return True
        return False


def verificar_primeira_pagina_tem_mes_alvo(ano_alvo, mes_alvo):
    print(f"Verificando se primeira pagina contem notas do mes alvo {mes_alvo:02d}/{ano_alvo}...")
    
    try:
        checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
        
        if len(checkboxes) == 0:
            print("[AVISO] Nenhum checkbox encontrado.")
            return False
        
        notas_no_alvo = 0
        notas_antigas = 0
        notas_novas = 0
        datas_invalidas = 0
        
        sample_size = min(10, len(checkboxes))
        
        for i in range(sample_size):
            try:
                checkbox = checkboxes[i]
                info = extrair_info_linha(checkbox)
                comp = comparar_competencia_nota(info, ano_alvo, mes_alvo)
                
                data_str = info.get('data_emissao', 'SEM_DATA')
                
                if comp == 0:
                    notas_no_alvo += 1
                    print(f"  [{i+1}] NO ALVO: {data_str}")
                elif comp == -1:
                    notas_antigas += 1
                    print(f"  [{i+1}] ANTIGA: {data_str}")
                elif comp == 1:
                    notas_novas += 1
                    print(f"  [{i+1}] NOVA: {data_str}")
                else:
                    datas_invalidas += 1
                    print(f"  [{i+1}] INVALIDA: {data_str}")
                    
            except Exception as e:
                print(f"[DEBUG] Erro ao verificar nota {i}: {e}")
                continue
        
        print(f"[SAMPLE] Novas: {notas_novas} | No alvo: {notas_no_alvo} | Antigas: {notas_antigas} | Invalidas: {datas_invalidas}")
        
        if datas_invalidas == sample_size or (notas_antigas > 0 and notas_no_alvo == 0 and notas_novas == 0):
            print("=" * 80)
            print("[DETECCAO RAPIDA] Empresa NAO possui notas validas na competencia alvo.")
            print("Encerrando como SEM COMPETENCIA.")
            print("=" * 80)
            return False
        
        if notas_no_alvo > 0:
            print(f"[OK] Encontradas {notas_no_alvo} notas no mes alvo. Processando...")
            return True
        
        if notas_novas > 0:
            print(f"[OK] Encontradas {notas_novas} notas novas. Continuando busca...")
            return True
        
        return True
        
    except Exception as e:
        print(f"[ERRO] Erro ao verificar primeira pagina: {e}")
        return True


def processar_nota_por_indice(i, ano_alvo, mes_alvo):
    global PARAR_PROCESSAMENTO, ENCONTROU_MES_ALVO, CONT_FORA_APOS_ALVO, CONT_FORA_ANTES_ALVO, CONT_DATA_INVALIDA, SEM_COMPETENCIA_NA_EMPRESA

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

            desmarcar_todas_notas()

            checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
            checkbox = checkboxes[i]

            info = extrair_info_linha(checkbox)
            key = chave_unica(info)

            if key in chaves_ok:
                print(f"[{i+1}] SKIP -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                return True

            if nota_cancelada(info):
                msg = f"Nota cancelada (situacao: {info.get('situacao', '')})"
                print(f"[{i+1}] SKIP CANCELADA -> NF={info.get('nf')} DATA={info.get('data_emissao')}")
                try:
                    salvar_log("", info, status="SKIP_CANCELADA", mensagem=msg[:180])
                except Exception:
                    pass
                return True

            comp = comparar_competencia_nota(info, ano_alvo, mes_alvo)
            
            data_str = info.get('data_emissao', 'SEM_DATA')
            comp_str = "ALVO" if comp == 0 else ("NOVA" if comp == 1 else ("ANTIGA" if comp == -1 else "INVALIDA"))
            print(f"[{i+1}] {comp_str}: {data_str} (alvo: {mes_alvo:02d}/{ano_alvo})")
            
            if comp is None:
                CONT_DATA_INVALIDA += 1
                print(f"  >> Contador datas invalidas: {CONT_DATA_INVALIDA}/{LIMITE_DATA_INVALIDA}")
                
                if CONT_DATA_INVALIDA >= LIMITE_DATA_INVALIDA:
                    print("=" * 80)
                    print(f"[HEURISTICA] {CONT_DATA_INVALIDA} notas consecutivas com data invalida.")
                    print("Encerrando como SEM COMPETENCIA.")
                    print("=" * 80)
                    SEM_COMPETENCIA_NA_EMPRESA = True
                    PARAR_PROCESSAMENTO = True
                    return True
                
                msg = f"Data de emissao invalida/ausente: {info.get('data_emissao', '')}"
                try:
                    salvar_log("", info, status="SKIP_DATA_INVALIDA", mensagem=msg[:180])
                except Exception:
                    pass
                return True

            CONT_DATA_INVALIDA = 0

            if comp == 1:
                CONT_FORA_ANTES_ALVO = 0
                msg = f"Nota mais nova que mes alvo {mes_alvo:02d}/{ano_alvo}"
                try:
                    salvar_log("", info, status="SKIP_FORA_COMPETENCIA", mensagem=msg[:180])
                except Exception:
                    pass
                return True

            if comp == -1:
                msg = f"Nota mais antiga que mes alvo {mes_alvo:02d}/{ano_alvo}"
                try:
                    salvar_log("", info, status="SKIP_FORA_COMPETENCIA", mensagem=msg[:180])
                except Exception:
                    pass

                if ENCONTROU_MES_ALVO:
                    CONT_FORA_APOS_ALVO += 1
                    print(f"  >> Contador notas antigas APOS alvo: {CONT_FORA_APOS_ALVO}/{LIMITE_HEURISTICA_FORA_ALVO}")
                    
                    if CONT_FORA_APOS_ALVO >= LIMITE_HEURISTICA_FORA_ALVO:
                        print("=" * 80)
                        print(f"[HEURISTICA] {CONT_FORA_APOS_ALVO} nota(s) antiga(s) APOS encontrar mes alvo.")
                        print("Encerrando processamento.")
                        print("=" * 80)
                        PARAR_PROCESSAMENTO = True
                else:
                    CONT_FORA_ANTES_ALVO += 1
                    print(f"  >> Contador notas antigas ANTES do alvo: {CONT_FORA_ANTES_ALVO}/{LIMITE_HEURISTICA_FORA_ALVO}")
                    
                    if CONT_FORA_ANTES_ALVO >= LIMITE_HEURISTICA_FORA_ALVO:
                        print("=" * 80)
                        print(f"[HEURISTICA] {CONT_FORA_ANTES_ALVO} nota(s) antiga(s) ANTES de encontrar mes alvo.")
                        print("Empresa NAO possui notas na competencia alvo.")
                        print("Encerrando como SEM COMPETENCIA.")
                        print("=" * 80)
                        SEM_COMPETENCIA_NA_EMPRESA = True
                        PARAR_PROCESSAMENTO = True
                
                return True

            ENCONTROU_MES_ALVO = True
            CONT_FORA_APOS_ALVO = 0
            CONT_FORA_ANTES_ALVO = 0

            print(f"[{i+1}] [BAIXANDO] NF={info.get('nf')} RPS={info.get('rps')} DATA={info.get('data_emissao')}")

            if not checkbox.is_selected():
                click_robusto(checkbox)

            xmls_antes = [a for a in os.listdir(TEMP_DOWNLOAD_DIR) if a.lower().endswith(".xml")]

            botao_xml = wait.until(EC.element_to_be_clickable((By.ID, "_imagebutton12")))
            click_robusto(botao_xml)

            if fechar_aviso_se_existir(timeout=2):
                desmarcar_todas_notas()
                checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
                checkbox = checkboxes[i]
                if not checkbox.is_selected():
                    click_robusto(checkbox)
                botao_xml = wait.until(EC.element_to_be_clickable((By.ID, "_imagebutton12")))
                click_robusto(botao_xml)

            modal = wait.until(EC.visibility_of_element_located((By.ID, "modalboxexportarnotas")))
            select = Select(modal.find_element(By.TAG_NAME, "select"))
            select.select_by_visible_text("NF Nacional")
            sleep(0.15)

            visualizar_btn = modal.find_element(By.XPATH, ".//button[contains(.,'Visualizar')]")
            click_robusto(visualizar_btn)

            xml_baixado = aguardar_xml_novo(TEMP_DOWNLOAD_DIR, timeout=80, xmls_antes=xmls_antes)

            destino_final = organizar_xml_por_pasta(xml_baixado, info)

            logp = salvar_log(destino_final, info, status="OK", mensagem="OK (download confirmado por arquivo)")
            chaves_ok.add(key)

            print(f"  >> Salvo: {destino_final}")

            fechar_modal_exportacao()

            return True

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


def main():
    global PARAR_PROCESSAMENTO, ENCONTROU_MES_ALVO, CONT_FORA_APOS_ALVO, CONT_FORA_ANTES_ALVO, CONT_DATA_INVALIDA, SEM_COMPETENCIA_NA_EMPRESA

    preencher_login_prefeitura_se_habilitado()

    ano_alvo, mes_alvo = calcular_mes_alvo(APURACAO_REFERENCIA)

    try:
        status_lista = esperar_lista_ou_sem_checkbox(timeout=20)
        if status_lista in {"sem_checkbox", "sem_checkbox_com_data"}:
            data_ref = primeira_data_emissao_visivel_sem_checkbox()
            comp = comparar_competencia_nota({"data_emissao": data_ref}, ano_alvo, mes_alvo) if data_ref else None

            if comp == -1:
                print(
                    f"Sem checkbox e data inicial antiga ({data_ref}) para alvo {mes_alvo:02d}/{ano_alvo}. "
                    "Encerrando como sem competencia."
                )
            else:
                print("Lista carregada sem checkbox; encerrando como sem competencia.")

            raise RuntimeError(MSG_SEM_COMPETENCIA)
    except TimeoutException:
        if STRICT_LISTA_INICIAL:
            raise RuntimeError(MSG_CAPTCHA_TIMEOUT)
        print("Aviso: lista inicial nao carregou em 20s; seguindo com tentativas por item.")

    ENCONTROU_MES_ALVO = False
    CONT_FORA_APOS_ALVO = 0
    CONT_FORA_ANTES_ALVO = 0
    CONT_DATA_INVALIDA = 0
    SEM_COMPETENCIA_NA_EMPRESA = False
    PARAR_PROCESSAMENTO = False
    
    print(f"Mes alvo de download: {mes_alvo:02d}/{ano_alvo}")
    print(f"Heuristica de parada: {LIMITE_HEURISTICA_FORA_ALVO} nota(s) antiga(s) consecutiva(s).")

    definir_page_size(100)

    if not verificar_primeira_pagina_tem_mes_alvo(ano_alvo, mes_alvo):
        SEM_COMPETENCIA_NA_EMPRESA = True
        raise RuntimeError(MSG_SEM_COMPETENCIA)

    pagina = 1
    while True:
        total = len(driver.find_elements(By.NAME, "gridListaCheck"))
        print(f"\nPagina {pagina}: {total} notas encontradas.")

        if total == 0 and not ENCONTROU_MES_ALVO:
            SEM_COMPETENCIA_NA_EMPRESA = True
            print("Lista carregada sem checkboxes; encerrando empresa como sem competencia.")
            break

        i = 0
        while True:
            checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
            total = len(checkboxes)
            if i >= total:
                break

            processar_nota_por_indice(i, ano_alvo, mes_alvo)
            
            if PARAR_PROCESSAMENTO:
                print("\n[PARADA] Saindo do loop.")
                break
            
            i += 1

        if PARAR_PROCESSAMENTO:
            print("Encerrando varredura por heuristica de competencia.")
            break

        if not ir_para_proxima_pagina():
            print("Nao ha mais paginas. Finalizando.")
            break
        pagina += 1

    if SEM_COMPETENCIA_NA_EMPRESA:
        raise RuntimeError(MSG_SEM_COMPETENCIA)

    print("Processo finalizado.")
    sleep(2)


try:
    main()
except RuntimeError as e:
    msg = str(e)
    print(msg)
    if MSG_CAPTCHA_TIMEOUT in msg:
        sys.exit(EXIT_CODE_CAPTCHA_TIMEOUT)
    if MSG_SEM_COMPETENCIA in msg:
        sys.exit(EXIT_CODE_SEM_COMPETENCIA)
    raise
finally:
    driver.quit()