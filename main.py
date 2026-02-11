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

# Se quiser travar a empresa SEM depender do texto do site, descomente:
# EMPRESA_PASTA_FORCADA = "H2_IMOBILIARIA"
EMPRESA_PASTA_FORCADA = ""

# Competência alvo: por padrão, mês anterior ao mês atual (apuração).
# Pode sobrescrever com APURACAO_REFERENCIA=MM/AAAA (ex.: 03/2026 -> alvo 02/2026).
APURACAO_REFERENCIA = os.environ.get("APURACAO_REFERENCIA", "").strip()

PARAR_PROCESSAMENTO = False

# =====================
# CHROME – PERFIL EXCLUSIVO
# =====================
options = Options()
options.add_argument("--start-maximized")

# Perfil exclusivo com fallback cross-platform.
DEFAULT_PROFILE_DIR = (
    r"C:\ChromeRobotProfile" if os.name == "nt"
    else os.path.join(tempfile.gettempdir(), "ChromeRobotProfile")
)
CHROME_PROFILE_DIR = os.environ.get("CHROME_PROFILE_DIR", DEFAULT_PROFILE_DIR)
options.add_argument(f"--user-data-dir={CHROME_PROFILE_DIR}")

prefs = {
    "download.default_directory": TEMP_DOWNLOAD_DIR,  # <<< importante
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

print("Chrome iniciado.")
print("Entre manualmente em: Nota Fiscal - Lista Nota Fiscais")
login_wait_seconds = int(os.environ.get("LOGIN_WAIT_SECONDS", "40"))
print(f"Voce tem {login_wait_seconds} segundos.")
sleep(login_wait_seconds)

# =====================
# HELPERS – CLIQUE / LISTA / 502
# =====================
def click_robusto(el):
    try:
        el.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", el)

def esperar_lista(timeout=WAIT_LISTA_TIMEOUT):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located((By.NAME, "gridListaCheck"))
    )

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

        # padrões comuns de página 502
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

# =====================
# EXTRAIR INFO DA LINHA (NF/DATA/RPS/CHAVE)
# =====================
def td_text(tr, col):
    el = tr.find_element(By.CSS_SELECTOR, f"td[id*=',{col}_gridLista']")
    return (el.text or "").strip()

def extrair_info_linha(checkbox):
    tr = checkbox.find_element(By.XPATH, "./ancestor::tr[1]")
    return {
        "id_interno": checkbox.get_attribute("value") or "",
        "nf": td_text(tr, 6),
        "situacao": td_text(tr, 9),
        "data_emissao": td_text(tr, 10),  # dd/mm/aaaa
        "rps": td_text(tr, 11),
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

def nota_no_mes_alvo(info: dict, ano_alvo: int, mes_alvo: int) -> bool:
    parsed = parse_data_emissao_site(info.get("data_emissao", ""))
    if not parsed:
        return False
    ano = int(parsed[0])
    mes = int(parsed[1])
    return ano == ano_alvo and mes == mes_alvo

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
# DOWNLOAD – esperar XML NOVO (na pasta _tmp)
# =====================
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
# EMPRESA – PASTA CANONICA (evita H2_IMOBILIARIA vs IMOBILIARIA_H2_LTDA)
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

# =====================
# ORGANIZAR (competencia MM.AAAA) + mover com retry
# =====================
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
      downloads/EMPRESA_PASTA/MM.AAAA/NFS_<NF>_<YYYY-MM-DD>.xml
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
        novo_nome = f"NFS_{nf}_{data_iso}.xml"
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

# =====================
# PROCESSAR UMA NOTA (sem falso 502)
# =====================
def processar_nota_por_indice(i, ano_alvo, mes_alvo):
    global PARAR_PROCESSAMENTO

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

            # garante só 1 selecionada
            desmarcar_todas_notas()

            checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
            checkbox = checkboxes[i]

            info = extrair_info_linha(checkbox)
            key = chave_unica(info)

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

            # Regra de negocio: ao encontrar nota fora do mês alvo, encerramos o robô.
            if not nota_no_mes_alvo(info, ano_alvo, mes_alvo):
                msg = f"Fora do mês alvo {mes_alvo:02d}/{ano_alvo} (data_emissao: {info.get('data_emissao', '')})"
                print(f"[{i+1}] FORA DA COMPETENCIA -> NF={info.get('nf')} DATA={info.get('data_emissao')} | Encerrando processo.")
                try:
                    salvar_log("", info, status="SKIP_FORA_COMPETENCIA", mensagem=msg[:180])
                except Exception:
                    pass
                PARAR_PROCESSAMENTO = True
                return True

            print(f"[{i+1}] Tentativa {tentativa}/{MAX_RETRIES_POR_NOTA} -> NF={info.get('nf')} RPS={info.get('rps')} DATA={info.get('data_emissao')}")

            if not checkbox.is_selected():
                click_robusto(checkbox)

            xmls_antes = [a for a in os.listdir(TEMP_DOWNLOAD_DIR) if a.lower().endswith(".xml")]

            # botão exportar XML (ID fixo)
            botao_xml = wait.until(EC.element_to_be_clickable((By.ID, "_imagebutton12")))
            click_robusto(botao_xml)

            # aviso de "só 1 por vez"
            if fechar_aviso_se_existir(timeout=2):
                desmarcar_todas_notas()
                checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
                checkbox = checkboxes[i]
                if not checkbox.is_selected():
                    click_robusto(checkbox)
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
    global PARAR_PROCESSAMENTO

    try:
        esperar_lista(timeout=20)
    except TimeoutException:
        print("Aviso: lista inicial nao carregou em 20s; seguindo com tentativas por item.")

    ano_alvo, mes_alvo = calcular_mes_alvo(APURACAO_REFERENCIA)
    print(f"Mes alvo de download: {mes_alvo:02d}/{ano_alvo}")

    definir_page_size(100)

    pagina = 1
    while True:
        total = len(driver.find_elements(By.NAME, "gridListaCheck"))
        print(f"Pagina {pagina}: {total} notas encontradas.")

        i = 0
        while True:
            checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
            total = len(checkboxes)
            if i >= total:
                break

            processar_nota_por_indice(i, ano_alvo, mes_alvo)
            if PARAR_PROCESSAMENTO:
                break
            i += 1

        if PARAR_PROCESSAMENTO:
            print("Encerrando varredura ao encontrar nota fora do mês alvo.")
            break

        if not ir_para_proxima_pagina():
            break
        pagina += 1

    print("Processo finalizado.")
    sleep(2)


try:
    main()
finally:
    driver.quit()
