from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from time import sleep
import os, time, csv
import xml.etree.ElementTree as ET
import shutil

# =====================
# CONFIGURACAO
# =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# =====================
# CHROME – PERFIL EXCLUSIVO
# =====================
options = Options()
options.add_argument("--start-maximized")
options.add_argument(r"--user-data-dir=C:\ChromeRobotProfile")

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

print("Chrome iniciado.")
print("Entre manualmente em: Nota Fiscal - Lista Nota Fiscais")
print("Voce tem 40 segundos.")
sleep(40)

# =====================
# ESPERAR DOWNLOAD REAL
# =====================
def aguardar_xml_novo(pasta, timeout=40):
    fim = time.time() + timeout
    while time.time() < fim:
        arquivos = os.listdir(pasta)

        # enquanto estiver baixando
        if any(a.endswith(".crdownload") for a in arquivos):
            sleep(0.5)
            continue

        xmls = [a for a in arquivos if a.lower().endswith(".xml")]
        if xmls:
            xmls.sort(
                key=lambda x: os.path.getmtime(os.path.join(pasta, x)),
                reverse=True
            )
            return os.path.join(pasta, xmls[0])

        sleep(0.5)

    raise TimeoutError("Download nao finalizou")

# =====================
# FECHAR MODAL (CORRETO)
# =====================
def fechar_modal():
    try:
        # clicar no X
        botao = driver.find_element(By.XPATH, "//button[contains(@class,'close')]")
        driver.execute_script("arguments[0].click();", botao)

        # ESPERA REAL: overlay desaparecer
        WebDriverWait(driver, 30).until_not(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class,'modal-backdrop') or contains(@class,'ui-widget-overlay')]")
            )
        )

        # também aguarda modal sumir
        WebDriverWait(driver, 30).until_not(
            EC.presence_of_element_located(
                (By.ID, "modalboxexportarnotas")
            )
        )

        sleep(0.5)

    except:
        pass

# =====================
# ORGANIZAR XML
# =====================
def organizar_xml_por_pasta(caminho_xml):
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()

        ns = {"ns": "http://www.sped.fazenda.gov.br/nfse"}

        def get(xpath):
            el = root.find(xpath, ns)
            return el.text.strip() if el is not None and el.text else ""

        empresa = get(".//ns:emit/ns:xNome")
        data_raw = get(".//ns:dhProc")
        numero = get(".//ns:nDFSe")

        if not empresa or not data_raw:
            print("Nao foi possivel identificar empresa/data.")
            return caminho_xml

        data = data_raw[:10]
        ano, mes, _ = data.split("-")

        empresa = empresa.replace("/", "").replace(".", "").replace("-", "").replace(" ", "_").upper()

        destino = os.path.join(DOWNLOAD_DIR, empresa, ano, mes)
        os.makedirs(destino, exist_ok=True)

        novo_nome = f"NFS_{numero}_{data}.xml"
        destino_final = os.path.join(destino, novo_nome)

        shutil.move(caminho_xml, destino_final)
        return destino_final

    except:
        return caminho_xml

# =====================
# PAGINACAO REAL
# =====================
def ir_para_proxima_pagina():
    try:
        pagina_atual = driver.find_element(By.NAME, "gridListaPage").get_attribute("value")

        botao = driver.find_element(By.XPATH, "//span[normalize-space(text())='»']")
        driver.execute_script("arguments[0].click();", botao)

        wait.until(
            lambda d: d.find_element(By.NAME, "gridListaPage").get_attribute("value") != pagina_atual
        )

        sleep(1)
        return True

    except:
        return False

# =====================
# LOOP PRINCIPAL
# =====================
pagina = 1

while True:
    print(f"\nProcessando pagina {pagina}")

    checkboxes = driver.find_elements(By.NAME, "gridListaCheck")
    print(f"Notas encontradas na pagina: {len(checkboxes)}")

    for checkbox in checkboxes:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", checkbox)
            sleep(0.3)

            if not checkbox.is_selected():
                checkbox.click()

            # abrir exportacao xml
            botao_xml = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//span[contains(@class,'fa-file-code-o')]/parent::*")
                )
            )
            botao_xml.click()

            # esperar modal
            modal = wait.until(
                EC.visibility_of_element_located((By.ID, "modalboxexportarnotas"))
            )

            # selecionar NF Nacional
            select = Select(modal.find_element(By.TAG_NAME, "select"))
            select.select_by_visible_text("NF Nacional")
            sleep(0.4)

            # visualizar XML
            modal.find_element(By.XPATH, ".//button[contains(.,'Visualizar')]").click()

            # aguardar download
            xml = aguardar_xml_novo(DOWNLOAD_DIR)
            organizar_xml_por_pasta(xml)

            # FECHAR MODAL CORRETAMENTE
            fechar_modal()

            # desmarcar
            if checkbox.is_selected():
                checkbox.click()

            sleep(0.4)

        except Exception as e:
            print("Erro:", e)
            continue

    if not ir_para_proxima_pagina():
        print("Nao ha mais paginas. Finalizando.")
        break

    pagina += 1

print("Processo finalizado com sucesso.")
sleep(3)
driver.quit()
