import csv
import json
import os
import re
import subprocess
import traceback
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

from core.app_info import APP_MAIN_BUNDLE_NAME, APP_MAIN_EXE_NAME
from core.company_paths import nome_pasta_empresa_por_dados


def get_runtime_base() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_main_command() -> list[str]:
    base = get_runtime_base()

    if getattr(sys, "frozen", False):
        main_exe = base / APP_MAIN_BUNDLE_NAME / APP_MAIN_EXE_NAME
        if not main_exe.exists():
            raise FileNotFoundError(f"Executavel do backend principal nao encontrado: {main_exe}")
        return [str(main_exe)]

    main_py = base / "main.py"
    if not main_py.exists():
        raise FileNotFoundError(f"Arquivo main.py nao encontrado: {main_py}")

    return [sys.executable, str(main_py)]


def executar_main_processo(env_extra: dict | None = None, timeout: int | None = None) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    if env_extra:
        env.update(env_extra)

    cmd = get_main_command()
    return subprocess.run(
        cmd,
        env=env,
        cwd=str(get_runtime_base()),
        timeout=timeout,
    )


BASE_DIR = str(get_runtime_base())
OUTPUT_BASE_DIR = os.environ.get("OUTPUT_BASE_DIR", BASE_DIR)
os.makedirs(OUTPUT_BASE_DIR, exist_ok=True)

CSV_EMPRESAS = os.environ.get("EMPRESAS_ARQUIVO", os.environ.get("EMPRESAS_CSV", "empresas.xlsx"))
REPORT_PATH = os.path.join(OUTPUT_BASE_DIR, "report_execucao_empresas.csv")
CHECKPOINT_PATH = os.path.join(OUTPUT_BASE_DIR, "checkpoint_execucao_empresas.json")
MAX_TENTATIVAS = int(os.environ.get("MAX_TENTATIVAS_EMPRESA", "3"))
LOGIN_WAIT_SECONDS = int(os.environ.get("LOGIN_WAIT_SECONDS", "120"))
TIMEOUT_PROCESSO_MAIN = int(os.environ.get("TIMEOUT_PROCESSO_MAIN", "1800"))
CONTINUAR_DE_ONDE_PAROU = os.environ.get("CONTINUAR_DE_ONDE_PAROU", "1").strip() == "1"
USAR_CHECKPOINT = os.environ.get("USAR_CHECKPOINT", "1").strip() == "1"
MSG_CAPTCHA_TIMEOUT = "CAPTCHA_NAO_RESOLVIDO_NO_TEMPO"
EXIT_CODE_CAPTCHA_TIMEOUT = 30
EXIT_CODE_SEM_COMPETENCIA = 40
EXIT_CODE_SEM_SERVICOS = 41
EXIT_CODE_CREDENCIAL_INVALIDA = 50
EXIT_CODE_TOMADOS_PDF_REVISAO_MANUAL = 51
EXIT_CODE_EMPRESA_MULTIPLA = 52
EXIT_CODE_APURACAO_COMPLETA_REVISAO_MANUAL = 53
EXIT_CODE_MULTI_CADASTRO_REVISAO_MANUAL = 54
EXIT_CODE_TOMADOS_FALHA = 42
EXIT_CODE_CHROME_INIT_FALHA = 60
ARQUIVO_STATUS_TOMADOS_PDF = "_tomados_pdf_status.json"
STATUS_SUCESSO = {"SUCESSO", "SUCESSO_SEM_COMPETENCIA", "SUCESSO_SEM_SERVICOS"}
STATUS_FINAIS_NAO_REPROCESSAR = STATUS_SUCESSO | {"REVISAO_MANUAL"}


def inferir_acao_recomendada(status: str, motivo: str = "") -> str:
    status_limpo = (status or "").strip().upper()
    motivo_limpo = (motivo or "").strip().lower()

    if status_limpo == "REVISAO_MANUAL":
        if "tomados" in motivo_limpo and "pdf" in motivo_limpo:
            return "ABRIR_DEBUG"
        return "CONFERIR_SENHA"
    if status_limpo != "FALHA":
        return ""
    if "captcha" in motivo_limpo:
        return "REPROCESSAR"
    return "ABRIR_DEBUG"


def caminho_status_tomados_pdf_empresa(empresa: dict) -> str:
    pasta_empresa = Path(OUTPUT_BASE_DIR) / "downloads" / nome_pasta_empresa_por_dados(empresa)
    legado = pasta_empresa / ARQUIVO_STATUS_TOMADOS_PDF
    if legado.exists():
        return str(legado)

    competencias = []
    if pasta_empresa.exists():
        for child in pasta_empresa.iterdir():
            if child.is_dir() and re.fullmatch(r"\d{2}\.\d{4}", child.name):
                competencias.append(child)

    competencias.sort(key=lambda path: (int(path.name.split(".")[1]), int(path.name.split(".")[0])), reverse=True)
    for competencia_dir in competencias:
        geral_dir = competencia_dir / "_GERAL"
        candidato = geral_dir / ARQUIVO_STATUS_TOMADOS_PDF
        if candidato.exists():
            return str(candidato)
        prefixed = sorted(
            geral_dir.glob(f"*{ARQUIVO_STATUS_TOMADOS_PDF}"),
            key=lambda path: path.stat().st_mtime,
            reverse=True,
        )
        if prefixed:
            return str(prefixed[0])

    return str(legado)


def carregar_status_tomados_pdf_empresa(empresa: dict) -> dict:
    path = caminho_status_tomados_pdf_empresa(empresa)
    if not os.path.exists(path):
        return {}

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def normalizar_header(h: str) -> str:
    txt = (h or "").strip().lower()
    txt = unicodedata.normalize("NFKD", txt)
    return "".join(ch for ch in txt if not unicodedata.combining(ch))


def mapear_colunas(headers_normalizados):
    col_codigo = None
    col_razao = None
    col_cnpj = None
    col_segmento = None
    col_senha = None

    for h in headers_normalizados:
        if h.startswith("cod"):
            col_codigo = h
        elif "razao" in h:
            col_razao = h
        elif h == "cnpj":
            col_cnpj = h
        elif "segmento" in h:
            col_segmento = h
        elif "senha" in h:
            col_senha = h

    if not all([col_codigo, col_razao, col_cnpj, col_segmento, col_senha]):
        raise ValueError(
            "Colunas obrigatÃ³rias nÃ£o encontradas. Esperado: CÃ³digo, RazÃ£o Social, CNPJ, Segmento, Senha Prefeitura"
        )

    return col_codigo, col_razao, col_cnpj, col_segmento, col_senha


def carregar_empresas_csv(path_csv: str):
    empresas = []
    with open(path_csv, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=';')
        headers_normalizados = [normalizar_header(h) for h in (reader.fieldnames or [])]
        col_codigo, col_razao, col_cnpj, col_segmento, col_senha = mapear_colunas(headers_normalizados)

        for linha_planilha, raw in enumerate(reader, start=2):
            row = {normalizar_header(k): (v or "").strip() for k, v in raw.items()}
            if not row.get(col_cnpj):
                continue
            empresas.append({
                "codigo": row.get(col_codigo, ""),
                "razao_social": row.get(col_razao, ""),
                "cnpj": row.get(col_cnpj, ""),
                "segmento": row.get(col_segmento, ""),
                "senha_prefeitura": row.get(col_senha, ""),
                "indice_lista": len(empresas) + 1,
                "linha_planilha": linha_planilha,
            })
    return empresas


def encontrar_header_xlsx(ws, max_linhas_busca=25):
    linhas = ws.iter_rows(values_only=True)

    for i, row in enumerate(linhas, start=1):
        if i > max_linhas_busca:
            break

        header_norm = [normalizar_header(str(h) if h is not None else "") for h in row]
        try:
            cols = mapear_colunas(header_norm)
            idx = {h: j for j, h in enumerate(header_norm)}
            return i, cols, idx
        except ValueError:
            continue

    raise ValueError(
        "Colunas obrigatÃ³rias nÃ£o encontradas nas primeiras linhas do XLSX. "
        "Esperado: CÃ³digo, RazÃ£o Social, CNPJ, Segmento, Senha Prefeitura"
    )


def carregar_empresas_xlsx(path_xlsx: str):
    try:
        from openpyxl import load_workbook
    except Exception as e:
        raise RuntimeError("Para usar .xlsx instale a dependÃªncia: pip install openpyxl") from e

    wb = load_workbook(path_xlsx, read_only=True, data_only=True)
    ws = wb.active

    linha_header, (col_codigo, col_razao, col_cnpj, col_segmento, col_senha), idx = encontrar_header_xlsx(ws)

    empresas = []
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if i <= linha_header:
            continue

        vals = ["" if v is None else str(v).strip() for v in row]
        cnpj = vals[idx[col_cnpj]] if idx[col_cnpj] < len(vals) else ""
        if not cnpj:
            continue

        empresas.append({
            "codigo": vals[idx[col_codigo]] if idx[col_codigo] < len(vals) else "",
            "razao_social": vals[idx[col_razao]] if idx[col_razao] < len(vals) else "",
            "cnpj": cnpj,
            "segmento": vals[idx[col_segmento]] if idx[col_segmento] < len(vals) else "",
            "senha_prefeitura": vals[idx[col_senha]] if idx[col_senha] < len(vals) else "",
            "indice_lista": len(empresas) + 1,
            "linha_planilha": i,
        })

    return empresas


def carregar_empresas(path_arquivo: str):
    ext = os.path.splitext(path_arquivo)[1].lower()
    if ext == ".xlsx":
        return carregar_empresas_xlsx(path_arquivo)
    return carregar_empresas_csv(path_arquivo)


def filtrar_empresas_para_execucao(empresas, inicio=None, fim=None):
    if inicio is None or fim is None:
        return empresas
    return [
        empresa
        for empresa in empresas
        if inicio <= int(empresa.get("indice_lista", 0)) <= fim
    ]


def resolver_faixa_execucao(empresas):
    raw_inicio = os.environ.get("EMPRESA_INICIO", "").strip()
    raw_fim = os.environ.get("EMPRESA_FIM", "").strip()

    if not raw_inicio and not raw_fim:
        return None, None, empresas

    if not raw_inicio or not raw_fim:
        raise ValueError(
            "Faixa de execucao invalida. Informe EMPRESA_INICIO e EMPRESA_FIM, ou deixe ambos vazios."
        )

    try:
        inicio = int(raw_inicio)
        fim = int(raw_fim)
    except Exception as e:
        raise ValueError(
            "Faixa de execucao invalida. EMPRESA_INICIO e EMPRESA_FIM devem ser inteiros positivos."
        ) from e

    if inicio <= 0 or fim <= 0:
        raise ValueError(
            "Faixa de execucao invalida. EMPRESA_INICIO e EMPRESA_FIM devem ser inteiros positivos."
        )
    if inicio > fim:
        raise ValueError("Faixa de execucao invalida. EMPRESA_INICIO nao pode ser maior que EMPRESA_FIM.")

    total = len(empresas)
    if inicio > total:
        raise ValueError(
            f"Faixa de execucao invalida. EMPRESA_INICIO={inicio} excede o total carregado ({total})."
        )

    fim = min(fim, total)
    empresas_filtradas = filtrar_empresas_para_execucao(empresas, inicio, fim)
    return inicio, fim, empresas_filtradas


def resolver_paths_execucao(output_base_dir, inicio=None, fim=None):
    report_path = os.path.join(output_base_dir, "report_execucao_empresas.csv")
    checkpoint_path = os.path.join(output_base_dir, "checkpoint_execucao_empresas.json")

    if inicio is None or fim is None:
        return report_path, checkpoint_path

    largura = max(3, len(str(max(int(inicio), int(fim)))))
    sufixo = f"__lote_{int(inicio):0{largura}d}_{int(fim):0{largura}d}"
    report_path = os.path.join(output_base_dir, f"report_execucao_empresas{sufixo}.csv")
    checkpoint_path = os.path.join(output_base_dir, f"checkpoint_execucao_empresas{sufixo}.json")
    return report_path, checkpoint_path


def carregar_report_existente(path_report: str):
    concluidas = set()
    if not os.path.exists(path_report):
        return concluidas

    try:
        with open(path_report, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                status = (row.get("status") or "").strip().upper()
                cnpj = re.sub(r"\D", "", row.get("cnpj") or "")
                if status in STATUS_FINAIS_NAO_REPROCESSAR and cnpj:
                    concluidas.add(cnpj)
    except Exception:
        pass

    return concluidas


def carregar_rows_report_existente(path_report: str):
    rows = []
    if not os.path.exists(path_report):
        return rows

    try:
        with open(path_report, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                ts_inicio_raw = (row.get("timestamp_inicio") or "").strip()
                ts_fim_raw = (row.get("timestamp_fim") or "").strip()
                try:
                    ts_inicio = datetime.strptime(ts_inicio_raw, "%Y-%m-%d %H:%M:%S")
                except Exception:
                    ts_inicio = datetime.now()
                try:
                    ts_fim = datetime.strptime(ts_fim_raw, "%Y-%m-%d %H:%M:%S")
                except Exception:
                    ts_fim = ts_inicio

                tentativas_raw = (row.get("tentativas") or "0").strip()
                try:
                    tentativas = int(tentativas_raw or 0)
                except Exception:
                    tentativas = 0

                empresa = {
                    "codigo": (row.get("codigo_empresa") or "").strip(),
                    "razao_social": (row.get("razao_social") or "").strip(),
                    "cnpj": (row.get("cnpj") or "").strip(),
                    "segmento": (row.get("segmento") or "").strip(),
                }
                resultado = {
                    "status": (row.get("status") or "").strip(),
                    "motivo": (row.get("motivo") or "").strip(),
                    "tentativas": tentativas,
                    "acao_recomendada": (row.get("acao_recomendada") or "").strip(),
                }
                if not resultado["acao_recomendada"]:
                    resultado["acao_recomendada"] = inferir_acao_recomendada(
                        resultado["status"],
                        resultado["motivo"],
                    )
                rows.append({"empresa": empresa, "resultado": resultado, "inicio": ts_inicio, "fim": ts_fim})
    except Exception:
        return rows

    return rows


def carregar_checkpoint(path_checkpoint: str):
    if not os.path.exists(path_checkpoint):
        return {"processadas": set(), "ultimo_indice": -1}

    try:
        with open(path_checkpoint, "r", encoding="utf-8") as f:
            data = json.load(f)
        processadas = set(re.sub(r"\D", "", c) for c in data.get("processadas", []) if c)
        ultimo_indice = int(data.get("ultimo_indice", -1))
        return {"processadas": processadas, "ultimo_indice": ultimo_indice}
    except Exception:
        return {"processadas": set(), "ultimo_indice": -1}


def salvar_checkpoint(path_checkpoint: str, processadas, ultimo_indice: int):
    payload = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ultimo_indice": int(ultimo_indice),
        "processadas": sorted(set(processadas)),
    }
    tmp_path = f"{path_checkpoint}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path_checkpoint)


def executar_empresa(empresa: dict):
    inicio = datetime.now()
    ultimo_motivo = ""

    for tentativa in range(1, MAX_TENTATIVAS + 1):
        print("\n" + "=" * 80)
        print(f"Empresa {empresa['codigo']} - {empresa['razao_social']}")
        print(f"CNPJ: {empresa['cnpj']} | Segmento: {empresa['segmento']}")
        print("Senha prefeitura (referencia): [OCULTA]")
        print(f"Tentativa {tentativa}/{MAX_TENTATIVAS}")
        print(
            f"Etapa 1: login automatico (CNPJ/senha) + captcha humano; apos Entrar o robo navega sozinho para Lista Nota Fiscais. Tempo: {LOGIN_WAIT_SECONDS}s."
        )

        env = os.environ.copy()
        env["LOGIN_WAIT_SECONDS"] = str(LOGIN_WAIT_SECONDS)
        env["STRICT_LISTA_INICIAL"] = "1"
        env["EMPRESA_PASTA_FORCADA"] = nome_pasta_empresa_por_dados(empresa)
        env["AUTO_LOGIN_PREFEITURA"] = "1"
        env["EMPRESA_CNPJ"] = re.sub(r"\D", "", empresa["cnpj"])
        env["EMPRESA_SENHA"] = empresa["senha_prefeitura"]
        env["FORCAR_SESSAO_LIMPA_LOGIN"] = os.environ.get("FORCAR_SESSAO_LIMPA_LOGIN", "1")

        try:
            proc = executar_main_processo(env_extra=env, timeout=TIMEOUT_PROCESSO_MAIN)
        except FileNotFoundError as e:
            ultimo_motivo = str(e)
            print(f"Falha tentativa {tentativa}: {ultimo_motivo}")
            break
        except subprocess.TimeoutExpired:
            ultimo_motivo = f"Timeout de execucao do backend principal (> {TIMEOUT_PROCESSO_MAIN}s)"
            print(f"Falha tentativa {tentativa}: {ultimo_motivo}")
            continue

        rc = int(proc.returncode)
        sucesso_por_codigo = {
            0: ("SUCESSO", "OK"),
            EXIT_CODE_SEM_COMPETENCIA: ("SUCESSO_SEM_COMPETENCIA", "Sem notas na competencia alvo"),
            EXIT_CODE_SEM_SERVICOS: ("SUCESSO_SEM_SERVICOS", "Contribuinte sem modulo de Nota Fiscal"),
            EXIT_CODE_CREDENCIAL_INVALIDA: ("REVISAO_MANUAL", "Credencial invalida no portal"),
        }

        if rc == EXIT_CODE_TOMADOS_PDF_REVISAO_MANUAL:
            status_pdf = carregar_status_tomados_pdf_empresa(empresa)
            motivo = (status_pdf.get("codigo_erro") or "").strip() or "PDF_TOMADOS_NAO_GERADO"
            return {
                "status": "REVISAO_MANUAL",
                "motivo": motivo,
                "tentativas": tentativa,
                "acao_recomendada": inferir_acao_recomendada("REVISAO_MANUAL", motivo),
                "inicio": inicio,
                "fim": datetime.now(),
            }

        if rc == EXIT_CODE_EMPRESA_MULTIPLA:
            motivo = "EMPRESA_MULTIPLA"
            return {
                "status": "REVISAO_MANUAL",
                "motivo": motivo,
                "tentativas": tentativa,
                "acao_recomendada": inferir_acao_recomendada("REVISAO_MANUAL", motivo),
                "inicio": inicio,
                "fim": datetime.now(),
            }

        if rc == EXIT_CODE_APURACAO_COMPLETA_REVISAO_MANUAL:
            motivo = "APURACAO_COMPLETA_COM_FALHAS"
            return {
                "status": "REVISAO_MANUAL",
                "motivo": motivo,
                "tentativas": tentativa,
                "acao_recomendada": inferir_acao_recomendada("REVISAO_MANUAL", motivo),
                "inicio": inicio,
                "fim": datetime.now(),
            }

        if rc == EXIT_CODE_MULTI_CADASTRO_REVISAO_MANUAL:
            motivo = "MULTIPLOS_CADASTROS_COM_FALHAS"
            return {
                "status": "REVISAO_MANUAL",
                "motivo": motivo,
                "tentativas": tentativa,
                "acao_recomendada": inferir_acao_recomendada("REVISAO_MANUAL", motivo),
                "inicio": inicio,
                "fim": datetime.now(),
            }

        if rc in sucesso_por_codigo:
            status, motivo = sucesso_por_codigo[rc]
            return {
                "status": status,
                "motivo": motivo,
                "tentativas": tentativa,
                "acao_recomendada": inferir_acao_recomendada(status, motivo),
                "inicio": inicio,
                "fim": datetime.now(),
            }

        if rc == EXIT_CODE_CAPTCHA_TIMEOUT:
            ultimo_motivo = "Captcha nao resolvido a tempo"
        elif rc == EXIT_CODE_TOMADOS_FALHA:
            ultimo_motivo = "Falha na etapa de Servicos Tomados"
            print(f"Falha tentativa {tentativa}: {ultimo_motivo}")
            break
        elif rc == EXIT_CODE_CHROME_INIT_FALHA:
            ultimo_motivo = "Falha na inicializacao do Chrome/WebDriver"
            print(f"Falha tentativa {tentativa}: {ultimo_motivo}")
            break
        else:
            ultimo_motivo = f"Falha execucao (exit={rc})"

        print(f"Falha tentativa {tentativa}: {ultimo_motivo}")

    return {
        "status": "FALHA",
        "motivo": ultimo_motivo or "Falha desconhecida",
        "tentativas": MAX_TENTATIVAS,
        "acao_recomendada": inferir_acao_recomendada("FALHA", ultimo_motivo or "Falha desconhecida"),
        "inicio": inicio,
        "fim": datetime.now(),
    }

REPORT_HEADER = [
    "timestamp_inicio",
    "timestamp_fim",
    "codigo_empresa",
    "razao_social",
    "cnpj",
    "segmento",
    "status",
    "motivo",
    "tentativas",
    "acao_recomendada",
]


def garantir_report_com_header(path_report: str):
    if not os.path.exists(path_report):
        with open(path_report, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f, delimiter=';').writerow(REPORT_HEADER)
        return

    with open(path_report, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=';')
        fieldnames = reader.fieldnames or []
        if fieldnames == REPORT_HEADER:
            return
        rows = list(reader)

    with open(path_report, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(REPORT_HEADER)
        for row in rows:
            status = (row.get("status") or "").strip()
            motivo = (row.get("motivo") or "").strip()
            writer.writerow([
                (row.get("timestamp_inicio") or "").strip(),
                (row.get("timestamp_fim") or "").strip(),
                (row.get("codigo_empresa") or "").strip(),
                (row.get("razao_social") or "").strip(),
                (row.get("cnpj") or "").strip(),
                (row.get("segmento") or "").strip(),
                status,
                motivo,
                (row.get("tentativas") or "").strip(),
                (row.get("acao_recomendada") or "").strip() or inferir_acao_recomendada(status, motivo),
            ])


def resetar_report(path_report: str):
    with open(path_report, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f, delimiter=';').writerow(REPORT_HEADER)


def append_report_row(path_report: str, row):
    garantir_report_com_header(path_report)
    with open(path_report, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=';')
        w.writerow([
            row["inicio"].strftime("%Y-%m-%d %H:%M:%S"),
            row["fim"].strftime("%Y-%m-%d %H:%M:%S"),
            row["empresa"]["codigo"],
            row["empresa"]["razao_social"],
            row["empresa"]["cnpj"],
            row["empresa"]["segmento"],
            row["resultado"]["status"],
            row["resultado"]["motivo"],
            row["resultado"]["tentativas"],
            row["resultado"].get("acao_recomendada", ""),
        ])


def main():
    if not os.path.exists(CSV_EMPRESAS):
        print(f"Arquivo de empresas nÃ£o encontrado: {CSV_EMPRESAS}")
        print("Informe EMPRESAS_ARQUIVO com caminho vÃ¡lido (.xlsx/.csv).")
        return

    empresas = carregar_empresas(CSV_EMPRESAS)
    print(f"Empresas carregadas: {len(empresas)}")

    try:
        faixa_inicio, faixa_fim, empresas = resolver_faixa_execucao(empresas)
    except ValueError as e:
        raise SystemExit(str(e)) from e
    report_path, checkpoint_path = resolver_paths_execucao(OUTPUT_BASE_DIR, faixa_inicio, faixa_fim)

    faixa_txt = "todas" if faixa_inicio is None else f"{faixa_inicio}-{faixa_fim}"
    print(f"Faixa aplicada: {faixa_txt}")
    print(f"Empresas selecionadas: {len(empresas)}")
    print(f"Report: {report_path}")
    print(f"Checkpoint: {checkpoint_path}")

    concluidas = set()
    if CONTINUAR_DE_ONDE_PAROU:
        concluidas = carregar_report_existente(report_path)
        if concluidas:
            print(f"Retomada ativa: {len(concluidas)} empresas jÃ¡ concluÃ­das serÃ£o puladas.")

    start_idx = 0
    processadas_checkpoint = set()
    ck_ultimo_indice = -1
    if USAR_CHECKPOINT:
        ck = carregar_checkpoint(checkpoint_path)
        processadas_checkpoint = ck["processadas"]
        ck_ultimo_indice = int(ck.get("ultimo_indice", -1))
        if processadas_checkpoint:
            print(f"Checkpoint ativo: {len(processadas_checkpoint)} empresas já processadas serão puladas.")
        if CONTINUAR_DE_ONDE_PAROU and ck_ultimo_indice >= 0:
            start_idx = ck_ultimo_indice + 1
            print(f"Retomando a partir do índice {start_idx} (checkpoint).")

    ja_processadas = set(concluidas) | set(processadas_checkpoint)

    resultados = carregar_rows_report_existente(report_path) if CONTINUAR_DE_ONDE_PAROU else []
    if CONTINUAR_DE_ONDE_PAROU:
        garantir_report_com_header(report_path)
    else:
        resetar_report(report_path)
    cnpjs_ja_no_report = {re.sub(r"\D", "", r["empresa"].get("cnpj", "")) for r in resultados if r.get("empresa")}

    last_idx_concluido = start_idx - 1
    try:
        for idx, empresa in enumerate(empresas[start_idx:], start=start_idx):
            cnpj_limpo = re.sub(r"\D", "", empresa.get("cnpj", ""))

            if cnpj_limpo and cnpj_limpo in ja_processadas:
                # Atualiza checkpoint mesmo quando pula, para retomar do ponto exato
                last_idx_concluido = idx
                if USAR_CHECKPOINT:
                    salvar_checkpoint(checkpoint_path, processadas_checkpoint, idx)

                if cnpj_limpo in cnpjs_ja_no_report:
                    continue

                agora = datetime.now()
                resultado = {
                    "status": "SUCESSO",
                    "motivo": "Pulada por retomada/checkpoint (já processada)",
                    "tentativas": 0,
                    "acao_recomendada": "",
                    "inicio": agora,
                    "fim": agora,
                }
                row = {"empresa": empresa, "resultado": resultado, "inicio": agora, "fim": agora}
                resultados.append(row)
                cnpjs_ja_no_report.add(cnpj_limpo)
                append_report_row(report_path, row)
                continue

            try:
                res = executar_empresa(empresa)
            except Exception as e:
                agora = datetime.now()
                res = {
                    "status": "FALHA",
                    "motivo": f"Erro inesperado no orquestrador: {str(e)[:160]}",
                    "tentativas": 0,
                    "acao_recomendada": inferir_acao_recomendada(
                        "FALHA",
                        f"Erro inesperado no orquestrador: {str(e)[:160]}",
                    ),
                    "inicio": agora,
                    "fim": agora,
                }

            if not res.get("acao_recomendada"):
                res["acao_recomendada"] = inferir_acao_recomendada(
                    res.get("status", ""),
                    res.get("motivo", ""),
                )

            row = {"empresa": empresa, "resultado": res, "inicio": res["inicio"], "fim": res["fim"]}
            resultados.append(row)
            if cnpj_limpo:
                cnpjs_ja_no_report.add(cnpj_limpo)
            append_report_row(report_path, row)

            if cnpj_limpo and res.get("status") in STATUS_FINAIS_NAO_REPROCESSAR:
                ja_processadas.add(cnpj_limpo)
                if USAR_CHECKPOINT:
                    processadas_checkpoint.add(cnpj_limpo)

            last_idx_concluido = idx
            if USAR_CHECKPOINT:
                # Atualiza SEMPRE o índice (sucesso ou falha) para retomar do ponto exato
                salvar_checkpoint(checkpoint_path, processadas_checkpoint, idx)

    except KeyboardInterrupt:
        print("\nInterrompido pelo usuário (Ctrl+C). Salvando checkpoint e encerrando...")
        if USAR_CHECKPOINT:
            salvar_checkpoint(checkpoint_path, processadas_checkpoint, last_idx_concluido)
        return
    except Exception:
        print("\nErro inesperado no loop principal. Salvando checkpoint antes de sair...")
        print(traceback.format_exc())
        if USAR_CHECKPOINT:
            salvar_checkpoint(checkpoint_path, processadas_checkpoint, last_idx_concluido)
        raise

    total = len(resultados)
    ok = sum(1 for r in resultados if r["resultado"]["status"] in STATUS_SUCESSO)
    revisao_manual = sum(1 for r in resultados if r["resultado"]["status"] == "REVISAO_MANUAL")
    falha = sum(1 for r in resultados if r["resultado"]["status"] not in STATUS_FINAIS_NAO_REPROCESSAR)
    print("\n" + "=" * 80)
    print(f"Processamento finalizado. Total={total} | Sucesso={ok} | RevisaoManual={revisao_manual} | Falha={falha}")
    print(f"Report: {report_path}")

    if USAR_CHECKPOINT and os.path.exists(checkpoint_path):
        if falha == 0:
            os.remove(checkpoint_path)
            print(f"Checkpoint removido ao final: {checkpoint_path}")
        else:
            print(f"Checkpoint mantido para retomada: {checkpoint_path}")


if __name__ == "__main__":
    main()
