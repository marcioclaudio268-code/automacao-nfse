import csv
import json
import os
import re
import subprocess
import traceback
import sys
import unicodedata
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_EMPRESAS = os.environ.get("EMPRESAS_ARQUIVO", os.environ.get("EMPRESAS_CSV", "empresas.xlsx"))
REPORT_PATH = os.path.join(BASE_DIR, "report_execucao_empresas.csv")
CHECKPOINT_PATH = os.path.join(BASE_DIR, "checkpoint_execucao_empresas.json")
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
EXIT_CODE_TOMADOS_FALHA = 42
EXIT_CODE_CHROME_INIT_FALHA = 60
STATUS_SUCESSO = {"SUCESSO", "SUCESSO_SEM_COMPETENCIA", "SUCESSO_SEM_SERVICOS"}


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

        for raw in reader:
            row = {normalizar_header(k): (v or "").strip() for k, v in raw.items()}
            if not row.get(col_cnpj):
                continue
            empresas.append({
                "codigo": row.get(col_codigo, ""),
                "razao_social": row.get(col_razao, ""),
                "cnpj": row.get(col_cnpj, ""),
                "segmento": row.get(col_segmento, ""),
                "senha_prefeitura": row.get(col_senha, ""),
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
        })

    return empresas


def carregar_empresas(path_arquivo: str):
    ext = os.path.splitext(path_arquivo)[1].lower()
    if ext == ".xlsx":
        return carregar_empresas_xlsx(path_arquivo)
    return carregar_empresas_csv(path_arquivo)


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
                if status in {"SUCESSO", "SUCESSO_SEM_COMPETENCIA", "SUCESSO_SEM_SERVICOS"} and cnpj:
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
                }
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
    main_script = os.path.join(BASE_DIR, "main.py")

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
        env["EMPRESA_PASTA_FORCADA"] = empresa["razao_social"]
        env["AUTO_LOGIN_PREFEITURA"] = "1"
        env["EMPRESA_CNPJ"] = re.sub(r"\D", "", empresa["cnpj"])
        env["EMPRESA_SENHA"] = empresa["senha_prefeitura"]
        env["FORCAR_SESSAO_LIMPA_LOGIN"] = os.environ.get("FORCAR_SESSAO_LIMPA_LOGIN", "1")

        try:
            proc = subprocess.run(
                [sys.executable, main_script],
                env=env,
                cwd=BASE_DIR,
                timeout=TIMEOUT_PROCESSO_MAIN,
            )
        except subprocess.TimeoutExpired:
            ultimo_motivo = f"Timeout de execucao do main.py (> {TIMEOUT_PROCESSO_MAIN}s)"
            print(f"Falha tentativa {tentativa}: {ultimo_motivo}")
            continue

        rc = int(proc.returncode)
        sucesso_por_codigo = {
            0: ("SUCESSO", "OK"),
            EXIT_CODE_SEM_COMPETENCIA: ("SUCESSO_SEM_COMPETENCIA", "Sem notas na competencia alvo"),
            EXIT_CODE_SEM_SERVICOS: ("SUCESSO_SEM_SERVICOS", "Contribuinte sem modulo de Nota Fiscal"),
            EXIT_CODE_CREDENCIAL_INVALIDA: ("SUCESSO", "Revisar manualmente: credencial inválida (usuário/senha) no portal"),
        }

        if rc in sucesso_por_codigo:
            status, motivo = sucesso_por_codigo[rc]
            return {
                "status": status,
                "motivo": motivo,
                "tentativas": tentativa,
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
]


def garantir_report_com_header(path_report: str):
    if os.path.exists(path_report):
        return
    with open(path_report, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f, delimiter=';').writerow(REPORT_HEADER)


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
        ])


def main():
    if not os.path.exists(CSV_EMPRESAS):
        print(f"Arquivo de empresas nÃ£o encontrado: {CSV_EMPRESAS}")
        print("Informe EMPRESAS_ARQUIVO com caminho vÃ¡lido (.xlsx/.csv).")
        return

    empresas = carregar_empresas(CSV_EMPRESAS)
    print(f"Empresas carregadas: {len(empresas)}")

    concluidas = set()
    if CONTINUAR_DE_ONDE_PAROU:
        concluidas = carregar_report_existente(REPORT_PATH)
        if concluidas:
            print(f"Retomada ativa: {len(concluidas)} empresas jÃ¡ concluÃ­das serÃ£o puladas.")

    start_idx = 0
    processadas_checkpoint = set()
    ck_ultimo_indice = -1
    if USAR_CHECKPOINT:
        ck = carregar_checkpoint(CHECKPOINT_PATH)
        processadas_checkpoint = ck["processadas"]
        ck_ultimo_indice = int(ck.get("ultimo_indice", -1))
        if processadas_checkpoint:
            print(f"Checkpoint ativo: {len(processadas_checkpoint)} empresas já processadas serão puladas.")
        if CONTINUAR_DE_ONDE_PAROU and ck_ultimo_indice >= 0:
            start_idx = ck_ultimo_indice + 1
            print(f"Retomando a partir do índice {start_idx} (checkpoint).")

    ja_processadas = set(concluidas) | set(processadas_checkpoint)

    resultados = carregar_rows_report_existente(REPORT_PATH) if CONTINUAR_DE_ONDE_PAROU else []
    if CONTINUAR_DE_ONDE_PAROU:
        garantir_report_com_header(REPORT_PATH)
    else:
        resetar_report(REPORT_PATH)
    cnpjs_ja_no_report = {re.sub(r"\D", "", r["empresa"].get("cnpj", "")) for r in resultados if r.get("empresa")}

    last_idx_concluido = start_idx - 1
    try:
        for idx, empresa in enumerate(empresas[start_idx:], start=start_idx):
            cnpj_limpo = re.sub(r"\D", "", empresa.get("cnpj", ""))

            if cnpj_limpo and cnpj_limpo in ja_processadas:
                # Atualiza checkpoint mesmo quando pula, para retomar do ponto exato
                last_idx_concluido = idx
                if USAR_CHECKPOINT:
                    salvar_checkpoint(CHECKPOINT_PATH, processadas_checkpoint, idx)

                if cnpj_limpo in cnpjs_ja_no_report:
                    continue

                agora = datetime.now()
                resultado = {
                    "status": "SUCESSO",
                    "motivo": "Pulada por retomada/checkpoint (já processada)",
                    "tentativas": 0,
                    "inicio": agora,
                    "fim": agora,
                }
                row = {"empresa": empresa, "resultado": resultado, "inicio": agora, "fim": agora}
                resultados.append(row)
                cnpjs_ja_no_report.add(cnpj_limpo)
                append_report_row(REPORT_PATH, row)
                continue

            try:
                res = executar_empresa(empresa)
            except Exception as e:
                agora = datetime.now()
                res = {
                    "status": "FALHA",
                    "motivo": f"Erro inesperado no orquestrador: {str(e)[:160]}",
                    "tentativas": 0,
                    "inicio": agora,
                    "fim": agora,
                }

            row = {"empresa": empresa, "resultado": res, "inicio": res["inicio"], "fim": res["fim"]}
            resultados.append(row)
            if cnpj_limpo:
                cnpjs_ja_no_report.add(cnpj_limpo)
            append_report_row(REPORT_PATH, row)

            if cnpj_limpo and res.get("status") in STATUS_SUCESSO:
                ja_processadas.add(cnpj_limpo)
                if USAR_CHECKPOINT:
                    processadas_checkpoint.add(cnpj_limpo)

            last_idx_concluido = idx
            if USAR_CHECKPOINT:
                # Atualiza SEMPRE o índice (sucesso ou falha) para retomar do ponto exato
                salvar_checkpoint(CHECKPOINT_PATH, processadas_checkpoint, idx)

    except KeyboardInterrupt:
        print("\nInterrompido pelo usuário (Ctrl+C). Salvando checkpoint e encerrando...")
        if USAR_CHECKPOINT:
            salvar_checkpoint(CHECKPOINT_PATH, processadas_checkpoint, last_idx_concluido)
        return
    except Exception:
        print("\nErro inesperado no loop principal. Salvando checkpoint antes de sair...")
        print(traceback.format_exc())
        if USAR_CHECKPOINT:
            salvar_checkpoint(CHECKPOINT_PATH, processadas_checkpoint, last_idx_concluido)
        raise

    total = len(resultados)
    ok = sum(1 for r in resultados if r["resultado"]["status"] in STATUS_SUCESSO)
    falha = total - ok
    print("\n" + "=" * 80)
    print(f"Processamento finalizado. Total={total} | Sucesso={ok} | Falha={falha}")
    print(f"Report: {REPORT_PATH}")

    if USAR_CHECKPOINT and os.path.exists(CHECKPOINT_PATH):
        if falha == 0:
            os.remove(CHECKPOINT_PATH)
            print(f"Checkpoint removido ao final: {CHECKPOINT_PATH}")
        else:
            print(f"Checkpoint mantido para retomada: {CHECKPOINT_PATH}")


if __name__ == "__main__":
    main()