import csv
import os
import subprocess
import sys
import unicodedata
import re
from datetime import datetime

CSV_EMPRESAS = os.environ.get("EMPRESAS_ARQUIVO", os.environ.get("EMPRESAS_CSV", "empresas.xlsx"))
REPORT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "report_execucao_empresas.csv")
MAX_TENTATIVAS = int(os.environ.get("MAX_TENTATIVAS_EMPRESA", "3"))
LOGIN_WAIT_SECONDS = int(os.environ.get("LOGIN_WAIT_SECONDS", "120"))
CONTINUAR_DE_ONDE_PAROU = os.environ.get("CONTINUAR_DE_ONDE_PAROU", "1").strip() == "1"
MSG_CAPTCHA_TIMEOUT = "CAPTCHA_NAO_RESOLVIDO_NO_TEMPO"
EXIT_CODE_CAPTCHA_TIMEOUT = 30
EXIT_CODE_SEM_COMPETENCIA = 40
EXIT_CODE_SEM_SERVICOS = 41
EXIT_CODE_CREDENCIAL_INVALIDA = 50


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
            "Colunas obrigatórias não encontradas. Esperado: Código, Razão Social, CNPJ, Segmento, Senha Prefeitura"
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
    cache = []

    for i, row in enumerate(linhas, start=1):
        cache.append(row)
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
        "Colunas obrigatórias não encontradas nas primeiras linhas do XLSX. "
        "Esperado: Código, Razão Social, CNPJ, Segmento, Senha Prefeitura"
    )

def carregar_empresas_xlsx(path_xlsx: str):
    try:
        from openpyxl import load_workbook
    except Exception as e:
        raise RuntimeError("Para usar .xlsx instale a dependência: pip install openpyxl") from e

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
                if status in {"SUCESSO", "SUCESSO_SEM_COMPETENCIA"} and cnpj:
                    concluidas.add(cnpj)
    except Exception:
        pass

    return concluidas


def executar_empresa(empresa: dict):
    inicio = datetime.now()
    ultimo_motivo = ""

    for tentativa in range(1, MAX_TENTATIVAS + 1):
        print("\n" + "=" * 80)
        print(f"Empresa {empresa['codigo']} - {empresa['razao_social']}")
        print(f"CNPJ: {empresa['cnpj']} | Segmento: {empresa['segmento']}")
        print(f"Senha prefeitura (referência): {empresa['senha_prefeitura']}")
        print(f"Tentativa {tentativa}/{MAX_TENTATIVAS}")
        print(
            f"Etapa 1: login automático (CNPJ/senha) + captcha humano; após Entrar o robô navega sozinho para Lista Nota Fiscais. Tempo: {LOGIN_WAIT_SECONDS}s."
        )

        env = os.environ.copy()
        env["LOGIN_WAIT_SECONDS"] = str(LOGIN_WAIT_SECONDS)
        env["STRICT_LISTA_INICIAL"] = "1"
        env["EMPRESA_PASTA_FORCADA"] = empresa["razao_social"]
        env["AUTO_LOGIN_PREFEITURA"] = "1"
        env["EMPRESA_CNPJ"] = re.sub(r"\D", "", empresa["cnpj"])
        env["EMPRESA_SENHA"] = empresa["senha_prefeitura"]

        proc = subprocess.run(
            [sys.executable, "main.py"],
            env=env,
        )

        if proc.returncode == 0:
            return {
                "status": "SUCESSO",
                "motivo": "OK",
                "tentativas": tentativa,
                "inicio": inicio,
                "fim": datetime.now(),
            }

        if proc.returncode == EXIT_CODE_SEM_COMPETENCIA:
            return {
                "status": "SUCESSO_SEM_COMPETENCIA",
                "motivo": "Sem notas na competência alvo",
                "tentativas": tentativa,
                "inicio": inicio,
                "fim": datetime.now(),
            }

        if proc.returncode == EXIT_CODE_SEM_SERVICOS:
            return {
                "status": "SUCESSO_SEM_SERVICOS",
                "motivo": "Contribuinte sem módulo de Nota Fiscal",
                "tentativas": tentativa,
                "inicio": inicio,
                "fim": datetime.now(),
            }

        if proc.returncode == EXIT_CODE_CREDENCIAL_INVALIDA:
            return {
                "status": "SUCESSO",
                "motivo": "Credencial inválida na prefeitura (revisar planilha)",
                "tentativas": tentativa,
                "inicio": inicio,
                "fim": datetime.now(),
            }

        if proc.returncode == EXIT_CODE_CAPTCHA_TIMEOUT:
            ultimo_motivo = "Captcha não resolvido a tempo"
        else:
            ultimo_motivo = f"Falha execução (exit={proc.returncode})"

        print(f"Falha tentativa {tentativa}: {ultimo_motivo}")

    return {
        "status": "FALHA",
        "motivo": ultimo_motivo or "Falha desconhecida",
        "tentativas": MAX_TENTATIVAS,
        "inicio": inicio,
        "fim": datetime.now(),
    }


def salvar_report(rows):
    header = [
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
    with open(REPORT_PATH, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(header)
        for r in rows:
            w.writerow([
                r["inicio"].strftime("%Y-%m-%d %H:%M:%S"),
                r["fim"].strftime("%Y-%m-%d %H:%M:%S"),
                r["empresa"]["codigo"],
                r["empresa"]["razao_social"],
                r["empresa"]["cnpj"],
                r["empresa"]["segmento"],
                r["resultado"]["status"],
                r["resultado"]["motivo"],
                r["resultado"]["tentativas"],
            ])


def main():
    if not os.path.exists(CSV_EMPRESAS):
        print(f"Arquivo de empresas não encontrado: {CSV_EMPRESAS}")
        print("Informe EMPRESAS_ARQUIVO com caminho válido (.xlsx/.csv).")
        return

    empresas = carregar_empresas(CSV_EMPRESAS)
    print(f"Empresas carregadas: {len(empresas)}")

    concluidas = set()
    if CONTINUAR_DE_ONDE_PAROU:
        concluidas = carregar_report_existente(REPORT_PATH)
        if concluidas:
            print(f"Retomada ativa: {len(concluidas)} empresas já concluídas serão puladas.")

    resultados = []
    for empresa in empresas:
        cnpj_limpo = re.sub(r"\D", "", empresa.get("cnpj", ""))
        if cnpj_limpo and cnpj_limpo in concluidas:
            agora = datetime.now()
            resultados.append({
                "empresa": empresa,
                "resultado": {
                    "status": "SUCESSO",
                    "motivo": "Pulada por retomada (já concluída em execução anterior)",
                    "tentativas": 0,
                    "inicio": agora,
                    "fim": agora,
                },
                "inicio": agora,
                "fim": agora,
            })
            continue

        res = executar_empresa(empresa)
        resultados.append({"empresa": empresa, "resultado": res, "inicio": res["inicio"], "fim": res["fim"]})

    salvar_report(resultados)
    total = len(resultados)
    ok = sum(1 for r in resultados if r["resultado"]["status"] in {"SUCESSO", "SUCESSO_SEM_COMPETENCIA", "SUCESSO_SEM_SERVICOS"})
    falha = total - ok
    print("\n" + "=" * 80)
    print(f"Processamento finalizado. Total={total} | Sucesso={ok} | Falha={falha}")
    print(f"Report: {REPORT_PATH}")


if __name__ == "__main__":
    main()
