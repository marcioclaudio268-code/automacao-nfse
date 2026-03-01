import csv
import json
import time
import os
import re
import subprocess
import sys
import unicodedata
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_EMPRESAS = os.environ.get("EMPRESAS_ARQUIVO", os.environ.get("EMPRESAS_CSV", "empresas.xlsx"))
REPORT_PATH = os.path.join(BASE_DIR, "report_execucao_empresas.txt")
REPORT_MULTIPLAS_PATH = os.path.join(BASE_DIR, "report_empresas_multiplas.txt")
CHECKPOINT_PATH = os.path.join(BASE_DIR, "checkpoint_execucao_empresas.json")
MAX_TENTATIVAS = int(os.environ.get("MAX_TENTATIVAS_EMPRESA", "3"))
LOGIN_WAIT_SECONDS = int(os.environ.get("LOGIN_WAIT_SECONDS", "120"))
TIMEOUT_PROCESSO_MAIN = int(os.environ.get("TIMEOUT_PROCESSO_MAIN", "1800"))
CONTINUAR_DE_ONDE_PAROU = os.environ.get("CONTINUAR_DE_ONDE_PAROU", "1").strip() == "1"
USAR_CHECKPOINT = os.environ.get("USAR_CHECKPOINT", "1").strip() == "1"
RETOMAR_POR_INDICE = os.environ.get("RETOMAR_POR_INDICE", "1").strip() == "1"
MSG_CAPTCHA_TIMEOUT = "CAPTCHA_NAO_RESOLVIDO_NO_TEMPO"
MSG_CAPTCHA_INCORRETO = "CAPTCHA_INCORRETO"
MSG_EMPRESA_MULTIPLA = "EMPRESA_MULTIPLA"
MSG_ALERTA_TOMADOS = "ALERTA_TOMADOS"
MSG_ALERTA_FECHAMENTO_TOMADOS = "ALERTA_FECHAMENTO_TOMADOS"
MSG_ALERTA_PRESTADOS = "ALERTA_PRESTADOS"
EXIT_CODE_CAPTCHA_TIMEOUT = 30
EXIT_CODE_SEM_COMPETENCIA = 40
EXIT_CODE_SEM_SERVICOS = 41
EXIT_CODE_CREDENCIAL_INVALIDA = 50
EXIT_CODE_TOMADOS_FALHA = 42
EXIT_CODE_CHROME_INIT_FALHA = 60
EXIT_CODE_EMPRESA_MULTIPLA = 61
MAX_RETRIES_CAPTCHA_INCORRETO = int(os.environ.get("MAX_RETRIES_CAPTCHA_INCORRETO", "2"))
STATUS_SUCESSO = {"SUCESSO", "SUCESSO_SEM_COMPETENCIA", "SUCESSO_SEM_SERVICOS"}


def _motivo_empresa_multipla(motivo: str) -> bool:
    txt = (motivo or "")
    return bool(re.search(r"EMPRESA_MULTIPLA|empresa\s*m[úu]ltipla|cadastros\s+relacionados", txt, re.I))


def append_txt(path: str, line: str):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(line.rstrip() + "\n")


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
            "Colunas obrigatÃƒÂ³rias nÃƒÂ£o encontradas. Esperado: CÃƒÂ³digo, RazÃƒÂ£o Social, CNPJ, Segmento, Senha Prefeitura"
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
        "Colunas obrigatÃƒÂ³rias nÃƒÂ£o encontradas nas primeiras linhas do XLSX. "
        "Esperado: CÃƒÂ³digo, RazÃƒÂ£o Social, CNPJ, Segmento, Senha Prefeitura"
    )


def carregar_empresas_xlsx(path_xlsx: str):
    try:
        from openpyxl import load_workbook
    except Exception as e:
        raise RuntimeError("Para usar .xlsx instale a dependÃƒÂªncia: pip install openpyxl") from e

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
        with open(path_report, "r", encoding="utf-8") as f:
            for line in f:
                if line.startswith("#") or not line.strip():
                    continue
                m_cnpj = re.search(r"CNPJ=(\d+)", line)
                m_status = re.search(r"STATUS=([^\s\|]+)", line)
                if not m_cnpj or not m_status:
                    continue
                cnpj = (m_cnpj.group(1) or "").strip()
                status = (m_status.group(1) or "").strip().upper()
                if status in STATUS_SUCESSO and cnpj:
                    concluidas.add(cnpj)
    except Exception:
        pass

    return concluidas

def carregar_rows_report_existente(path_report: str):
    rows = []
    if not os.path.exists(path_report):
        return rows

    def _get(k: str, line: str) -> str:
        m = re.search(rf"{re.escape(k)}=([^\|]+)", line)
        return (m.group(1).strip() if m else "")

    try:
        with open(path_report, "r", encoding="utf-8") as f:
            for line in f:
                if line.startswith("#") or not line.strip():
                    continue

                inicio_raw = _get("INICIO", line)
                fim_raw = _get("FIM", line)
                try:
                    inicio = datetime.strptime(inicio_raw, "%Y-%m-%d %H:%M:%S") if inicio_raw else datetime.now()
                except Exception:
                    inicio = datetime.now()
                try:
                    fim = datetime.strptime(fim_raw, "%Y-%m-%d %H:%M:%S") if fim_raw else inicio
                except Exception:
                    fim = inicio

                emp = {
                    "codigo": _get("CODIGO", line),
                    "razao_social": _get("RAZAO", line),
                    "cnpj": _get("CNPJ", line),
                    "segmento": _get("SEGMENTO", line),
                }
                status = _get("STATUS", line).upper()
                motivo = _get("MOTIVO", line)
                tentativas_raw = _get("TENTATIVAS", line)
                try:
                    tentativas = int(tentativas_raw) if tentativas_raw else 0
                except Exception:
                    tentativas = 0

                competencia = _get("COMPETENCIA", line)

                res = {
                    "status": status,
                    "motivo": motivo,
                    "tentativas": tentativas,
                    "inicio": inicio,
                    "fim": fim,
                }
                row = {"empresa": emp, "resultado": res, "inicio": inicio, "fim": fim}
                if competencia:
                    row["competencia"] = competencia
                rows.append(row)
    except Exception:
        pass

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


def salvar_checkpoint(path_checkpoint: str, processadas, ultimo_indice: int, tentativas: int = 12):
    """
    Windows: os.replace pode falhar com arquivo bloqueado (WinError 5).
    EstratÃ©gia:
      1) grava tmp Ãºnico
      2) tenta replace com retries (backoff leve)
      3) fallback: grava direto no destino
    """
    payload = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ultimo_indice": int(ultimo_indice),
        # mantÃ©m compatÃ­vel: lista ordenada e sem duplicatas
        "processadas": sorted(set(processadas)),
    }

    dir_ = os.path.dirname(path_checkpoint) or "."
    os.makedirs(dir_, exist_ok=True)

    pid = os.getpid()
    stamp = int(time.time() * 1000)
    tmp_path = f"{path_checkpoint}.{pid}.{stamp}.tmp"

    data = json.dumps(payload, ensure_ascii=False, indent=2)

    # 1) escreve tmp
    with open(tmp_path, "w", encoding="utf-8") as f:
        f.write(data)
        f.flush()
        try:
            os.fsync(f.fileno())
        except Exception:
            pass

    # 2) tenta replace com retry
    last_err = None
    for i in range(int(tentativas)):
        try:
            os.replace(tmp_path, path_checkpoint)
            return
        except PermissionError as e:
            last_err = e
            time.sleep(0.25 + i * 0.05)
        except OSError as e:
            last_err = e
            time.sleep(0.25 + i * 0.05)

    # 3) fallback: grava direto no destino
    try:
        with open(path_checkpoint, "w", encoding="utf-8") as f:
            f.write(data)
            f.flush()
            try:
                os.fsync(f.fileno())
            except Exception:
                pass
    finally:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass

    # opcional: logar o erro original (nÃ£o quebra a execuÃ§Ã£o)
    try:
        append_txt(REPORT_PATH, f"[WARN] Falha ao substituir checkpoint via os.replace: {last_err}")
    except Exception:
        pass




def executar_empresa(empresa: dict):
    inicio = datetime.now()
    ultimo_motivo = ""
    main_script = os.path.join(BASE_DIR, "main.py")

    tentativa_consumida = 0
    captcha_incorreto_retries = 0

    while tentativa_consumida < MAX_TENTATIVAS:
        tentativa = tentativa_consumida + 1
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
                capture_output=True,
                text=True,
                timeout=TIMEOUT_PROCESSO_MAIN,
            )
        except subprocess.TimeoutExpired:
            ultimo_motivo = f"Timeout de execucao do main.py (> {TIMEOUT_PROCESSO_MAIN}s)"
            print(f"Falha tentativa {tentativa}: {ultimo_motivo}")
            tentativa_consumida += 1
            continue

        rc = int(proc.returncode)
        saida = (proc.stdout or "") + "\n" + (proc.stderr or "")

        sucesso_por_codigo = {
            0: ("SUCESSO", "FECHADO"),
            EXIT_CODE_SEM_COMPETENCIA: ("SUCESSO", "FECHADO (sem movimento em Prestados na competência alvo)"),
            EXIT_CODE_SEM_SERVICOS: ("SUCESSO_SEM_SERVICOS", "Contribuinte sem modulo de Nota Fiscal"),
            EXIT_CODE_CREDENCIAL_INVALIDA: ("SUCESSO", "Revisar manualmente: credencial invÃ¡lida (usuÃ¡rio/senha) no portal"),
            EXIT_CODE_EMPRESA_MULTIPLA: ("SUCESSO", "Revisar manualmente: EMPRESA_MULTIPLA (Cadastros Relacionados)"),
        }

        if rc in sucesso_por_codigo:
            status, motivo = sucesso_por_codigo[rc]
            if rc == 0:
                alertas = []
                if MSG_ALERTA_TOMADOS in saida:
                    alertas.append("tomados")
                if MSG_ALERTA_FECHAMENTO_TOMADOS in saida:
                    alertas.append("fechamento_tomados")
                if MSG_ALERTA_PRESTADOS in saida:
                    alertas.append("prestados")
                if alertas:
                    motivo = f"{motivo} (com alertas: {', '.join(alertas)})"
            return {
                "status": status,
                "motivo": motivo,
                "tentativas": tentativa,
                "inicio": inicio,
                "fim": datetime.now(),
            }

        if rc == EXIT_CODE_CAPTCHA_TIMEOUT:
            # captcha digitado errado (Texto da imagem incorreto.) -> retry automÃ¡tico "grÃ¡tis"
            if MSG_CAPTCHA_INCORRETO in saida:
                captcha_incorreto_retries += 1
                ultimo_motivo = "Captcha incorreto (texto da imagem incorreto) - retry automatico"
                print(f"{ultimo_motivo} | retry {captcha_incorreto_retries}/{MAX_RETRIES_CAPTCHA_INCORRETO}")
                if captcha_incorreto_retries > MAX_RETRIES_CAPTCHA_INCORRETO:
                    ultimo_motivo = "Muitos captchas incorretos seguidos - consumindo tentativa"
                    print(ultimo_motivo)
                    tentativa_consumida += 1
                continue

            ultimo_motivo = "Captcha nao resolvido a tempo"
            print(f"Falha tentativa {tentativa}: {ultimo_motivo}")
            tentativa_consumida += 1
            continue

        if rc == EXIT_CODE_TOMADOS_FALHA:
            ultimo_motivo = "Falha na etapa de Servicos Tomados"
            print(f"Falha tentativa {tentativa}: {ultimo_motivo}")
            break

        if rc == EXIT_CODE_CHROME_INIT_FALHA:
            ultimo_motivo = "Falha na inicializacao do Chrome/WebDriver"
            print(f"Falha tentativa {tentativa}: {ultimo_motivo}")
            break

        resumo_saida = " ".join((proc.stderr or proc.stdout or "").strip().split())[:160]
        if resumo_saida:
            ultimo_motivo = f"Falha execucao (exit={rc}): {resumo_saida}"
        else:
            ultimo_motivo = f"Falha execucao (exit={rc})"
        print(f"Falha tentativa {tentativa}: {ultimo_motivo}")
        tentativa_consumida += 1

    return {
        "status": "FALHA",
        "motivo": ultimo_motivo or "Falha desconhecida",
        "tentativas": MAX_TENTATIVAS,
        "inicio": inicio,
        "fim": datetime.now(),
    }


def garantir_report_com_header(path_report: str):
    if os.path.exists(path_report):
        return
    os.makedirs(os.path.dirname(path_report) or ".", exist_ok=True)
    with open(path_report, "w", encoding="utf-8") as f:
        f.write("# Report execuÃ§Ã£o empresas (uma linha por empresa)\n")

def resetar_report(path_report: str):
    os.makedirs(os.path.dirname(path_report) or ".", exist_ok=True)
    with open(path_report, "w", encoding="utf-8") as f:
        f.write("# Report execuÃ§Ã£o empresas (uma linha por empresa)\n")

def append_report_row(path_report: str, row):
    garantir_report_com_header(path_report)

    emp = row.get("empresa") or {}
    res = row.get("resultado") or {}
    inicio = row.get("inicio")
    fim = row.get("fim")
    competencia = row.get("competencia") or res.get("competencia") or ""

    inicio_str = inicio.strftime("%Y-%m-%d %H:%M:%S") if hasattr(inicio, "strftime") else ""
    fim_str = fim.strftime("%Y-%m-%d %H:%M:%S") if hasattr(fim, "strftime") else ""

    motivo = (res.get("motivo") or "").replace("\n", " ").replace("\r", " ").replace("|", "/").strip()
    cnpj_limpo = re.sub(r"\D", "", emp.get("cnpj", "") or "")

    parts = [
        f"INICIO={inicio_str}",
        f"FIM={fim_str}",
        f"CODIGO={emp.get('codigo','')}",
        f"RAZAO={emp.get('razao_social','')}",
        f"CNPJ={cnpj_limpo}",
        f"SEGMENTO={emp.get('segmento','')}",
    ]
    if competencia:
        parts.append(f"COMPETENCIA={competencia}")

    parts.append(f"STATUS={res.get('status','')}")
    parts.append(f"TENTATIVAS={res.get('tentativas','')}")

    if motivo:
        parts.append(f"MOTIVO={motivo}")

    append_txt(path_report, " | ".join(parts))


def garantir_report_multiplas_com_header(path_report: str):
    if os.path.exists(path_report):
        return
    os.makedirs(os.path.dirname(path_report) or '.', exist_ok=True)
    with open(path_report, 'w', encoding='utf-8') as f:
        f.write('# Empresas mÃºltiplas (Cadastros Relacionados) - revisar manualmente\n')

def append_report_multiplas_row(path_report: str, row):
    garantir_report_multiplas_com_header(path_report)
    emp = row.get('empresa') or {}
    res = row.get('resultado') or {}
    inicio = row.get('inicio')
    fim = row.get('fim')
    inicio_str = inicio.strftime('%Y-%m-%d %H:%M:%S') if hasattr(inicio, 'strftime') else ''
    fim_str = fim.strftime('%Y-%m-%d %H:%M:%S') if hasattr(fim, 'strftime') else ''
    motivo = (res.get('motivo') or '').replace('\n',' ').replace('\r',' ').replace('|','/').strip()
    cnpj_limpo = re.sub(r'\D', '', emp.get('cnpj', '') or '')
    parts = [
        f'INICIO={inicio_str}',
        f'FIM={fim_str}',
        f'CODIGO={emp.get("codigo","")}',
        f'RAZAO={emp.get("razao_social","")}',
        f'CNPJ={cnpj_limpo}',
        f'SEGMENTO={emp.get("segmento","")}',
        'STATUS=EMPRESA_MULTIPLA',
    ]
    if motivo:
        parts.append(f'MOTIVO={motivo}')
    append_txt(path_report, ' | '.join(parts))

def main():
    if not os.path.exists(CSV_EMPRESAS):
        print(f"Arquivo de empresas nÃƒÂ£o encontrado: {CSV_EMPRESAS}")
        print("Informe EMPRESAS_ARQUIVO com caminho vÃƒÂ¡lido (.xlsx/.csv).")
        return

    empresas = carregar_empresas(CSV_EMPRESAS)
    print(f"Empresas carregadas: {len(empresas)}")

    concluidas = set()
    if CONTINUAR_DE_ONDE_PAROU:
        concluidas = carregar_report_existente(REPORT_PATH)
        if concluidas:
            print(f"Retomada ativa: {len(concluidas)} empresas jÃƒÂ¡ concluÃƒÂ­das serÃƒÂ£o puladas.")

    processadas_checkpoint = set()
    if USAR_CHECKPOINT:
        ck = carregar_checkpoint(CHECKPOINT_PATH)
        processadas_checkpoint = ck["processadas"]
        ultimo_indice_checkpoint = int(ck.get("ultimo_indice", -1))
        if processadas_checkpoint:
            print(f"Checkpoint ativo: {len(processadas_checkpoint)} empresas jÃƒÂ¡ processadas serÃƒÂ£o puladas.")

    start_index = 0
    if USAR_CHECKPOINT and RETOMAR_POR_INDICE and ultimo_indice_checkpoint >= 0:
        start_index = min(max(0, ultimo_indice_checkpoint + 1), len(empresas))
        if start_index > 0:
            print(f"Retomando a partir do Ã­ndice {start_index} (checkpoint ultimo_indice={ultimo_indice_checkpoint}).")
    ja_processadas = set(concluidas) | set(processadas_checkpoint)

    resultados = carregar_rows_report_existente(REPORT_PATH) if CONTINUAR_DE_ONDE_PAROU else []
    if CONTINUAR_DE_ONDE_PAROU:
        garantir_report_com_header(REPORT_PATH)
    else:
        resetar_report(REPORT_PATH)
    cnpjs_ja_no_report = {re.sub(r"\D", "", r["empresa"].get("cnpj", "")) for r in resultados if r.get("empresa")}
    atualizar_resumo_falhas_parcial(resultados)
    for idx, empresa in enumerate(empresas[start_index:], start=start_index):
        cnpj_limpo = re.sub(r"\D", "", empresa.get("cnpj", ""))

        if cnpj_limpo and cnpj_limpo in ja_processadas:
            if cnpj_limpo in cnpjs_ja_no_report:
                continue
            agora = datetime.now()
            resultado = {
                "status": "SUCESSO",
                "motivo": "Pulada por retomada/checkpoint (jÃƒÂ¡ processada)",
                "tentativas": 0,
                "inicio": agora,
                "fim": agora,
            }
            row = {"empresa": empresa, "resultado": resultado, "inicio": agora, "fim": agora}
            resultados.append(row)
            cnpjs_ja_no_report.add(cnpj_limpo)
            append_report_row(REPORT_PATH, row)
            # Report dedicado: empresas com mÃºltiplos cadastros (bloqueio para automaÃ§Ã£o)
            try:
                motivo_txt = row.get("resultado", {}).get("motivo", "")
                if _motivo_empresa_multipla(motivo_txt):
                    append_report_multiplas_row(REPORT_MULTIPLAS_PATH, row)
            except Exception:
                pass
            atualizar_resumo_falhas_parcial(resultados)
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
        try:
            if _motivo_empresa_multipla(res.get("motivo", "")):
                append_report_multiplas_row(REPORT_MULTIPLAS_PATH, row)
        except Exception:
            pass
        atualizar_resumo_falhas_parcial(resultados)
        if USAR_CHECKPOINT:
            salvar_checkpoint(CHECKPOINT_PATH, processadas_checkpoint, idx)

        if cnpj_limpo and res.get("status") in STATUS_SUCESSO:
            ja_processadas.add(cnpj_limpo)
            if USAR_CHECKPOINT:
                processadas_checkpoint.add(cnpj_limpo)
                salvar_checkpoint(CHECKPOINT_PATH, processadas_checkpoint, idx)

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

RESUMO_FALHAS_PATH = os.environ.get(
    "RESUMO_FALHAS_PATH",
    os.path.join(BASE_DIR, "resumo_falhas.txt"),
)

def _nome_empresa(empresa: dict) -> str:
    razao = empresa.get("razao_social") or empresa.get("razao") or empresa.get("RazÃ£o Social") or ""
    cnpj = re.sub(r"\D", "", empresa.get("cnpj", "") or "")
    codigo = str(empresa.get("codigo") or "").strip()
    base = razao.strip() if razao else "(SEM RAZAO)"
    if codigo:
        base = f"{codigo} - {base}"
    if cnpj:
        base = f"{base} ({cnpj})"
    return base

def escrever_resumo_falhas(path_txt: str, resultados: list):
    falhas = []
    revisar = []

    for row in resultados:
        empresa = row.get("empresa") or {}
        res = row.get("resultado") or {}
        status = (res.get("status") or "").upper().strip()
        motivo = (res.get("motivo") or "").strip()
        nome = _nome_empresa(empresa)

        if status == "FALHA":
            falhas.append((nome, motivo))
            continue

        # Regra A: SUCESSO mas precisa revisÃ£o (ex.: credencial invÃ¡lida)
        if status.startswith("SUCESSO"):
            if re.search(r"(revisar|credencial|inv[aÃ¡]lid|senha inv[aÃ¡]lida|usuario inv[aÃ¡]lid)", motivo, re.I):
                revisar.append((nome, motivo))

    linhas = []
    linhas.append("EMPRESAS COM FALHAS")
    linhas.append("")

    if not falhas:
        linhas.append("Nenhuma falha.")
    else:
        for nome, motivo in falhas:
            linhas.append(f"{nome} - FALHA ({motivo})")

    linhas.append("")
    linhas.append("EMPRESAS PARA REVISAR")
    linhas.append("")

    if not revisar:
        linhas.append("Nenhuma empresa para revisar.")
    else:
        for nome, motivo in revisar:
            linhas.append(f"{nome} - REVISAR ({motivo})")

    with open(path_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas) + "\n")


def atualizar_resumo_falhas_parcial(resultados: list):
    try:
        escrever_resumo_falhas(RESUMO_FALHAS_PATH, resultados)
    except Exception as e:
        print(f"[WARN] Falha ao atualizar resumo parcial: {e}")

if __name__ == "__main__":
    try:
        main()
    finally:
        try:
            resultados_finais = carregar_rows_report_existente(REPORT_PATH)
            escrever_resumo_falhas(RESUMO_FALHAS_PATH, resultados_finais)
            print(f"Resumo de falhas gerado: {RESUMO_FALHAS_PATH}")
        except Exception as e:
            print(f"[WARN] Falha ao gerar resumo de falhas: {e}")
