from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
ORQUESTRADOR_PATH = PROJECT_DIR / "orquestrador_empresas.py"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Worker do lote da Automação Prefeitura")
    parser.add_argument("--planilha", required=True)
    parser.add_argument("--apuracao", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--login-wait-seconds", default="120")
    parser.add_argument("--timeout-processo-main", default="1800")
    parser.add_argument("--continuar", default="1")
    parser.add_argument("--usar-checkpoint", default="1")
    return parser


def main() -> int:
    args = build_parser().parse_args()

    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["EMPRESAS_ARQUIVO"] = args.planilha
    env["APURACAO_REFERENCIA"] = args.apuracao
    env["OUTPUT_BASE_DIR"] = args.output_dir
    env["LOGIN_WAIT_SECONDS"] = str(args.login_wait_seconds)
    env["TIMEOUT_PROCESSO_MAIN"] = str(args.timeout_processo_main)
    env["CONTINUAR_DE_ONDE_PAROU"] = str(args.continuar)
    env["USAR_CHECKPOINT"] = str(args.usar_checkpoint)

    cmd = [sys.executable, "-u", str(ORQUESTRADOR_PATH)]

    print("=" * 80, flush=True)
    print("WORKER_INICIADO", flush=True)
    print(f"PLANILHA={args.planilha}", flush=True)
    print(f"APURACAO_REFERENCIA={args.apuracao}", flush=True)
    print(f"OUTPUT_BASE_DIR={args.output_dir}", flush=True)
    print("=" * 80, flush=True)

    proc = subprocess.Popen(
        cmd,
        cwd=str(PROJECT_DIR),
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        universal_newlines=True,
    )

    assert proc.stdout is not None
    try:
        for line in proc.stdout:
            print(line.rstrip("\n"), flush=True)
    except KeyboardInterrupt:
        proc.terminate()
        return 130

    return proc.wait()


if __name__ == "__main__":
    raise SystemExit(main())