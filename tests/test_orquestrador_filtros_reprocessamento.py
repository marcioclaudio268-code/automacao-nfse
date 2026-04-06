import csv

import orquestrador_empresas as oq


def _empresa(codigo: str, indice_lista: int, linha_planilha: int):
    return {
        "codigo": codigo,
        "razao_social": f"Empresa {codigo}",
        "cnpj": f"12.345.678/0001-{int(codigo) % 100:02d}",
        "segmento": "SERVICOS",
        "senha_prefeitura": "segredo",
        "indice_lista": indice_lista,
        "linha_planilha": linha_planilha,
    }


def _escrever_resumo_csv(path, linhas):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, delimiter=";", fieldnames=oq.RESUMO_HEADER)
        writer.writeheader()
        writer.writerows(linhas)


def test_classificar_erro_execucao_cobre_categorias_e_fallback():
    assert oq.classificar_erro_execucao("REVISAO_MANUAL", "Credencial invalida no portal") == "LOGIN_INVALIDO"
    assert oq.classificar_erro_execucao("FALHA", "Falha no download do arquivo XML") == "ARQUIVO"
    assert oq.classificar_erro_execucao("FALHA", "Captcha nao resolvido a tempo") == "CAPTCHA"
    assert oq.classificar_erro_execucao("FALHA", "Timeout ao falar com o portal") == "TIMEOUT"
    assert oq.classificar_erro_execucao("FALHA", "Falha na inicializacao do Chrome/WebDriver") == "ERRO_PORTAL"
    assert oq.classificar_erro_execucao("SUCESSO_SEM_SERVICOS", "Contribuinte sem modulo de Nota Fiscal") == "SEM_MOVIMENTO"
    assert oq.classificar_erro_execucao("FALHA", "Mensagem sem contexto") == "DESCONHECIDO"
    assert oq.classificar_erro_execucao("", "", "CAPTCHA_TIMEOUT") == "CAPTCHA"


def test_filtrar_empresas_por_criterios_combina_codigo_e_erro():
    empresas = [_empresa("101", 1, 2), _empresa("115", 2, 3), _empresa("833", 3, 4)]
    registros_fonte = [
        {
            "empresa": {"codigo": "101"},
            "resultado": {"status": "REVISAO_MANUAL", "motivo": "Credencial invalida no portal", "erro_tipo": "LOGIN_INVALIDO"},
        },
        {
            "empresa": {"codigo": "115"},
            "resultado": {"status": "FALHA", "motivo": "Falha no download do XML", "erro_tipo": "ARQUIVO"},
        },
        {
            "empresa": {"codigo": "833"},
            "resultado": {"status": "REVISAO_MANUAL", "motivo": "Credencial invalida no portal", "erro_tipo": "LOGIN_INVALIDO"},
        },
    ]

    selecionadas = oq.filtrar_empresas_por_criterios(
        empresas,
        empresas_explicitadas=("115", "833"),
        tipos_erro=("LOGIN_INVALIDO",),
        registros_fonte=registros_fonte,
    )

    assert [empresa["codigo"] for empresa in selecionadas] == ["833"]

    selecionadas_empresas = oq.filtrar_empresas_por_criterios(empresas, empresas_explicitadas=("115", "833"))
    assert [empresa["codigo"] for empresa in selecionadas_empresas] == ["115", "833"]


def test_main_reprocessa_lote_com_faixa_empresas_e_erro(monkeypatch, tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    empresas_arquivo = tmp_path / "empresas.csv"
    empresas_arquivo.write_text("stub", encoding="utf-8")

    empresas = [_empresa("101", 1, 2), _empresa("115", 2, 3), _empresa("833", 3, 4), _empresa("900", 4, 5)]
    resumo_fonte_path = output_dir / "resumo_execucao_empresas__lote_002_004.csv"
    _escrever_resumo_csv(
        resumo_fonte_path,
        [
            {
                "indice_lista": "1",
                "linha_planilha": "2",
                "empresa": "Empresa 101",
                "codigo": "101",
                "cnpj": "12.345.678/0001-01",
                "competencia": "01.2026",
                "status_execucao": "REVISAO_MANUAL",
                "teve_iss": "NAO",
                "teve_prestados": "NAO",
                "teve_tomados": "NAO",
                "erro_tipo": "LOGIN_INVALIDO",
                "erro_resumo": "Credencial invalida no portal",
            },
            {
                "indice_lista": "2",
                "linha_planilha": "3",
                "empresa": "Empresa 115",
                "codigo": "115",
                "cnpj": "12.345.678/0001-15",
                "competencia": "01.2026",
                "status_execucao": "FALHA",
                "teve_iss": "NAO",
                "teve_prestados": "NAO",
                "teve_tomados": "NAO",
                "erro_tipo": "ARQUIVO",
                "erro_resumo": "Falha no download do XML",
            },
            {
                "indice_lista": "3",
                "linha_planilha": "4",
                "empresa": "Empresa 833",
                "codigo": "833",
                "cnpj": "12.345.678/0001-33",
                "competencia": "01.2026",
                "status_execucao": "REVISAO_MANUAL",
                "teve_iss": "NAO",
                "teve_prestados": "NAO",
                "teve_tomados": "NAO",
                "erro_tipo": "LOGIN_INVALIDO",
                "erro_resumo": "Credencial invalida no portal",
            },
            {
                "indice_lista": "4",
                "linha_planilha": "5",
                "empresa": "Empresa 900",
                "codigo": "900",
                "cnpj": "12.345.678/0001-00",
                "competencia": "01.2026",
                "status_execucao": "REVISAO_MANUAL",
                "teve_iss": "NAO",
                "teve_prestados": "NAO",
                "teve_tomados": "NAO",
                "erro_tipo": "LOGIN_INVALIDO",
                "erro_resumo": "Credencial invalida no portal",
            },
        ],
    )

    monkeypatch.setattr(oq, "CSV_EMPRESAS", str(empresas_arquivo))
    monkeypatch.setattr(oq, "OUTPUT_BASE_DIR", str(output_dir))
    monkeypatch.setattr(oq, "RESUMO_USA_XLSX", False)
    monkeypatch.setattr(oq, "CONTINUAR_DE_ONDE_PAROU", True)
    monkeypatch.setattr(oq, "USAR_CHECKPOINT", True)
    monkeypatch.setattr(oq, "carregar_empresas", lambda _: list(empresas))
    monkeypatch.setenv("EMPRESA_INICIO", "2")
    monkeypatch.setenv("EMPRESA_FIM", "4")
    monkeypatch.setenv("EMPRESAS", "115,833,900")
    monkeypatch.setenv("FILTRAR_ERRO_TIPO", "LOGIN_INVALIDO")

    executadas = []
    rows_report = []

    monkeypatch.setattr(
        oq,
        "executar_empresa",
        lambda empresa: executadas.append(empresa["codigo"]) or {
            "status": "SUCESSO",
            "motivo": "OK",
            "tentativas": 1,
            "acao_recomendada": "",
            "inicio": oq.datetime.now(),
            "fim": oq.datetime.now(),
        },
    )
    monkeypatch.setattr(oq, "append_report_row", lambda *_args, **_kwargs: rows_report.append(_args[1]))
    monkeypatch.setattr(oq, "resetar_report", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(oq, "garantir_report_com_header", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(oq, "carregar_report_existente", lambda *_args, **_kwargs: {"101"})
    monkeypatch.setattr(oq, "carregar_checkpoint", lambda *_args, **_kwargs: {"processadas": {"101"}, "ultimo_indice": 0})
    monkeypatch.setattr(oq, "salvar_checkpoint", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(oq, "salvar_resumo_execucao", lambda *args, **kwargs: str(output_dir / "resumo_execucao_empresas__lote_002_004.csv"))

    oq.main()

    assert executadas == ["833", "900"]
    assert [row["empresa"]["codigo"] for row in rows_report] == ["833", "900"]
