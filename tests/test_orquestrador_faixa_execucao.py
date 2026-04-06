import pytest

import orquestrador_empresas as oq


def test_carregar_empresas_csv_preserva_metadados(tmp_path):
    csv_path = tmp_path / "empresas.csv"
    csv_path.write_text(
        "Codigo;Razao Social;CNPJ;Segmento;Senha Prefeitura\n"
        "1;Empresa A;12.345.678/0001-90;SERVICOS;segredo1\n"
        "2;Empresa B;98.765.432/0001-10;COMERCIO;segredo2\n",
        encoding="utf-8",
    )

    empresas = oq.carregar_empresas(str(csv_path))

    assert [empresa["indice_lista"] for empresa in empresas] == [1, 2]
    assert [empresa["linha_planilha"] for empresa in empresas] == [2, 3]


def test_resolver_faixa_execucao_sem_variaveis_retorna_lista_integral(monkeypatch):
    empresas = [
        {"indice_lista": 1, "linha_planilha": 10},
        {"indice_lista": 2, "linha_planilha": 11},
    ]

    monkeypatch.delenv("EMPRESA_INICIO", raising=False)
    monkeypatch.delenv("EMPRESA_FIM", raising=False)

    inicio, fim, filtradas = oq.resolver_faixa_execucao(empresas)

    assert inicio is None
    assert fim is None
    assert filtradas == empresas


def test_resolver_faixa_execucao_filtra_por_indice_lista(monkeypatch):
    empresas = [
        {"indice_lista": 1, "linha_planilha": 10},
        {"indice_lista": 2, "linha_planilha": 11},
        {"indice_lista": 3, "linha_planilha": 12},
        {"indice_lista": 4, "linha_planilha": 13},
    ]

    monkeypatch.setenv("EMPRESA_INICIO", "2")
    monkeypatch.setenv("EMPRESA_FIM", "3")

    inicio, fim, filtradas = oq.resolver_faixa_execucao(empresas)

    assert inicio == 2
    assert fim == 3
    assert [empresa["indice_lista"] for empresa in filtradas] == [2, 3]
    assert [empresa["linha_planilha"] for empresa in filtradas] == [11, 12]


def test_resolver_paths_execucao_usa_sufixo_de_lote(tmp_path):
    report_path, checkpoint_path = oq.resolver_paths_execucao(str(tmp_path), 1, 100)

    assert report_path.endswith("report_execucao_empresas__lote_001_100.csv")
    assert checkpoint_path.endswith("checkpoint_execucao_empresas__lote_001_100.json")


def test_resolver_paths_execucao_sem_faixa_mantem_nomes_base(tmp_path):
    report_path, checkpoint_path = oq.resolver_paths_execucao(str(tmp_path), None, None)

    assert report_path.endswith("report_execucao_empresas.csv")
    assert checkpoint_path.endswith("checkpoint_execucao_empresas.json")


def test_resolver_faixa_execucao_rejeita_intervalo_invalido(monkeypatch):
    monkeypatch.setenv("EMPRESA_INICIO", "5")
    monkeypatch.setenv("EMPRESA_FIM", "3")

    with pytest.raises(ValueError, match="nao pode ser maior"):
        oq.resolver_faixa_execucao([{"indice_lista": 1}])
