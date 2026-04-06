\# 🤖 Automação NFSe – Prefeitura (Padrão Nacional)



\## Objetivo



Este projeto automatiza o download em lote dos XMLs de Notas Fiscais de Serviço (NFSe) diretamente do portal da prefeitura, evitando a necessidade de baixar manualmente nota por nota.



O sistema:



\* abre o Chrome com perfil exclusivo

\* o usuário apenas faz login manualmente

\* o robô seleciona cada nota automaticamente

\* seleciona "NF Nacional"

\* clica em "Visualizar XML"

\* baixa o arquivo

\* fecha o modal corretamente

\* passa para a próxima nota

\* organiza por empresa/ano/mês

\* registra em log para não repetir downloads



---



\## IMPORTANTE (como usar)



1\. Execute:



```

python main.py

```



2\. O Chrome abrirá.



3\. Faça login manualmente no portal da prefeitura.



4\. Navegue até:



```

Nota Fiscal → Lista Nota Fiscais

```



5\. Aguarde 40 segundos.



Após isso o robô assumirá automaticamente.



---



\## Estrutura de Pastas



```

automacao-nfse/

│

├── main.py

├── log.csv

├── downloads/

│   └── EMPRESA/

│       └── ANO/

│           └── MES/

│               └── arquivos.xml

```



---



\## Tecnologias



\* Python 3.11+

\* Selenium

\* ChromeDriver automático

\* XML parsing



---



\## Observações importantes



• O site da prefeitura bloqueia login automatizado

→ por isso o login SEMPRE é manual.



• O sistema utiliza perfil exclusivo do Chrome:



```

C:\\ChromeRobotProfile

```



• Não usar o Chrome pessoal.



---



\## Problemas conhecidos



O site ocasionalmente retorna:

HTTP 502



O robô possui tratamento de falha e tenta novamente automaticamente.



---



## Execucao em lote por faixa



Para rodar a lista inteira, sem faixa, use o orquestrador:



```powershell
python orquestrador_empresas.py
```



Para processar apenas um lote da planilha:



```powershell
$env:EMPRESA_INICIO = "1"
$env:EMPRESA_FIM = "100"
python orquestrador_empresas.py
```



Se `EMPRESA_INICIO` e `EMPRESA_FIM` ficarem vazios, o comportamento continua igual ao atual: todas as empresas validas da planilha sao processadas.



## Resumo geral por empresa

Ao final de cada execucao, o orquestrador gera um resumo consolidado por empresa na mesma pasta de saida do lote.

Sem faixa:

- `resumo_execucao_empresas.xlsx`
- se `openpyxl` nao estiver disponivel no ambiente, o fallback e `.csv`

Com faixa:

- `resumo_execucao_empresas__lote_001_100.xlsx`
- se `openpyxl` nao estiver disponivel no ambiente, o fallback e `.csv`

Exemplo em PowerShell:

```powershell
$env:EMPRESA_INICIO = "1"
$env:EMPRESA_FIM = "100"
python orquestrador_empresas.py
```



## Reprocessamento seletivo

O orquestrador aceita filtros para reexecutar somente subconjuntos uteis.

- `EMPRESAS` usa o `codigo` da planilha como identificador principal, separado por virgula.
- `FILTRAR_ERRO_TIPO` aceita: `LOGIN_INVALIDO`, `ARQUIVO`, `CAPTCHA`, `TIMEOUT`, `ERRO_PORTAL`, `SEM_MOVIMENTO` e `DESCONHECIDO`.
- A ordem aplicada e: faixa -> `EMPRESAS` -> `FILTRAR_ERRO_TIPO`.
- O filtro por erro usa o `resumo_execucao_empresas` do mesmo escopo e, se ele nao existir, faz fallback para `report_execucao_empresas`.

Exemplos em PowerShell:

```powershell
$env:EMPRESAS = "101,115,833"
python orquestrador_empresas.py
```


## Saidas gerais centralizadas

Os arquivos continuam sendo salvos por empresa e, depois disso, sao espelhados em uma pasta geral para consolidacao operacional.

- `SAIDAS_GERAIS/ISS/<competencia>/`
- `SAIDAS_GERAIS/XML_TOMADOS/<competencia>/`

Exemplos de nome final:

- `SAIDAS_GERAIS/ISS/02.2026/00833_ABS_REPRESENTACOES_ISS_02-2026.pdf`
- `SAIDAS_GERAIS/XML_TOMADOS/02.2026/00833_ABS_REPRESENTACOES_SERVICOS_TOMADOS_02-2026.xml`

O nome espelhado usa, quando disponivel:

- codigo da empresa
- slug ou nome curto da empresa
- tipo do documento
- competencia
- extensao original

O salvamento por empresa continua sendo a fonte primaria de verdade; o espelhamento e secundario e nao substitui o arquivo original.

```powershell
$env:FILTRAR_ERRO_TIPO = "LOGIN_INVALIDO"
python orquestrador_empresas.py
```

```powershell
$env:EMPRESA_INICIO = "1"
$env:EMPRESA_FIM = "100"
$env:EMPRESAS = "101,115,833"
$env:FILTRAR_ERRO_TIPO = "LOGIN_INVALIDO"
python orquestrador_empresas.py
```



