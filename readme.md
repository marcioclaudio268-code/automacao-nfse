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






## Execução em lote (semi-automático)

Para processar várias empresas com captcha manual e relatório final CSV:

1. Prepare `empresas.xlsx` (recomendado) **ou** `empresas.csv` (separador `;`) com colunas:
   - Código
   - Razão Social
   - CNPJ
   - Segmento
   - Senha Prefeitura

2. Execute:

```bash
python orquestrador_empresas.py
```

Opcional: para usar outro arquivo, defina `EMPRESAS_ARQUIVO` (ex.: `empresas.csv`).

Opcional: para retomar de onde parou e pular empresas já concluídas no report anterior, use `CONTINUAR_DE_ONDE_PAROU=1` (padrão).

Opcional: para checkpoint incremental durante a execução, use `USAR_CHECKPOINT=1` (padrão). O orquestrador salva `checkpoint_execucao_empresas.json` e atualiza o report a cada empresa processada.

Observação: ao encontrar a primeira nota mais antiga que a competência alvo, a empresa é encerrada como sem competência (ordem decrescente).

3. Para cada empresa (etapa atual):
   - o robô abre a URL de login da prefeitura;
   - preenche CNPJ/senha automaticamente;
   - você resolve o captcha e clica em `Entrar` manualmente;
   - após o login, o robô clica em `Nota Fiscal` e no segundo dashboard clica em `Lista Nota Fiscais` automaticamente.

4. O relatório consolidado será gerado na raiz:
   - `report_execucao_empresas.csv`
   - quando não houver notas da competência alvo, o status registrado será `SUCESSO_SEM_COMPETENCIA` (sem retries desnecessários).
   - quando houver usuário/senha inválidos na prefeitura, a empresa é marcada como `SUCESSO` com motivo no report para revisão da planilha.
   - quando o contribuinte não possui módulo `Nota Fiscal`, a empresa é marcada como `SUCESSO_SEM_SERVICOS`.
   - quando a lista carregar sem checkboxes (empresa sem notas selecionáveis), também finaliza como `SUCESSO_SEM_COMPETENCIA`; se a primeira Data Emissão já vier mais antiga que a competência alvo, encerra imediatamente.
