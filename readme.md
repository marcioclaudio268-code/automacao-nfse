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

Opcional: para reduzir/aumentar parada por notas antigas fora da competência, use `LIMITE_HEURISTICA_FORA_ALVO` (padrão atual: `2`).

3. Para cada empresa (etapa atual):
   - o robô abre a URL de login da prefeitura;
   - preenche CNPJ/senha automaticamente;
   - você resolve o captcha e clica em `Entrar` manualmente;
   - após o login, o robô clica em `Nota Fiscal` e no segundo dashboard clica em `Lista Nota Fiscais` automaticamente.

4. O relatório consolidado será gerado na raiz:
   - `report_execucao_empresas.csv`
