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



