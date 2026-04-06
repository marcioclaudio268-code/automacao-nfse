# GUI Rollout Controlado

Status: `READY`

Related:
- [backend_gui_contract.md](/c:/Users/Windows%2011/Desktop/NFS%20XML/app%20come%C3%A7o/automacao-nfse/docs/backend_gui_contract.md)
- [manual_homologation_checklist.md](/c:/Users/Windows%2011/Desktop/NFS%20XML/app%20come%C3%A7o/automacao-nfse/docs/manual_homologation_checklist.md)

Objetivo:
- colocar a GUI como frente principal de operacao interna em rollout controlado
- manter terminal apenas como fallback de transicao
- coletar atritos reais sem reabrir refatoracao ampla

Estado de partida:
- backend empacotado homologado para o objetivo atual
- contrato backend<->GUI congelado em `v1.1`
- GUI validada como cliente do contrato em smoke controlado
- risco remanescente principal concentrado em portal real e observabilidade geral do login/captcha

## Janela de rollout

Fase recomendada:
1. piloto com `1` ou `2` operadores
2. `3` a `5` execucoes reais
3. iniciar com lote pequeno
4. aumentar volume apenas se nao houver bloqueio operacional da GUI

Tamanho inicial de lote:
- preferir `1` empresa no primeiro uso de cada operador
- depois evoluir para `2` a `5` empresas por lote
- evitar lote grande enquanto houver coleta ativa de atritos

Criterio para expandir o uso:
- GUI inicia e acompanha o lote sem ambiguidade
- `manual_wait` fica claro e liberavel
- `report_execucao_empresas.csv` fecha corretamente
- logs e evidencias abrem pelo resolvedor da GUI
- nenhum atrito novo exige inferencia manual de estado escondido

## Checklist operacional

### Antes do turno

Confirmar:
- release da GUI correta
- release do orquestrador empacotado disponivel como fallback
- planilha de entrada revisada
- pasta de saida definida por run
- operador sabe a apuracao do dia e a competencia alvo esperada

Preparar:
- pasta de saida nova para o lote
- operador pronto para captcha e pausas manuais
- canal de registro de atritos do piloto

### Antes de iniciar cada lote

Na GUI:
- selecionar planilha correta
- validar planilha
- confirmar `Mes de apuracao`
- confirmar `Pasta base`
- anotar ou printar o inicio do lote se o piloto exigir evidencia

Conferir:
- a competencia esperada e o mes anterior ao de apuracao
- exemplo: `02/2026 -> 01.2026`

### Durante a execucao

Confirmar na GUI:
- empresa ativa identificavel na grade
- etapa atual legivel
- ultima mensagem compreensivel
- transicao para `EM_EXECUCAO`

Se houver pausa manual:
- ler o contexto mostrado na area manual
- executar a acao humana no portal
- clicar `Continuar etapa manual` apenas depois da acao no portal
- confirmar que a pausa some e o fluxo retoma

Se houver login/captcha manual:
- resolver o captcha
- clicar `Entrar` no portal
- confirmar que a automacao segue sozinha sem `Continuar etapa manual`
- usar `Continuar etapa manual` apenas para `MANUAL-FINAL` ou outra pausa formal equivalente

Se o lote fechar:
- conferir resultado final na grade
- abrir `Report`
- abrir `log manual` se houve pausa manual
- abrir `evidencias`

### Depois de cada lote

Registrar:
- quantidade de empresas
- status final por empresa
- se houve `MANUAL-LOGIN`
- se houve `MANUAL-FINAL`
- se o `MANUAL-LOGIN` retomou automaticamente como esperado
- se o operador precisou recorrer ao terminal
- qualquer atrito de entendimento da interface

Conferir:
- `report_execucao_empresas.csv` com header correto
- `motivo` e `acao_recomendada` legiveis para operacao
- pasta da empresa e competencia corretas
- a competencia aberta pela GUI confere com a apuracao

## Criterios de fallback

Regra geral:
- problema do portal real nao implica fallback automatico para terminal
- usar fallback apenas quando a GUI nao consegue cumprir seu papel de cliente do contrato

Continuar na GUI quando:
- o portal retorna `HTTP 500`
- o portal entra em loading infinito
- o captcha e rejeitado
- o operador precisa repetir tentativa humana

Motivo:
- nesses casos, o problema principal e externo ou operacional do portal
- a GUI ainda pode continuar sendo o ponto correto de controle e evidencia

Acionar fallback para terminal quando ocorrer qualquer um destes:
- GUI nao inicia o lote
- grade nao atualiza empresa ativa nem etapa de forma minimamente utilizavel
- `manual_wait` ocorre, mas `Continuar etapa manual` nao libera o backend
- login/captcha conclui no portal, mas a automacao nao retoma sozinha
- report nao fecha ou nao fica acessivel ao final
- botoes de artefato abrem caminho errado ou competencia errada
- GUI fecha, trava ou perde o controle do processo
- a interface entra em estado inconsistente e o operador nao sabe mais se o lote continua vivo

## Procedimento de fallback

Se a GUI falhar como frente operacional:
1. preservar a pasta de saida do lote atual
2. registrar screenshot da GUI e da pasta de saida
3. registrar se havia `manual_wait.json`
4. parar a GUI apenas depois de salvar a evidencia minima
5. decidir entre retomada ou rerun limpo

Usar retomada com o mesmo `output_base_dir` quando:
- o report e o checkpoint existem
- o estado do lote esta compreensivel
- a intencao e continuar do ponto ja processado

Usar rerun limpo com nova pasta de saida quando:
- a execucao ficou ambigua
- o operador nao confia no estado atual do lote
- a GUI caiu antes de deixar evidencia suficiente

No fallback terminal, manter:
- mesmo arquivo de empresas
- mesma `APURACAO_REFERENCIA`
- `CONTINUAR_DE_ONDE_PAROU=1`
- `USAR_CHECKPOINT=1`
- mesmo `OUTPUT_BASE_DIR` se a decisao for retomada

## Evidencia minima do piloto

Por lote:
- pasta de saida usada
- `report_execucao_empresas.csv`
- logs relevantes
- `manual_wait.json` se ainda existir no momento do incidente
- print curto quando houver falha da GUI ou atrito relevante

Por atrito:
- o que o operador tentou fazer
- o que a GUI mostrou
- se precisou fallback
- impacto real: atraso, repeticao, perda de evidencia ou bloqueio

## Lista curta de issues de endurecimento

Prioridade alta:
- `OBS-001`: persistir evidencia melhor de `MANUAL-LOGIN`
  - objetivo: deixar rastro confiavel em arquivo para login/captcha manual, nao so em evento transitivo
- `PORTAL-001`: classificar melhor falhas externas do portal
  - objetivo: distinguir `HTTP 500`, loading infinito e captcha rejeitado com mensagens operacionais
- `OBS-002`: persistir log estruturado de runtime do orquestrador
  - objetivo: nao depender apenas da janela ao vivo para reconstruir o fluxo

Prioridade media:
- `GUI-001`: reforcar feedback visual do login/captcha em andamento
  - objetivo: deixar explicito na interface a sequencia `resolver captcha -> clicar Entrar -> seguir automaticamente`
- `GUI-002`: resumir falha final com linguagem operacional
  - objetivo: exibir `motivo` e `acao_recomendada` de forma mais amigavel sem quebrar o contrato

## Criterio de encerramento do rollout controlado

Considerar a GUI pronta para operacao interna normal quando:
- o piloto completar `3` a `5` execucoes reais
- nao houver bloqueio recorrente de GUI
- o fallback terminal virar excecao e nao rotina
- os atritos restantes forem de portal real ou observabilidade, nao de integracao

## Template curto de acompanhamento

```text
Rollout GUI:
- Data:
- Operador:
- Release GUI:
- Release backend:
- Lote:
- Empresas:
- Apuracao:
- Output dir:
- Resultado geral:
- Manual wait ocorreu: sim/nao
- Fallback para terminal: sim/nao
- Atritos observados:
- Acoes de endurecimento abertas:
```
