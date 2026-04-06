# Backend GUI Contract

Status: `FROZEN`

Version: `v1.1`

Purpose: document the frozen integration contract that the GUI consumes from the current backend implementation validated in `MAN-001`.

Validation status:
- accepted on `2026-03-11` for GUI integration
- validated with real run `C:\AAAAAAAAA\20260311_090718`
- rollout update on `2026-03-11`: `MANUAL-LOGIN` changed from formal pause to informational stage with automatic resume after portal submit
- `MANUAL-FINAL` remains the only formal manual gate in normal operation

Scope:
- Desktop GUI launching and monitoring the orchestrator.
- Runtime progress updates.
- Manual pause and resume handshake.
- Final result ingestion from `report_execucao_empresas.csv`.
- Artifact location rules.

Non-scope:
- Free-form console wording not covered below.
- Internal Selenium behavior.
- Site-specific DOM details from the prefeitura portal.

## Integration Boundary

Stable backend entrypoint for the GUI:
- `orquestrador_empresas.py` in dev mode.
- `AutomacaoNFSe-Orquestrador.exe` in packaged mode.

The GUI must treat the orchestrator as the integration boundary.

The GUI must not call `main.py` directly.

The GUI must treat `report_execucao_empresas.csv` as the source of truth for final outcome.

The GUI may use runtime events only for live progress and user feedback.

## Inputs

Current execution inputs map to `RuntimeConfig`.

Required:
- `empresas_arquivo`: absolute or relative path to `.xlsx` or `.csv`.
- `apuracao_referencia`: `MM/AAAA`.
- `output_base_dir`: writable directory.

Optional:
- `continuar_de_onde_parou`: `true|false`, default `true`.
- `usar_checkpoint`: `true|false`, default `true`.
- `login_wait_seconds`: integer `> 0`, default `120`.
- `timeout_processo_main`: integer `> 0`, default `1800`.

Boolean handling rule:
- inside the GUI, `continuar_de_onde_parou` and `usar_checkpoint` are boolean values.
- on the process boundary, the controller must serialize them as `1` for true and `0` for false.
- in the current backend implementation, only the string `1` is interpreted as true.
- the GUI must not send textual boolean values such as `true`, `false`, `yes` or `no` to the backend process.

Current validation rules:
- input file must exist.
- file extension must be `.xlsx` or `.csv`.
- `apuracao_referencia` must match `MM/AAAA`.
- `output_base_dir` must be non-empty and writable.
- `login_wait_seconds` and `timeout_processo_main` must be positive.

Environment variables currently set by the GUI controller:
- `EMPRESAS_ARQUIVO`
- `APURACAO_REFERENCIA`
- `OUTPUT_BASE_DIR`
- `EXECUTION_CONTROL_DIR`
- `LOGIN_WAIT_SECONDS`
- `TIMEOUT_PROCESSO_MAIN`
- `CONTINUAR_DE_ONDE_PAROU` as `1|0`
- `USAR_CHECKPOINT` as `1|0`
- `PYTHONUNBUFFERED=1`

## Status Model

Final result values currently used by the GUI and report:
- `AGUARDANDO`
- `EM_EXECUCAO`
- `SUCESSO`
- `SUCESSO_SEM_COMPETENCIA`
- `SUCESSO_SEM_SERVICOS`
- `REVISAO_MANUAL`
- `FALHA`
- `INTERROMPIDO`

Current recommended action values from the report:
- `""`
- `CONFERIR_SENHA`
- `REPROCESSAR`
- `ABRIR_DEBUG`

Current final-result semantics:
- `SUCESSO`: backend run completed successfully.
- `SUCESSO_SEM_COMPETENCIA`: no notas in target competence.
- `SUCESSO_SEM_SERVICOS`: company has no Nota Fiscal module.
- `REVISAO_MANUAL`: backend completed with manual follow-up required, currently used for invalid portal credential.
- `FALHA`: orchestrator or company execution failed.
- `INTERROMPIDO`: stop requested by the user while execution was active.

Current minimum valid result transition:
- `AGUARDANDO -> EM_EXECUCAO -> {SUCESSO|SUCESSO_SEM_COMPETENCIA|SUCESSO_SEM_SERVICOS|REVISAO_MANUAL|FALHA|INTERROMPIDO}`

The GUI must not infer additional final states on its own.

## Progress Model

Current runtime `etapa` values are human-readable strings, not strict enums.

Observed current values:
- `Preparando empresa`
- `Login`
- `Aguardando captcha`
- `Tomados`
- `Prestados`
- `Livros / Manual`
- `Concluido`
- `Erro`
- `Interrompido`

Important:
- `ui.models.EtapaExecucao` is a useful canonical list, but current live payloads still use human-readable labels.
- The GUI must display these values as received.
- The GUI must not assume `etapa` is normalized to enum literal names such as `PREPARANDO` or `LIVROS_MANUAL`.

## Runtime Events

Current event parser emits `ParsedEvent.kind` values:
- `empresa_atual`
- `stderr`
- `manual_wait`
- `manual_resumed`
- `interrompido`
- `erro`
- `final`
- `etapa`
- `log`
- `empty`

Current minimal event payload fields:
- `kind`
- `raw`
- `codigo`
- `razao`
- `etapa`
- `mensagem`

Current event handling contract for the GUI:
- `empresa_atual`: identifies the active company and moves it to `EM_EXECUCAO`.
- `manual_wait`: requires GUI reaction only for formal manual gates such as `MANUAL` and `MANUAL-FINAL`.
- `manual_resumed`: requires GUI reaction only for formal manual gates such as `MANUAL` and `MANUAL-FINAL`.
- `interrompido`: updates active row to `INTERROMPIDO`.
- `etapa`, `erro`, `log`, `final`: informational progress updates.
- `stderr`: observability only. Not a final-state contract.

Current heuristics used by `parse_log_line`:
- `Empresa <codigo> - <razao>` starts a new active company.
- `[MANUAL-LOGIN]` maps to an informational login/captcha stage with `etapa = Aguardando captcha`.
- `[MANUAL]` or `[MANUAL-FINAL]` lines map to manual-step events.
- text containing `captcha` maps to `Aguardando captcha`.
- text containing `tomados` maps to `Tomados`.
- text containing `prestados` maps to `Prestados`.
- text containing `livro`, `guia` or `manual` maps to `Livros / Manual`.

The GUI must not expand this parser logic on its own.

If parser logic changes, backend code must change first and the contract must be updated.

## Manual Pause Contract

Manual control directory root:
- `<output_base_dir>/_control/run_<timestamp>_<id>/`

Per-company manual control directory:
- `<control_dir>/<empresa_pasta>/`

Current files:
- `manual_wait.json`
- `continue.signal`

Current `manual_wait.json` payload:
- `tag`
- `contexto`
- `empresa_pasta`
- `timestamp`

Current backend behavior:
- before waiting, backend removes stale `continue.signal` if present.
- backend writes `manual_wait.json`.
- backend emits console line `[MANUAL] AGUARDANDO_LIBERACAO_APP | <contexto>` or `[MANUAL-FINAL] ...`.
- backend waits indefinitely until `continue.signal` appears.
- when `continue.signal` is detected, backend removes it, emits `CONTINUANDO_EXECUCAO_MANUAL`, and removes `manual_wait.json`.

Current login/captcha handshake:
- backend fills CNPJ and password.
- backend emits informational lines tagged as `MANUAL-LOGIN`.
- operator resolves the captcha and clicks `Entrar`.
- backend starts or continues the timed wait for dashboard confirmation automatically after the portal submit.
- GUI must not require `continue.signal` for `MANUAL-LOGIN`.

Current GUI behavior:
- detects formal manual wait from control files and parsed log events.
- writes `<control_dir>/<empresa_pasta>/continue.signal` to release the blocked step.

Current GUI rule for `MANUAL-LOGIN`:
- display it as progress information only.
- do not block the operator with a mandatory continue action.
- expect automatic transition to the next phase after successful portal submit.

Current limitations:
- no backend timeout for manual wait.
- no backend retry counter for manual wait.
- cancel is implemented by stopping the orchestrator process, not by a dedicated manual-wait API.

GUI obligations:
- show active manual-wait state clearly.
- expose a single `continue` action.
- keep `MANUAL-LOGIN` as informational UX, not as a release gate.
- expose stop/cancel through orchestrator stop, not by inventing another signal file.

GUI must not:
- edit `manual_wait.json`.
- rely on the textual timestamp written inside `continue.signal`.
- guess manual control paths without the company-folder rule below.

## Outputs

Base output layout:
- `<output_base_dir>/report_execucao_empresas.csv`
- `<output_base_dir>/checkpoint_execucao_empresas.json`
- `<output_base_dir>/downloads/`
- `<output_base_dir>/_control/`

Legacy fallback locations still exist for artifact lookup:
- `<project_dir>/report_execucao_empresas.csv`
- `<project_dir>/checkpoint_execucao_empresas.json`
- `<project_dir>/downloads/`

Compatibility rule for legacy paths:
- these fallback paths exist for backward compatibility with old runs and migrated artifacts.
- they are `compat-only`, not primary output targets.
- the new GUI must write to and read from `output_base_dir` paths as the primary flow.
- the new GUI may consult legacy paths only when opening historical artifacts through the artifact locator.
- the new GUI must not treat legacy paths as the default runtime destination.

### Report Contract

File:
- `report_execucao_empresas.csv`

Encoding:
- `utf-8-sig`

Delimiter:
- `;`

Current required header order:
1. `timestamp_inicio`
2. `timestamp_fim`
3. `codigo_empresa`
4. `razao_social`
5. `cnpj`
6. `segmento`
7. `status`
8. `motivo`
9. `tentativas`
10. `acao_recomendada`

Current semantics:
- one line per company result.
- report is append-only during execution after header initialization.
- final status must be read from this file, not inferred only from live events.

Current GUI consumption from the report:
- `status -> resultado`
- `motivo -> ultima_mensagem`
- `tentativas`
- `timestamp_inicio -> inicio`
- `timestamp_fim -> fim`
- `acao_recomendada`

### Checkpoint Contract

File:
- `checkpoint_execucao_empresas.json`

Current payload fields:
- `timestamp`
- `ultimo_indice`
- `processadas`

Current semantics:
- resumability aid for the orchestrator.
- not a GUI display contract.

The GUI may expose it as diagnostic information, but must not depend on it for result rendering.

### Company Artifact Contract

Current company folder naming rule:
- preferred: `nome_pasta_empresa_por_dados(empresa)`
- fallback legacy: `nome_pasta_empresa_legada_por_dados(empresa)`

The GUI must resolve artifacts through `ArtifactLocatorService`.

The GUI must not hardcode company folder names.

The GUI must prefer artifact resolution under `output_base_dir`.

Legacy artifact resolution under `project_dir` is `compat-only`.

Current company artifact paths:
- company root: `<downloads_dir>/<empresa_pasta>/`
- NFSe log: `log_downloads_nfse.txt`
- Tomados log: `log_tomados.txt`
- manual log: `log_fechamento_manual.txt`
- debug directory: `_debug/`
- evidence directory: `<competencia>/_evidencias/`

Common output artifacts currently produced by the backend:
- downloaded NFSe XML files
- `SERVICOS_TOMADOS_MM-AAAA.xml`
- `LIVRO_SERVICOS_PRESTADOS_MM-AAAA.pdf`
- `LIVRO_SERVICOS_TOMADOS_MM-AAAA.pdf`
- `GUIA_ISS_PRESTADOS_MM-AAAA.pdf`
- `GUIA_ISS_TOMADOS_MM-AAAA.pdf`
- `_debug/<timestamp>/...`
- `_evidencias/*.png`

## Exit-Code Mapping

This mapping is internal between `main.py` and the orchestrator, but is useful for contract validation during `MAN-001`.

Current `main.py` exit codes:
- `0 -> SUCESSO`
- `30 -> CAPTCHA timeout`
- `40 -> SUCESSO_SEM_COMPETENCIA`
- `41 -> SUCESSO_SEM_SERVICOS`
- `42 -> falha em Tomados`
- `50 -> REVISAO_MANUAL`
- `60 -> falha de inicializacao do Chrome/WebDriver`

Current orchestrator mapping:
- `0 -> SUCESSO`
- `40 -> SUCESSO_SEM_COMPETENCIA`
- `41 -> SUCESSO_SEM_SERVICOS`
- `50 -> REVISAO_MANUAL`
- `30|42|60|other -> FALHA`

The GUI should consume orchestrator results, not raw `main.py` exit codes.

## Responsibilities

Backend responsibilities:
- validate runtime inputs.
- execute companies in order.
- write `report_execucao_empresas.csv`.
- manage `checkpoint_execucao_empresas.json`.
- emit live console events.
- create and consume manual control files.
- produce downloadable and diagnostic artifacts.
- decide final status and recommended action.

GUI responsibilities:
- collect valid runtime config.
- launch and stop the orchestrator.
- render live progress conservatively.
- render final result from the report.
- surface manual wait and allow continue.
- open artifacts using the artifact locator.
- keep user-facing state aligned with the report and control files.

The GUI must never infer on its own:
- final status from free-form logs only.
- company folder names by string concatenation.
- report schema changes.
- manual-wait completion without `manual_resumed` or the release action.

## Stability Rules

Stable in `v1`:
- report file name
- report header order
- final result values
- recommended action values
- manual control file names
- base output directory layout
- artifact lookup via `ArtifactLocatorService`

Still non-contract in `v1`:
- free-form console text beyond the event heuristics above
- exact wording inside `motivo`
- exact wording of non-contract log lines
- internal Selenium and site navigation details

## MAN-001 Outcome

Validated in `MAN-001`:
- a real run reached `SUCESSO`
- final rendering through `report_execucao_empresas.csv` is sufficient
- artifact lookup is sufficient for logs, evidence and company folders
- manual wait is detectable and releasable without ambiguity for `MANUAL-FINAL`

Accepted caveat from `MAN-001`:
- `MANUAL-LOGIN` is now intentionally informational and auto-resuming
- persisted evidence for login/captcha remains useful operationally, but it is no longer part of the formal manual control handshake
