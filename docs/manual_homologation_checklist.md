# Manual Homologation Checklist

Status: `DRAFT`

Related contract:
- [backend_gui_contract.md](/c:/Users/Windows%2011/Desktop/NFS%20XML/app%20come%C3%A7o/automacao-nfse/docs/backend_gui_contract.md)

Purpose: execute `MAN-001` against the real portal and use the result to validate or adjust the draft backend/GUI contract.

## Goal

Accept the backend for GUI integration only after a real execution outside the purely local automated harness.

Minimum acceptance for `MAN-001`:
- at least one real company finishes as `SUCESSO` or `REVISAO_MANUAL`
- report is generated with the contracted schema
- manual-wait flow is understandable and releasable
- artifact locations are sufficient for GUI consumption

## Preconditions

Required:
- packaged release available
- smoke official passing
- packaged validation passing
- valid company input file
- valid portal credentials for at least one company
- writable output directory

Recommended:
- one company expected to complete successfully
- one company likely to exercise manual wait or artifact generation in Tomados/Prestados

## Test Setup

Record before execution:
- release path used
- input file used
- target `apuracao_referencia`
- output directory used
- operator name
- date/time

Preferred execution path:
1. Run through the current desktop app if the objective includes validating GUI consumption of manual wait.
2. If the GUI is not the target for this run, run the packaged orchestrator and inspect the control files and outputs directly.

## Execution Checklist

### 1. Start

Confirm:
- input file loads without validation error
- `apuracao_referencia` is correct
- output directory is empty enough to isolate the run

Capture evidence:
- screenshot of launch configuration

### 2. Runtime Progress

During execution, confirm:
- active company is identifiable
- progress messages are understandable enough to map the current phase
- the run can be distinguished between login, captcha wait, Prestados, Tomados and manual/livros stages
- after captcha submit, the automation resumes automatically without requiring a GUI continue action at login

Record any mismatch with the draft contract:
- missing stage
- ambiguous stage
- stage text too unstable for UI display

Capture evidence:
- screenshot of live progress
- excerpt of stdout/stderr if needed

### 3. Manual Wait

If manual wait occurs, confirm:
- whether it occurred at login/captcha or later manual/livros stages
- the waiting company is identifiable
- the waiting context is understandable
- a control directory exists under `<output_base_dir>/_control/run_<...>/`
- `manual_wait.json` exists for the company
- resume can be triggered cleanly
- after resume, `manual_wait.json` is removed

Current contract nuance:
- `MANUAL-LOGIN` is no longer a formal manual wait.
- login/captcha may still appear in live progress and logs, but it should auto-resume after clicking `Entrar`.
- `manual_wait.json` and `continue.signal` are expected only for formal manual stages such as `MANUAL` and `MANUAL-FINAL`.

Inspect and record:
- `tag`
- `contexto`
- `empresa_pasta`
- whether the context was sufficient for a GUI button label or manual status text

If manual wait does not occur in this run:
- mark this section as `not exercised`
- note whether another company/test month is needed

Capture evidence:
- screenshot of manual wait UI or filesystem state
- copy of `manual_wait.json`
- screenshot or note after resume

### 4. Final Report

Confirm `report_execucao_empresas.csv` exists and validate:
- encoding loads correctly
- header matches the contract exactly
- one line exists per processed company
- `status` is sufficient for final rendering
- `motivo` is understandable enough for user display
- `tentativas` is populated
- `acao_recomendada` is populated when expected

Capture evidence:
- report file path
- first relevant rows

### 5. Artifacts

For at least one processed company, confirm the backend generated locatable artifacts:
- company folder
- `log_downloads_nfse.txt`
- `log_tomados.txt`
- `log_fechamento_manual.txt` if manual flow occurred
- `_debug/` if a debug case occurred
- `_evidencias/` if evidence screenshots were produced

Confirm artifact lookup is possible without hardcoding the company folder.

Capture evidence:
- screenshot of folder tree
- paths opened successfully

### 6. Final Outcome

Confirm the run reaches one of:
- `SUCESSO`
- `REVISAO_MANUAL`

If only `FALHA` occurs:
- do not freeze the contract
- register the failure details and decide whether the issue is backend, environment, or portal data

Capture evidence:
- final screen or console output
- relevant report row

## Contract Review Questions

Use these questions while reviewing the run:
- Are current runtime `etapa` values sufficient for the GUI?
- Is any new final status needed?
- Is `motivo` enough, or is another structured field needed?
- Is `manual_wait.json.contexto` sufficient for the operator?
- Is any required artifact missing from the current lookup service?
- Is any path rule too implicit for safe GUI integration?

## Decision Matrix

Freeze contract after `MAN-001` only if:
- runtime progress is usable
- manual wait is clear or explicitly marked as not exercised
- final report schema is sufficient
- artifact lookup is sufficient
- no new backend field is required for GUI correctness

Update the contract and rerun `MAN-001` if:
- the GUI would need to infer hidden state
- a final state is ambiguous
- manual wait is not actionable enough
- report fields are insufficient
- artifact paths are not reliably discoverable

## Evidence Bundle

Keep together:
- release path
- input file used
- output directory
- `report_execucao_empresas.csv`
- `checkpoint_execucao_empresas.json` if retained
- relevant logs
- `manual_wait.json` if exercised
- screenshots of progress, manual wait and final result

## Result Template

Fill this at the end of the run:

```text
MAN-001 result:
- Date:
- Operator:
- Release:
- Input:
- Apuracao:
- Output dir:
- Companies exercised:
- Final statuses observed:
- Manual wait exercised: yes/no
- Report schema valid: yes/no
- Artifact lookup sufficient: yes/no
- Contract changes required:
- Decision: freeze draft / update draft and rerun
```
