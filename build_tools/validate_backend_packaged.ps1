param(
    [string]$ReleaseDir = "dist\\AutomacaoNFSe-Orquestrador",
    [string]$OutputDir = "build_tools\\homologacao_output",
    [int]$CaseTimeoutSeconds = 30
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$releasePath = Join-Path $projectDir $ReleaseDir
$outputRoot = Join-Path $projectDir $OutputDir
$resultsCsvPath = Join-Path $outputRoot "backend_homologation_results.csv"
$resultsMdPath = Join-Path $outputRoot "backend_homologation_report.md"
$runStamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
$tempRoot = Join-Path $env:TEMP "Homologacao NFSe\\Teste 01\\Run_$runStamp"
$copiedReleasePath = Join-Path $tempRoot "AutomacaoNFSe-Orquestrador"
$buildWarnPath = Join-Path $projectDir "build\\AutomacaoNFSe-Orquestrador\\warn-AutomacaoNFSe-Orquestrador.txt"
$mainWarnPath = Join-Path $projectDir "build\\AutomacaoNFSe-Main\\warn-AutomacaoNFSe-Main.txt"
$buildDelegatorScript = Join-Path $scriptDir "build_orchestrator.ps1"
$smokeScript = Join-Path $scriptDir "test_orchestrator_smoke.ps1"
$venvPython = Join-Path $projectDir ".venv\\Scripts\\python.exe"
$orchestratorExeName = "AutomacaoNFSe-Orquestrador.exe"
$mainBundleName = "AutomacaoNFSe-Main"
$mainExeName = "AutomacaoNFSe-Main.exe"
$reportHeader = @(
    "timestamp_inicio",
    "timestamp_fim",
    "codigo_empresa",
    "razao_social",
    "cnpj",
    "segmento",
    "status",
    "motivo",
    "tentativas",
    "acao_recomendada"
)

New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

$results = New-Object System.Collections.Generic.List[object]

function Stop-ProcessTree {
    param([int]$ProcessId)

    if ($ProcessId -le 0) {
        return
    }

    & taskkill /PID $ProcessId /T /F *> $null
}

function Get-AutomationBrowserProcesses {
    try {
        return @(
            Get-CimInstance Win32_Process -ErrorAction Stop |
                Where-Object {
                    $_.Name -in @("chrome.exe", "chromedriver.exe") -and (
                        $_.Name -eq "chromedriver.exe" -or
                        [string]$_.CommandLine -match "--enable-automation" -or
                        [string]$_.CommandLine -match "--remote-debugging-port=" -or
                        [string]$_.CommandLine -match "--test-type=webdriver" -or
                        [string]$_.CommandLine -match "--user-data-dir="
                    )
                } |
                ForEach-Object {
                    [pscustomobject]@{
                        Id = [int]$_.ProcessId
                        Name = [string]$_.Name
                    }
                }
        )
    }
    catch {
        return @(
            Get-Process -ErrorAction SilentlyContinue |
                Where-Object { $_.ProcessName -in @("chrome", "chromedriver") } |
                ForEach-Object {
                    [pscustomobject]@{
                        Id = [int]$_.Id
                        Name = "$($_.ProcessName).exe"
                    }
                }
        )
    }
}

function Stop-AutomationBrowserProcesses {
    $browserProcesses = @(Get-AutomationBrowserProcesses)
    if ($browserProcesses.Count -eq 0) {
        return
    }

    foreach ($proc in @($browserProcesses | Where-Object { $_.Name -ieq "chromedriver.exe" })) {
        Stop-ProcessTree -ProcessId $proc.Id
    }
    foreach ($proc in @($browserProcesses | Where-Object { $_.Name -ieq "chrome.exe" })) {
        Stop-ProcessTree -ProcessId $proc.Id
    }

    Start-Sleep -Milliseconds 300
}

function New-TestResult {
    param(
        [string]$Id,
        [string]$Objective,
        [string]$Input,
        [string]$Steps,
        [string]$Expected,
        [string]$Actual,
        [string]$Status,
        [string]$Evidence
    )

    $results.Add([pscustomobject]@{
        id = $Id
        objective = $Objective
        input = $Input
        steps = $Steps
        expected = $Expected
        actual = $Actual
        status = $Status
        evidence = $Evidence
    }) | Out-Null
}

function Reset-PathIfExists {
    param([string]$PathToReset)

    if (Test-Path $PathToReset) {
        Remove-Item -Recurse -Force $PathToReset
    }
}

function New-CaseDir {
    param([string]$CaseId)

    $caseDir = Join-Path $outputRoot $CaseId
    Reset-PathIfExists $caseDir
    New-Item -ItemType Directory -Force -Path $caseDir | Out-Null
    return $caseDir
}

function Copy-ReleaseToTemp {
    New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
    Copy-Item -Path $releasePath -Destination $copiedReleasePath -Recurse -Force
}

function Invoke-Executable {
    param(
        [string]$ExePath,
        [string]$WorkingDirectory,
        [hashtable]$EnvironmentOverrides = @{},
        [string[]]$ArgumentList = @(),
        [string]$StdOutPath,
        [string]$StdErrPath,
        [int]$TimeoutSeconds = $CaseTimeoutSeconds
    )

    $previousEnv = @{}
    foreach ($entry in $EnvironmentOverrides.GetEnumerator()) {
        $previousEnv[$entry.Key] = [Environment]::GetEnvironmentVariable($entry.Key, "Process")
        [Environment]::SetEnvironmentVariable($entry.Key, [string]$entry.Value, "Process")
    }

    $proc = $null
    $timedOut = $false
    try {
        try { Stop-AutomationBrowserProcesses } catch {}

        if (Test-Path $StdOutPath) { Remove-Item -Force $StdOutPath }
        if (Test-Path $StdErrPath) { Remove-Item -Force $StdErrPath }

        $startParams = @{
            FilePath = $ExePath
            WorkingDirectory = $WorkingDirectory
            RedirectStandardOutput = $StdOutPath
            RedirectStandardError = $StdErrPath
            PassThru = $true
        }
        if ($ArgumentList.Count -gt 0) {
            $startParams.ArgumentList = $ArgumentList
        }

        $proc = Start-Process @startParams

        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            $timedOut = $true
            Stop-ProcessTree -ProcessId $proc.Id
            $proc.WaitForExit(5000) | Out-Null
        }
    }
    finally {
        try {
            if ($proc -and -not $proc.HasExited) {
                Stop-ProcessTree -ProcessId $proc.Id
                $proc.WaitForExit(5000) | Out-Null
            }
        } catch {}

        foreach ($entry in $previousEnv.GetEnumerator()) {
            [Environment]::SetEnvironmentVariable($entry.Key, $entry.Value, "Process")
        }
    }

    $stdout = if (Test-Path $StdOutPath) { Get-Content -Path $StdOutPath -Raw -ErrorAction SilentlyContinue } else { "" }
    $stderr = if (Test-Path $StdErrPath) { Get-Content -Path $StdErrPath -Raw -ErrorAction SilentlyContinue } else { "" }

    try { Stop-AutomationBrowserProcesses } catch {}

    return @{
        ExitCode = if ($proc) { [int]$proc.ExitCode } else { -1 }
        StdOut = $stdout
        StdErr = $stderr
        TimedOut = $timedOut
    }
}

function New-CsvFixture {
    param(
        [string]$Path,
        [string[]]$Headers,
        [object[]]$Rows,
        [string]$Delimiter = ";",
        [string]$Encoding = "utf8"
    )

    $normalizedHeaders = foreach ($header in $Headers) {
        if ($header -match "^C.*digo$") {
            "Codigo"
            continue
        }
        if ($header -match "^Raz.* Social$") {
            "Razao Social"
            continue
        }
        $header
    }

    $rowGroups = @()
    $rawRows = @($Rows)

    if ($rawRows.Count -gt 0) {
        $firstItem = $rawRows[0]
        $isFlatRowPayload = ($firstItem -is [string]) -or ($firstItem -isnot [System.Collections.IEnumerable])

        if ($isFlatRowPayload) {
            if (($rawRows.Count % $normalizedHeaders.Count) -ne 0) {
                throw "Fixture invalida para ${Path}: quantidade de valores nao bate com o cabecalho."
            }

            for ($i = 0; $i -lt $rawRows.Count; $i += $normalizedHeaders.Count) {
                $rowGroups += ,@($rawRows[$i..($i + $normalizedHeaders.Count - 1)])
            }
        } else {
            foreach ($row in $rawRows) {
                $rowGroups += ,@($row)
            }
        }
    }

    $lines = @()
    $lines += ($normalizedHeaders -join $Delimiter)
    foreach ($row in $rowGroups) {
        $values = @()
        foreach ($value in $row) {
            $values += [string]$value
        }
        if ($values.Count -ne $normalizedHeaders.Count) {
            throw "Fixture invalida para ${Path}: linha com $($values.Count) colunas; esperado $($normalizedHeaders.Count)."
        }
        $lines += ($values -join $Delimiter)
    }

    Set-Content -Path $Path -Value $lines -Encoding $Encoding
}

function Read-ReportRows {
    param([string]$ReportPath)

    if (-not (Test-Path $ReportPath)) {
        return ,@()
    }

    return ,@(Import-Csv -Path $ReportPath -Delimiter ';')
}

function Join-PathDisplay {
    param([string[]]$Items)

    return (($Items | Where-Object { $_ }) -join " | ")
}

function Test-ReportHeader {
    param([string]$ReportPath)

    if (-not (Test-Path $ReportPath)) {
        return $false
    }

    $firstLine = Get-Content -Path $ReportPath -TotalCount 1
    if (-not $firstLine) {
        return $false
    }

    $currentHeader = $firstLine -split ';'
    if ($currentHeader[0] -match "^\uFEFF") {
        $currentHeader[0] = $currentHeader[0].TrimStart([char]0xFEFF)
    }

    return (@($currentHeader) -join ';') -eq ($reportHeader -join ';')
}

try {
    if (-not (Test-Path $releasePath)) {
        throw "Release do backend nao encontrado em: $releasePath"
    }

    $caseDir = New-CaseDir "INT-001"
    $orchestratorExePath = Join-Path $releasePath $orchestratorExeName
    $mainExePath = Join-Path $releasePath "$mainBundleName\\$mainExeName"
    $sourceMainPath = Join-Path $releasePath "main.py"
    $missingReleaseItems = @()
    if (-not (Test-Path $orchestratorExePath)) { $missingReleaseItems += $orchestratorExeName }
    if (-not (Test-Path $mainExePath)) { $missingReleaseItems += "$mainBundleName\\$mainExeName" }
    $releaseStatus = if ($missingReleaseItems.Count -eq 0 -and -not (Test-Path $sourceMainPath)) { "APROVADO" } else { "REPROVADO" }
    $releaseActual = if ($missingReleaseItems.Count -eq 0) {
        "Executaveis presentes e release nao contem main.py solto."
    } else {
        "Itens ausentes: $($missingReleaseItems -join ', ')"
    }
    if (Test-Path $sourceMainPath) {
        $releaseActual += " main.py foi encontrado no release."
    }
    New-TestResult `
        -Id "INT-001" `
        -Objective "Confirmar integridade minima do release empacotado." `
        -Input $releasePath `
        -Steps "Verificar executavel do Orquestrador, executavel do Main e ausencia de main.py solto." `
        -Expected "Release contem os dois executaveis e nao depende de fonte Python solta." `
        -Actual $releaseActual `
        -Status $releaseStatus `
        -Evidence (Join-PathDisplay @($orchestratorExePath, $mainExePath))

    $caseDir = New-CaseDir "INT-002"
    if (Test-Path $buildWarnPath) {
        $warnContent = Get-Content -Path $buildWarnPath
        $mainWarnContent = if (Test-Path $mainWarnPath) { Get-Content -Path $mainWarnPath } else { @() }
        $relevantMissing = @()
        if ($warnContent -match "missing module named openpyxl($|')") { $relevantMissing += "openpyxl" }
        if ($warnContent -match "missing module named selenium($|')") { $relevantMissing += "selenium" }
        if ($warnContent -match "missing module named core\.app_info") { $relevantMissing += "core.app_info" }
        if ($warnContent -match "missing module named core\.company_paths") { $relevantMissing += "core.company_paths" }
        if ($mainWarnContent -match "missing module named selenium\.webdriver\.chrome\.webdriver") { $relevantMissing += "selenium.webdriver.chrome.webdriver" }
        $warnStatus = if ($relevantMissing.Count -eq 0) { "APROVADO" } else { "REPROVADO" }
        $warnActual = if ($relevantMissing.Count -eq 0) {
            "Nao ha missing module relevante no warn do build."
        } else {
            "Warn acusa dependencias relevantes: $($relevantMissing -join ', ')"
        }
    } else {
        $warnStatus = "PENDENTE_MANUAL"
        $warnActual = "Arquivo de warn do PyInstaller nao foi encontrado."
    }
    New-TestResult `
        -Id "INT-002" `
        -Objective "Validar que o build log nao acusa dependencias relevantes ausentes." `
        -Input $buildWarnPath `
        -Steps "Ler warn-AutomacaoNFSe-Orquestrador.txt e procurar modulos obrigatorios." `
        -Expected "Sem openpyxl, selenium ou modulos internos marcados como missing." `
        -Actual $warnActual `
        -Status $warnStatus `
        -Evidence $buildWarnPath

    $caseDir = New-CaseDir "DEP-001"
    if (-not (Test-Path $venvPython)) {
        New-TestResult `
            -Id "DEP-001" `
            -Objective "Confirmar build a partir da .venv correta." `
            -Input $venvPython `
            -Steps "Verificar existencia da .venv e pacotes obrigatorios." `
            -Expected ".venv presente com pyinstaller, openpyxl e selenium." `
            -Actual "Python da .venv nao encontrado." `
            -Status "REPROVADO" `
            -Evidence $venvPython
    } else {
        $missingPackages = @()
        foreach ($packageName in @("pyinstaller", "openpyxl", "selenium")) {
            & $venvPython -m pip show $packageName | Out-Null
            if ($LASTEXITCODE -ne 0) {
                $missingPackages += $packageName
            }
        }
        $depStatus = if ([string]::IsNullOrWhiteSpace($missingPackages)) { "APROVADO" } else { "REPROVADO" }
        $depActual = if ($depStatus -eq "APROVADO") {
            ".venv contem pyinstaller, openpyxl e selenium."
        } else {
            "Pacotes ausentes na .venv: $($missingPackages -join ', ')"
        }
        New-TestResult `
            -Id "DEP-001" `
            -Objective "Confirmar build a partir da .venv correta." `
            -Input $venvPython `
            -Steps "Importar pyinstaller, openpyxl e selenium usando o Python da .venv." `
            -Expected ".venv presente com dependencias obrigatorias." `
            -Actual $depActual `
            -Status $depStatus `
            -Evidence $venvPython
    }

    Copy-ReleaseToTemp

    $caseDir = New-CaseDir "RUN-001"
    $fixturesDir = Join-Path $caseDir "fixtures"
    $executionDir = Join-Path $caseDir "saida"
    New-Item -ItemType Directory -Force -Path $fixturesDir, $executionDir | Out-Null
    $validOneCsv = Join-Path $fixturesDir "valid_1.csv"
    New-CsvFixture -Path $validOneCsv -Headers @(
        "Código",
        "Razão Social",
        "CNPJ",
        "Segmento",
        "Senha Prefeitura"
    ) -Rows @(
        @("1001", "EMPRESA TESTE", "00.000.000/0001-00", "TESTE", "senha")
    )
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $reportPath = Join-Path $executionDir "report_execucao_empresas.csv"
    $checkpointPath = Join-Path $executionDir "checkpoint_execucao_empresas.json"
    $runOne = Invoke-Executable `
        -ExePath (Join-Path $copiedReleasePath $orchestratorExeName) `
        -WorkingDirectory $copiedReleasePath `
        -EnvironmentOverrides @{
            EMPRESAS_ARQUIVO = $validOneCsv
            OUTPUT_BASE_DIR = $executionDir
            CONTINUAR_DE_ONDE_PAROU = "0"
            USAR_CHECKPOINT = "1"
            MAX_TENTATIVAS_EMPRESA = "1"
            LOGIN_WAIT_SECONDS = "1"
            TIMEOUT_PROCESSO_MAIN = "1"
        } `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath
    $rowsRunOne = Read-ReportRows -ReportPath $reportPath
    $rowRunOne = if ($rowsRunOne.Count -ge 1) { $rowsRunOne[0] } else { $null }
    $runOneApproved = (
        $runOne.ExitCode -eq 0 -and
        (Test-Path $reportPath) -and
        (Test-ReportHeader -ReportPath $reportPath) -and
        $rowsRunOne.Count -eq 1 -and
        $rowRunOne.status -eq "FALHA" -and
        $rowRunOne.motivo -match "Timeout de execucao do backend principal" -and
        $rowRunOne.acao_recomendada -eq "ABRIR_DEBUG" -and
        (Test-Path $checkpointPath)
    )
    $runOneActual = "TimedOut=$($runOne.TimedOut); ExitCode=$($runOne.ExitCode); report=$([bool](Test-Path $reportPath)); rows=$($rowsRunOne.Count); status=$($rowRunOne.status); acao=$($rowRunOne.acao_recomendada); checkpoint=$([bool](Test-Path $checkpointPath))"
    New-TestResult `
        -Id "RUN-001" `
        -Objective "Executar release copiado para fora do projeto com path absoluto e falha controlada." `
        -Input $validOneCsv `
        -Steps "Copiar release para pasta com espaco/acento e rodar com TIMEOUT_PROCESSO_MAIN=1." `
        -Expected "Release funciona fora do projeto, gera report/checkpoint e registra timeout de forma rastreavel." `
        -Actual $runOneActual `
        -Status $(if ($runOneApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence (Join-PathDisplay @($copiedReleasePath, $reportPath, $stdoutPath, $stderrPath))

    $caseDir = New-CaseDir "RUN-002"
    $relativeFixturesDir = Join-Path $copiedReleasePath "fixtures_rel"
    $relativeOutputDir = Join-Path $copiedReleasePath "Saida Relativa"
    Reset-PathIfExists $relativeFixturesDir
    Reset-PathIfExists $relativeOutputDir
    New-Item -ItemType Directory -Force -Path $relativeFixturesDir | Out-Null
    $relativeCsv = Join-Path $relativeFixturesDir "valid_1_rel.csv"
    New-CsvFixture -Path $relativeCsv -Headers @(
        "Código",
        "Razão Social",
        "CNPJ",
        "Segmento",
        "Senha Prefeitura"
    ) -Rows @(
        @("2001", "EMPRESA RELATIVA", "11.111.111/0001-11", "TESTE", "senha")
    )
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $relativeReportPath = Join-Path $relativeOutputDir "report_execucao_empresas.csv"
    $runRelative = Invoke-Executable `
        -ExePath (Join-Path $copiedReleasePath $orchestratorExeName) `
        -WorkingDirectory $copiedReleasePath `
        -EnvironmentOverrides @{
            EMPRESAS_ARQUIVO = ".\\fixtures_rel\\valid_1_rel.csv"
            OUTPUT_BASE_DIR = ".\\Saida Relativa"
            CONTINUAR_DE_ONDE_PAROU = "0"
            USAR_CHECKPOINT = "0"
            MAX_TENTATIVAS_EMPRESA = "1"
            LOGIN_WAIT_SECONDS = "1"
            TIMEOUT_PROCESSO_MAIN = "1"
        } `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath
    $rowsRelative = Read-ReportRows -ReportPath $relativeReportPath
    $rowRelative = if ($rowsRelative.Count -ge 1) { $rowsRelative[0] } else { $null }
    $relativeApproved = (
        $runRelative.ExitCode -eq 0 -and
        (Test-Path $relativeReportPath) -and
        $rowsRelative.Count -eq 1 -and
        $rowRelative.status -eq "FALHA"
    )
    $relativeActual = "TimedOut=$($runRelative.TimedOut); ExitCode=$($runRelative.ExitCode); report=$([bool](Test-Path $relativeReportPath)); rows=$($rowsRelative.Count); status=$($rowRelative.status)"
    New-TestResult `
        -Id "RUN-002" `
        -Objective "Validar leitura e escrita por path relativo em modo empacotado." `
        -Input ".\\fixtures_rel\\valid_1_rel.csv" `
        -Steps "Rodar release copiado com EMPRESAS_ARQUIVO e OUTPUT_BASE_DIR relativos ao cwd empacotado." `
        -Expected "Input relativo e output relativo funcionam sem depender da arvore do dev." `
        -Actual $relativeActual `
        -Status $(if ($relativeApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence (Join-PathDisplay @($relativeReportPath, $stdoutPath, $stderrPath))

    $caseDir = New-CaseDir "RUN-003"
    $fixturesDir = Join-Path $caseDir "fixtures"
    $executionDir = Join-Path $caseDir "saida"
    New-Item -ItemType Directory -Force -Path $fixturesDir, $executionDir | Out-Null
    $validThreeCsv = Join-Path $fixturesDir "valid_3.csv"
    New-CsvFixture -Path $validThreeCsv -Headers @(
        "Código",
        "Razão Social",
        "CNPJ",
        "Segmento",
        "Senha Prefeitura"
    ) -Rows @(
        @("3001", "EMPRESA A", "22.222.222/0001-22", "TESTE", "senha"),
        @("3002", "EMPRESA B", "33.333.333/0001-33", "TESTE", "senha"),
        @("3003", "EMPRESA C", "44.444.444/0001-44", "TESTE", "senha")
    )
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $reportPath = Join-Path $executionDir "report_execucao_empresas.csv"
    $runThree = Invoke-Executable `
        -ExePath (Join-Path $copiedReleasePath $orchestratorExeName) `
        -WorkingDirectory $copiedReleasePath `
        -EnvironmentOverrides @{
            EMPRESAS_ARQUIVO = $validThreeCsv
            OUTPUT_BASE_DIR = $executionDir
            CONTINUAR_DE_ONDE_PAROU = "0"
            USAR_CHECKPOINT = "1"
            MAX_TENTATIVAS_EMPRESA = "1"
            LOGIN_WAIT_SECONDS = "1"
            TIMEOUT_PROCESSO_MAIN = "1"
        } `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath
    $rowsRunThree = Read-ReportRows -ReportPath $reportPath
    $runThreeApproved = (
        $runThree.ExitCode -eq 0 -and
        (Test-Path $reportPath) -and
        $rowsRunThree.Count -eq 3 -and
        (@($rowsRunThree | Where-Object { $_.status -eq "FALHA" }).Count -eq 3)
    )
    $runThreeActual = "TimedOut=$($runThree.TimedOut); ExitCode=$($runThree.ExitCode); report=$([bool](Test-Path $reportPath)); rows=$($rowsRunThree.Count); falhas=$(@($rowsRunThree | Where-Object { $_.status -eq 'FALHA' }).Count)"
    New-TestResult `
        -Id "RUN-003" `
        -Objective "Validar lote pequeno e contrato de uma linha por empresa." `
        -Input $validThreeCsv `
        -Steps "Executar release com 3 empresas e timeout controlado." `
        -Expected "Report com 3 linhas de resultado, uma por empresa." `
        -Actual $runThreeActual `
        -Status $(if ($runThreeApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence (Join-PathDisplay @($reportPath, $stdoutPath, $stderrPath))

    $caseDir = New-CaseDir "FAIL-001"
    $mutatedRelease = Join-Path $caseDir "release_sem_main_exe"
    Copy-Item -Path $copiedReleasePath -Destination $mutatedRelease -Recurse -Force
    Remove-Item -Force (Join-Path $mutatedRelease "$mainBundleName\\$mainExeName")
    $fixturesDir = Join-Path $caseDir "fixtures"
    $executionDir = Join-Path $caseDir "saida"
    New-Item -ItemType Directory -Force -Path $fixturesDir, $executionDir | Out-Null
    $validCsv = Join-Path $fixturesDir "valid.csv"
    New-CsvFixture -Path $validCsv -Headers @(
        "Código",
        "Razão Social",
        "CNPJ",
        "Segmento",
        "Senha Prefeitura"
    ) -Rows @(
        @("4001", "EMPRESA SEM MAIN", "55.555.555/0001-55", "TESTE", "senha")
    )
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $reportPath = Join-Path $executionDir "report_execucao_empresas.csv"
    $failMissingExe = Invoke-Executable `
        -ExePath (Join-Path $mutatedRelease $orchestratorExeName) `
        -WorkingDirectory $mutatedRelease `
        -EnvironmentOverrides @{
            EMPRESAS_ARQUIVO = $validCsv
            OUTPUT_BASE_DIR = $executionDir
            CONTINUAR_DE_ONDE_PAROU = "0"
            USAR_CHECKPOINT = "0"
            MAX_TENTATIVAS_EMPRESA = "1"
            LOGIN_WAIT_SECONDS = "1"
            TIMEOUT_PROCESSO_MAIN = "1"
        } `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath
    $rowsMissingExe = Read-ReportRows -ReportPath $reportPath
    $rowMissingExe = if ($rowsMissingExe.Count -ge 1) { $rowsMissingExe[0] } else { $null }
    $missingExeApproved = (
        $failMissingExe.ExitCode -eq 0 -and
        (Test-Path $reportPath) -and
        $rowsMissingExe.Count -eq 1 -and
        $rowMissingExe.motivo -match "Executavel do backend principal nao encontrado"
    )
    $missingExeActual = "TimedOut=$($failMissingExe.TimedOut); ExitCode=$($failMissingExe.ExitCode); rows=$($rowsMissingExe.Count); motivo=$($rowMissingExe.motivo)"
    New-TestResult `
        -Id "FAIL-001" `
        -Objective "Validar falha explicita quando o executavel do Main esta ausente." `
        -Input $validCsv `
        -Steps "Remover AutomacaoNFSe-Main.exe do release copiado e executar o orquestrador." `
        -Expected "Report com FALHA explicita informando que o executavel do Main nao foi encontrado." `
        -Actual $missingExeActual `
        -Status $(if ($missingExeApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence (Join-PathDisplay @($reportPath, $stdoutPath, $stderrPath))

    $caseDir = New-CaseDir "FAIL-002"
    $mutatedRelease = Join-Path $caseDir "release_sem_pasta_main"
    Copy-Item -Path $copiedReleasePath -Destination $mutatedRelease -Recurse -Force
    Remove-Item -Recurse -Force (Join-Path $mutatedRelease $mainBundleName)
    $fixturesDir = Join-Path $caseDir "fixtures"
    $executionDir = Join-Path $caseDir "saida"
    New-Item -ItemType Directory -Force -Path $fixturesDir, $executionDir | Out-Null
    $validCsv = Join-Path $fixturesDir "valid.csv"
    New-CsvFixture -Path $validCsv -Headers @(
        "Código",
        "Razão Social",
        "CNPJ",
        "Segmento",
        "Senha Prefeitura"
    ) -Rows @(
        @("5001", "EMPRESA SEM PASTA", "66.666.666/0001-66", "TESTE", "senha")
    )
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $reportPath = Join-Path $executionDir "report_execucao_empresas.csv"
    $failMissingDir = Invoke-Executable `
        -ExePath (Join-Path $mutatedRelease $orchestratorExeName) `
        -WorkingDirectory $mutatedRelease `
        -EnvironmentOverrides @{
            EMPRESAS_ARQUIVO = $validCsv
            OUTPUT_BASE_DIR = $executionDir
            CONTINUAR_DE_ONDE_PAROU = "0"
            USAR_CHECKPOINT = "0"
            MAX_TENTATIVAS_EMPRESA = "1"
            LOGIN_WAIT_SECONDS = "1"
            TIMEOUT_PROCESSO_MAIN = "1"
        } `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath
    $rowsMissingDir = Read-ReportRows -ReportPath $reportPath
    $rowMissingDir = if ($rowsMissingDir.Count -ge 1) { $rowsMissingDir[0] } else { $null }
    $missingDirApproved = (
        $failMissingDir.ExitCode -eq 0 -and
        (Test-Path $reportPath) -and
        $rowsMissingDir.Count -eq 1 -and
        $rowMissingDir.motivo -match "Executavel do backend principal nao encontrado"
    )
    $missingDirActual = "TimedOut=$($failMissingDir.TimedOut); ExitCode=$($failMissingDir.ExitCode); rows=$($rowsMissingDir.Count); motivo=$($rowMissingDir.motivo)"
    New-TestResult `
        -Id "FAIL-002" `
        -Objective "Validar falha explicita quando a pasta do Main esta ausente." `
        -Input $validCsv `
        -Steps "Remover a pasta AutomacaoNFSe-Main do release copiado e executar o orquestrador." `
        -Expected "Report com FALHA explicita informando ausencia do backend principal." `
        -Actual $missingDirActual `
        -Status $(if ($missingDirApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence (Join-PathDisplay @($reportPath, $stdoutPath, $stderrPath))

    $caseDir = New-CaseDir "INP-001"
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $missingInputRun = Invoke-Executable `
        -ExePath (Join-Path $copiedReleasePath $orchestratorExeName) `
        -WorkingDirectory $copiedReleasePath `
        -EnvironmentOverrides @{
            EMPRESAS_ARQUIVO = (Join-Path $caseDir "arquivo_inexistente.csv")
            OUTPUT_BASE_DIR = (Join-Path $caseDir "saida")
            CONTINUAR_DE_ONDE_PAROU = "0"
            USAR_CHECKPOINT = "0"
        } `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath
    $missingInputApproved = (
        $missingInputRun.ExitCode -eq 0 -and
        $missingInputRun.StdOut -match "Arquivo de empresas"
    )
    $missingInputActual = "TimedOut=$($missingInputRun.TimedOut); ExitCode=$($missingInputRun.ExitCode); stdout contem mensagem=$($missingInputRun.StdOut -match 'Arquivo de empresas')"
    New-TestResult `
        -Id "INP-001" `
        -Objective "Validar tratamento de arquivo de entrada inexistente." `
        -Input (Join-Path $caseDir "arquivo_inexistente.csv") `
        -Steps "Executar release com EMPRESAS_ARQUIVO apontando para arquivo inexistente." `
        -Expected "Mensagem clara de arquivo ausente, sem travamento." `
        -Actual $missingInputActual `
        -Status $(if ($missingInputApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence (Join-PathDisplay @($stdoutPath, $stderrPath))

    $caseDir = New-CaseDir "INP-002"
    $emptyCsv = Join-Path $caseDir "vazio.csv"
    Set-Content -Path $emptyCsv -Value @() -Encoding UTF8
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $emptyRun = Invoke-Executable `
        -ExePath (Join-Path $copiedReleasePath $orchestratorExeName) `
        -WorkingDirectory $copiedReleasePath `
        -EnvironmentOverrides @{
            EMPRESAS_ARQUIVO = $emptyCsv
            OUTPUT_BASE_DIR = (Join-Path $caseDir "saida")
            CONTINUAR_DE_ONDE_PAROU = "0"
            USAR_CHECKPOINT = "0"
        } `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath
    $emptyApproved = (
        $emptyRun.ExitCode -ne 0 -and
        (($emptyRun.StdErr + $emptyRun.StdOut) -match "Colunas obrigat")
    )
    $emptyActual = "TimedOut=$($emptyRun.TimedOut); ExitCode=$($emptyRun.ExitCode); erro explicito=$((($emptyRun.StdErr + $emptyRun.StdOut) -match 'Colunas obrigat'))"
    New-TestResult `
        -Id "INP-002" `
        -Objective "Validar tratamento de CSV vazio." `
        -Input $emptyCsv `
        -Steps "Executar release com arquivo CSV vazio." `
        -Expected "Falha explicita informando problema de colunas obrigatorias." `
        -Actual $emptyActual `
        -Status $(if ($emptyApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence (Join-PathDisplay @($stdoutPath, $stderrPath))

    $caseDir = New-CaseDir "INP-003"
    $headerOnlyCsv = Join-Path $caseDir "somente_header.csv"
    New-CsvFixture -Path $headerOnlyCsv -Headers @(
        "Código",
        "Razão Social",
        "CNPJ",
        "Segmento",
        "Senha Prefeitura"
    ) -Rows @()
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $headerOnlyOutput = Join-Path $caseDir "saida"
    $headerOnlyReport = Join-Path $headerOnlyOutput "report_execucao_empresas.csv"
    $headerOnlyRun = Invoke-Executable `
        -ExePath (Join-Path $copiedReleasePath $orchestratorExeName) `
        -WorkingDirectory $copiedReleasePath `
        -EnvironmentOverrides @{
            EMPRESAS_ARQUIVO = $headerOnlyCsv
            OUTPUT_BASE_DIR = $headerOnlyOutput
            CONTINUAR_DE_ONDE_PAROU = "0"
            USAR_CHECKPOINT = "0"
        } `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath
    $headerOnlyRows = Read-ReportRows -ReportPath $headerOnlyReport
    New-TestResult `
        -Id "INP-003" `
        -Objective "Validar comportamento para CSV com somente cabecalho." `
        -Input $headerOnlyCsv `
        -Steps "Executar release com CSV contendo apenas cabecalho." `
        -Expected "Erro claro e acionavel para lote sem empresas." `
        -Actual "TimedOut=$($headerOnlyRun.TimedOut); ExitCode=$($headerOnlyRun.ExitCode); empresas carregadas zero=$($headerOnlyRun.StdOut -match 'Empresas carregadas: 0'); linhas report=$($headerOnlyRows.Count)" `
        -Status "REPROVADO" `
        -Evidence (Join-PathDisplay @($headerOnlyReport, $stdoutPath, $stderrPath))

    $caseDir = New-CaseDir "INP-004"
    $missingColumnCsv = Join-Path $caseDir "sem_coluna.csv"
    New-CsvFixture -Path $missingColumnCsv -Headers @(
        "Código",
        "Razão Social",
        "CNPJ",
        "Segmento"
    ) -Rows @(
        @("7001", "EMPRESA SEM SENHA", "77.777.777/0001-77", "TESTE")
    )
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $missingColumnRun = Invoke-Executable `
        -ExePath (Join-Path $copiedReleasePath $orchestratorExeName) `
        -WorkingDirectory $copiedReleasePath `
        -EnvironmentOverrides @{
            EMPRESAS_ARQUIVO = $missingColumnCsv
            OUTPUT_BASE_DIR = (Join-Path $caseDir "saida")
            CONTINUAR_DE_ONDE_PAROU = "0"
            USAR_CHECKPOINT = "0"
        } `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath
    $missingColumnApproved = (
        $missingColumnRun.ExitCode -ne 0 -and
        (($missingColumnRun.StdErr + $missingColumnRun.StdOut) -match "Colunas obrigat")
    )
    $missingColumnActual = "TimedOut=$($missingColumnRun.TimedOut); ExitCode=$($missingColumnRun.ExitCode); erro explicito=$((($missingColumnRun.StdErr + $missingColumnRun.StdOut) -match 'Colunas obrigat'))"
    New-TestResult `
        -Id "INP-004" `
        -Objective "Validar tratamento de coluna obrigatoria ausente." `
        -Input $missingColumnCsv `
        -Steps "Executar release com CSV sem a coluna Senha Prefeitura." `
        -Expected "Falha explicita informando colunas obrigatorias ausentes." `
        -Actual $missingColumnActual `
        -Status $(if ($missingColumnApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence (Join-PathDisplay @($stdoutPath, $stderrPath))

    $caseDir = New-CaseDir "INP-005"
    $wrongDelimiterCsv = Join-Path $caseDir "delimitador_incorreto.csv"
    New-CsvFixture -Path $wrongDelimiterCsv -Headers @(
        "Código",
        "Razão Social",
        "CNPJ",
        "Segmento",
        "Senha Prefeitura"
    ) -Rows @(
        @("8001", "EMPRESA DELIM", "88.888.888/0001-88", "TESTE", "senha")
    ) -Delimiter ","
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $wrongDelimiterRun = Invoke-Executable `
        -ExePath (Join-Path $copiedReleasePath $orchestratorExeName) `
        -WorkingDirectory $copiedReleasePath `
        -EnvironmentOverrides @{
            EMPRESAS_ARQUIVO = $wrongDelimiterCsv
            OUTPUT_BASE_DIR = (Join-Path $caseDir "saida")
            CONTINUAR_DE_ONDE_PAROU = "0"
            USAR_CHECKPOINT = "0"
        } `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath
    $wrongDelimiterApproved = (
        $wrongDelimiterRun.ExitCode -ne 0 -and
        (($wrongDelimiterRun.StdErr + $wrongDelimiterRun.StdOut) -match "Colunas obrigat")
    )
    $wrongDelimiterActual = "TimedOut=$($wrongDelimiterRun.TimedOut); ExitCode=$($wrongDelimiterRun.ExitCode); erro explicito=$((($wrongDelimiterRun.StdErr + $wrongDelimiterRun.StdOut) -match 'Colunas obrigat'))"
    New-TestResult `
        -Id "INP-005" `
        -Objective "Validar tratamento de delimitador incorreto." `
        -Input $wrongDelimiterCsv `
        -Steps "Executar release com CSV separado por virgula." `
        -Expected "Falha explicita informando que as colunas obrigatorias nao foram encontradas." `
        -Actual $wrongDelimiterActual `
        -Status $(if ($wrongDelimiterApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence (Join-PathDisplay @($stdoutPath, $stderrPath))

    $reportFromRunThree = Join-Path $outputRoot "RUN-003\\saida\\report_execucao_empresas.csv"
    $repApproved = (Test-Path $reportFromRunThree) -and (Test-ReportHeader -ReportPath $reportFromRunThree)
    New-TestResult `
        -Id "REP-001" `
        -Objective "Validar estabilidade do cabecalho do report." `
        -Input $reportFromRunThree `
        -Steps "Comparar cabecalho do report gerado com o contrato esperado." `
        -Expected "Cabecalho exatamente igual ao contrato atual do backend." `
        -Actual $(if ($repApproved) { "Cabecalho do report corresponde ao contrato." } else { "Cabecalho divergente ou report ausente." }) `
        -Status $(if ($repApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence $reportFromRunThree

    $caseDir = New-CaseDir "BUILD-001"
    $smokeOutputDir = "build_tools\\homologacao_output\\smoke_run"
    $stdoutPath = Join-Path $caseDir "stdout.txt"
    $stderrPath = Join-Path $caseDir "stderr.txt"
    $smokeRun = Invoke-Executable `
        -ExePath "powershell.exe" `
        -ArgumentList @(
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            $smokeScript,
            "-OutputDir",
            $smokeOutputDir,
            "-TimeoutSeconds",
            [string]$CaseTimeoutSeconds
        ) `
        -WorkingDirectory $projectDir `
        -EnvironmentOverrides @{} `
        -StdOutPath $stdoutPath `
        -StdErrPath $stderrPath `
        -TimeoutSeconds ($CaseTimeoutSeconds + 10)
    $smokeReport = Join-Path $projectDir "$smokeOutputDir\\report_execucao_empresas.csv"
    $smokeApproved = (-not $smokeRun.TimedOut) -and $smokeRun.ExitCode -eq 0 -and (Test-Path $smokeReport)
    New-TestResult `
        -Id "BUILD-001" `
        -Objective "Validar repetibilidade minima do release com o smoke oficial." `
        -Input $smokeScript `
        -Steps "Executar test_orchestrator_smoke.ps1 com timeout externo e limpeza preventiva de browser de automacao." `
        -Expected "Smoke passa e gera report de saida." `
        -Actual "TimedOut=$($smokeRun.TimedOut); ExitCode=$($smokeRun.ExitCode); report=$([bool](Test-Path $smokeReport))" `
        -Status $(if ($smokeApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence (Join-PathDisplay @($smokeScript, $smokeReport, $stdoutPath, $stderrPath))

    $delegatorContent = if (Test-Path $buildDelegatorScript) { Get-Content -Path $buildDelegatorScript -Raw } else { "" }
    $delegatorApproved = $delegatorContent -match "build_backend\.ps1"
    New-TestResult `
        -Id "BUILD-002" `
        -Objective "Validar que build_orchestrator.ps1 delega para o build consolidado." `
        -Input $buildDelegatorScript `
        -Steps "Inspecionar conteudo do script de build legado." `
        -Expected "Script delega para build_backend.ps1 e nao gera release incompleto." `
        -Actual $(if ($delegatorApproved) { "Delegacao encontrada no script." } else { "Nao foi encontrada delegacao para build_backend.ps1." }) `
        -Status $(if ($delegatorApproved) { "APROVADO" } else { "REPROVADO" }) `
        -Evidence $buildDelegatorScript

    New-TestResult `
        -Id "MAN-001" `
        -Objective "Validar sucesso operacional real contra o portal NFSe." `
        -Input "Credenciais validas e ambiente de homologacao/portal" `
        -Steps "Executar lote real com timeout folgado e observar download, logs e saida final." `
        -Expected "Ao menos um caso SUCESSO ou REVISAO_MANUAL homologado fora do ambiente de desenvolvimento." `
        -Actual "Nao executado nesta validacao automatizada local." `
        -Status "PENDENTE_MANUAL" `
        -Evidence "Executar manualmente apos homologacao local do release."

    $results | Export-Csv -Path $resultsCsvPath -NoTypeInformation -Encoding UTF8

    $approvedCount = @($results | Where-Object { $_.status -eq "APROVADO" }).Count
    $failedCount = @($results | Where-Object { $_.status -eq "REPROVADO" }).Count
    $pendingCount = @($results | Where-Object { $_.status -eq "PENDENTE_MANUAL" }).Count
    $summary = @()
    $summary += "# Homologacao do Backend Empacotado"
    $summary += ""
    $summary += "- Data: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $summary += "- Release avaliado: $releasePath"
    $summary += "- Release copiado para teste externo: $copiedReleasePath"
    $summary += "- Resultado geral: $approvedCount aprovados, $failedCount reprovados, $pendingCount pendentes manuais"
    $summary += ""
    $summary += "## Casos"
    $summary += ""

    foreach ($result in $results) {
        $summary += "### $($result.id)"
        $summary += "- Objetivo: $($result.objective)"
        $summary += "- Entrada: $($result.input)"
        $summary += "- Passos: $($result.steps)"
        $summary += "- Resultado esperado: $($result.expected)"
        $summary += "- Resultado obtido: $($result.actual)"
        $summary += "- Status: $($result.status)"
        $summary += "- Evidencia: $($result.evidence)"
        $summary += ""
    }

    Set-Content -Path $resultsMdPath -Value $summary -Encoding UTF8

    Write-Host "Homologacao concluida."
    Write-Host "CSV: $resultsCsvPath"
    Write-Host "Markdown: $resultsMdPath"
}
finally {
    Stop-AutomationBrowserProcesses
    if (Test-Path $copiedReleasePath) {
        Write-Host "Release copiado mantido para evidencia em: $copiedReleasePath"
    }
}

exit 0
