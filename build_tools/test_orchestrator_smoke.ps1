param(
    [string]$EmpresasArquivo = "build_tools\smoke_empresas.csv",
    [string]$OutputDir = "build_tools\smoke_output",
    [int]$TimeoutSeconds = 30
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$exePath = Join-Path $projectDir "dist\AutomacaoNFSe-Orquestrador\AutomacaoNFSe-Orquestrador.exe"
$mainExePath = Join-Path $projectDir "dist\AutomacaoNFSe-Orquestrador\AutomacaoNFSe-Main\AutomacaoNFSe-Main.exe"
$empresasPath = Join-Path $projectDir $EmpresasArquivo
$outputPath = Join-Path $projectDir $OutputDir
$reportPath = Join-Path $outputPath "report_execucao_empresas.csv"

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

if (-not (Test-Path $exePath)) {
    throw "Executavel nao encontrado em: $exePath"
}

if (-not (Test-Path $mainExePath)) {
    throw "Executavel do Main nao encontrado em: $mainExePath"
}

if (-not (Test-Path $empresasPath)) {
    throw "Arquivo de empresas nao encontrado em: $empresasPath"
}

Push-Location $projectDir
try {
    if (Test-Path $outputPath) {
        Remove-Item -Recurse -Force $outputPath
    }
    New-Item -ItemType Directory -Force -Path $outputPath | Out-Null

    $env:EMPRESAS_ARQUIVO = $empresasPath
    $env:OUTPUT_BASE_DIR = $outputPath
    $env:CONTINUAR_DE_ONDE_PAROU = "0"
    $env:USAR_CHECKPOINT = "0"
    $env:MAX_TENTATIVAS_EMPRESA = "1"
    $env:LOGIN_WAIT_SECONDS = "1"
    $env:TIMEOUT_PROCESSO_MAIN = "1"

    $stdoutPath = Join-Path $outputPath "smoke_stdout.txt"
    $stderrPath = Join-Path $outputPath "smoke_stderr.txt"

    $proc = $null
    try {
        Stop-AutomationBrowserProcesses

        if (Test-Path $stdoutPath) { Remove-Item -Force $stdoutPath }
        if (Test-Path $stderrPath) { Remove-Item -Force $stderrPath }

        $proc = Start-Process `
            -FilePath $exePath `
            -WorkingDirectory $projectDir `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath `
            -PassThru

        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            Stop-ProcessTree -ProcessId $proc.Id
            $proc.WaitForExit(5000) | Out-Null
            throw "Smoke test excedeu timeout externo de ${TimeoutSeconds}s."
        }
    }
    finally {
        try {
            if ($proc -and -not $proc.HasExited) {
                Stop-ProcessTree -ProcessId $proc.Id
                $proc.WaitForExit(5000) | Out-Null
            }
        } catch {}

        Stop-AutomationBrowserProcesses
    }

    $exitCode = if ($proc) { [int]$proc.ExitCode } else { -1 }
    if ($exitCode -ne 0) {
        throw "Smoke test falhou com exit code $exitCode"
    }

    if (-not (Test-Path $reportPath)) {
        throw "Report nao foi gerado em: $reportPath"
    }

    $reportLines = Get-Content -Path $reportPath
    if ($reportLines.Count -lt 2) {
        throw "Smoke test inconclusivo: report foi gerado sem linhas de resultado."
    }

    $reportRows = @(Import-Csv -Path $reportPath -Delimiter ';')
    if ($reportRows.Count -lt 1) {
        throw "Smoke test inconclusivo: nenhuma linha de report foi carregada."
    }

    $firstRow = $reportRows[0]
    if ($firstRow.status -ne "FALHA") {
        throw "Smoke test falhou: status inesperado no report: $($firstRow.status)"
    }

    if ($firstRow.motivo -notmatch "Timeout de execucao do backend principal") {
        throw "Smoke test falhou: motivo inesperado no report: $($firstRow.motivo)"
    }

    if ($firstRow.acao_recomendada -ne "ABRIR_DEBUG") {
        throw "Smoke test falhou: acao_recomendada inesperada: $($firstRow.acao_recomendada)"
    }

    Write-Host "Smoke test concluido. Verifique a saida em: $outputPath"
}
finally {
    try { Stop-AutomationBrowserProcesses } catch {}
    Pop-Location
}
