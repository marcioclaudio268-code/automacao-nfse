param(
    [switch]$InstallMissing
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$venvPython = Join-Path $projectDir ".venv\Scripts\python.exe"
$mainSpecPath = Join-Path $projectDir "AutomacaoNFSe-Main.spec"
$orchestratorSpecPath = Join-Path $projectDir "AutomacaoNFSe-Orquestrador.spec"
$browserFetchScript = Join-Path $scriptDir "fetch_embedded_browser.ps1"
$embeddedBrowserSourceDir = Join-Path $projectDir "installer\chrome-for-testing\browser"
$mainBuildDir = Join-Path $projectDir "build\AutomacaoNFSe-Main"
$mainDistDir = Join-Path $projectDir "dist\AutomacaoNFSe-Main"
$orchestratorBuildDir = Join-Path $projectDir "build\AutomacaoNFSe-Orquestrador"
$orchestratorDistDir = Join-Path $projectDir "dist\AutomacaoNFSe-Orquestrador"
$orchestratorMainDir = Join-Path $orchestratorDistDir "AutomacaoNFSe-Main"

if (-not (Test-Path $venvPython)) {
    throw "Python da .venv nao encontrado em: $venvPython"
}

if (-not (Test-Path $mainSpecPath)) {
    throw "Arquivo .spec do Main nao encontrado em: $mainSpecPath"
}

if (-not (Test-Path $orchestratorSpecPath)) {
    throw "Arquivo .spec do Orquestrador nao encontrado em: $orchestratorSpecPath"
}

if (-not (Test-Path $browserFetchScript)) {
    throw "Script de preparo do browser embutido nao encontrado em: $browserFetchScript"
}

Push-Location $projectDir
try {
    function Remove-PathIfExists([string]$PathToRemove) {
        if (-not (Test-Path $PathToRemove)) {
            return
        }

        Remove-Item -Recurse -Force $PathToRemove
        Start-Sleep -Milliseconds 300
    }

    function Ensure-Package([string]$PackageName) {
        & $venvPython -m pip show $PackageName | Out-Null
        if ($LASTEXITCODE -eq 0) {
            return
        }

        if (-not $InstallMissing) {
            throw "$PackageName nao esta instalado na .venv. Rode novamente com -InstallMissing."
        }

        Write-Host "Instalando $PackageName na .venv..."
        & $venvPython -m pip install $PackageName
        if ($LASTEXITCODE -ne 0) {
            throw "Falha ao instalar $PackageName na .venv."
        }
    }

    & $venvPython -m pip show pyinstaller | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller nao esta instalado na .venv."
    }

    Ensure-Package "openpyxl"
    Ensure-Package "selenium"

    Write-Host "Preparando Chrome for Testing embutido..."
    & $browserFetchScript
    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao preparar o browser embutido."
    }

    Remove-PathIfExists $mainBuildDir
    Remove-PathIfExists $mainDistDir

    Write-Host "Gerando build do Main..."
    & $venvPython -m PyInstaller --noconfirm --clean $mainSpecPath
    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao executar o PyInstaller para o Main."
    }

    Remove-PathIfExists $orchestratorBuildDir
    Remove-PathIfExists $orchestratorDistDir

    Write-Host "Gerando build do Orquestrador..."
    & $venvPython -m PyInstaller --noconfirm --clean $orchestratorSpecPath
    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao executar o PyInstaller para o Orquestrador."
    }

    if (-not (Test-Path $mainDistDir)) {
        throw "Pasta do Main nao foi gerada em: $mainDistDir"
    }

    if (-not (Test-Path $orchestratorDistDir)) {
        throw "Pasta do Orquestrador nao foi gerada em: $orchestratorDistDir"
    }

    if (-not (Test-Path $embeddedBrowserSourceDir)) {
        throw "Pasta do browser embutido nao encontrada em: $embeddedBrowserSourceDir"
    }

    $mainBrowserDir = Join-Path $mainDistDir "browser"
    Remove-PathIfExists $mainBrowserDir
    New-Item -ItemType Directory -Force -Path $mainBrowserDir | Out-Null
    Copy-Item -Path (Join-Path $embeddedBrowserSourceDir '*') -Destination $mainBrowserDir -Recurse -Force

    if (Test-Path $orchestratorMainDir) {
        Remove-Item -Recurse -Force $orchestratorMainDir
    }

    Copy-Item -Path $mainDistDir -Destination $orchestratorMainDir -Recurse -Force

    $mainExePath = Join-Path $orchestratorMainDir "AutomacaoNFSe-Main.exe"
    $embeddedChromePath = Join-Path $orchestratorMainDir "browser\chrome-win64\chrome.exe"
    $embeddedDriverPath = Join-Path $orchestratorMainDir "browser\chromedriver-win64\chromedriver.exe"
    if (-not (Test-Path $mainExePath)) {
        throw "Executavel do Main nao encontrado apos copia: $mainExePath"
    }
    if (-not (Test-Path $embeddedChromePath)) {
        throw "Chrome embutido nao encontrado apos copia: $embeddedChromePath"
    }
    if (-not (Test-Path $embeddedDriverPath)) {
        throw "Chromedriver embutido nao encontrado apos copia: $embeddedDriverPath"
    }

    Write-Host "Backend empacotado com sucesso em dist\\AutomacaoNFSe-Orquestrador"
}
finally {
    Pop-Location
}
