param(
    [switch]$InstallMissing
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$venvPython = Join-Path $projectDir ".venv\Scripts\python.exe"
$guiSpecPath = Join-Path $projectDir "AutomacaoNFSe.spec"
$backendBuildScript = Join-Path $scriptDir "build_backend.ps1"
$guiBuildDir = Join-Path $projectDir "build\AutomacaoNFSe"
$guiDistDir = Join-Path $projectDir "dist\AutomacaoNFSe"
$backendDistDir = Join-Path $projectDir "dist\AutomacaoNFSe-Orquestrador"
$guiBackendDir = Join-Path $guiDistDir "backend"

if (-not (Test-Path $venvPython)) {
    throw "Python da .venv nao encontrado em: $venvPython"
}

if (-not (Test-Path $guiSpecPath)) {
    throw "Arquivo .spec da GUI nao encontrado em: $guiSpecPath"
}

if (-not (Test-Path $backendBuildScript)) {
    throw "Script de build do backend nao encontrado em: $backendBuildScript"
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

    Ensure-Package "PySide6"
    Ensure-Package "openpyxl"
    Ensure-Package "selenium"

    Write-Host "Gerando build do backend para empacotar junto da GUI..."
    & $backendBuildScript @PSBoundParameters
    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao executar o build do backend."
    }

    Remove-PathIfExists $guiBuildDir
    Remove-PathIfExists $guiDistDir

    Write-Host "Gerando build da GUI..."
    & $venvPython -m PyInstaller --noconfirm --clean $guiSpecPath
    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao executar o PyInstaller para a GUI."
    }

    if (-not (Test-Path $guiDistDir)) {
        throw "Pasta da GUI nao foi gerada em: $guiDistDir"
    }

    if (-not (Test-Path $backendDistDir)) {
        throw "Pasta do backend nao foi gerada em: $backendDistDir"
    }

    Remove-PathIfExists $guiBackendDir
    New-Item -ItemType Directory -Force -Path $guiBackendDir | Out-Null
    Copy-Item -Path (Join-Path $backendDistDir '*') -Destination $guiBackendDir -Recurse -Force

    $guiExePath = Join-Path $guiDistDir "AutomacaoNFSe.exe"
    $backendExePath = Join-Path $guiBackendDir "AutomacaoNFSe-Orquestrador.exe"
    $embeddedChromePath = Join-Path $guiBackendDir "AutomacaoNFSe-Main\browser\chrome-win64\chrome.exe"
    $embeddedDriverPath = Join-Path $guiBackendDir "AutomacaoNFSe-Main\browser\chromedriver-win64\chromedriver.exe"
    if (-not (Test-Path $guiExePath)) {
        throw "Executavel da GUI nao encontrado apos build: $guiExePath"
    }
    if (-not (Test-Path $backendExePath)) {
        throw "Executavel do backend nao encontrado dentro da GUI: $backendExePath"
    }
    if (-not (Test-Path $embeddedChromePath)) {
        throw "Chrome embutido nao encontrado dentro da GUI: $embeddedChromePath"
    }
    if (-not (Test-Path $embeddedDriverPath)) {
        throw "Chromedriver embutido nao encontrado dentro da GUI: $embeddedDriverPath"
    }

    Write-Host "GUI empacotada com sucesso em dist\\AutomacaoNFSe"
}
finally {
    Pop-Location
}
