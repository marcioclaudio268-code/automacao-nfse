param(
    [switch]$ForceRefresh
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$browserRoot = Join-Path $projectDir "installer\chrome-for-testing"
$browserOutputDir = Join-Path $browserRoot "browser"
$manifestPath = Join-Path $browserOutputDir "VERSION.json"
$metadataUrl = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json"

function Remove-PathIfExists([string]$PathToRemove) {
    if (Test-Path $PathToRemove) {
        Remove-Item -Recurse -Force $PathToRemove
    }
}

function Download-File([string]$Url, [string]$DestinationPath) {
    $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
    if ($curl) {
        & $curl.Source -L --fail --retry 3 --output $DestinationPath $Url
        if ($LASTEXITCODE -ne 0) {
            throw "Falha ao baixar arquivo via curl.exe: $Url"
        }
        return
    }

    Invoke-WebRequest -Uri $Url -OutFile $DestinationPath
}

function Get-StableDownload([object]$Metadata, [string]$Platform, [string]$Key) {
    $channel = $Metadata.channels.Stable
    if (-not $channel) {
        throw "Canal Stable nao encontrado nos metadados do Chrome for Testing."
    }

    $download = $channel.downloads.$Key | Where-Object { $_.platform -eq $Platform } | Select-Object -First 1
    if (-not $download) {
        throw "Download '$Key' para plataforma '$Platform' nao encontrado."
    }

    return @{
        Version = $channel.version
        Url = $download.url
    }
}

New-Item -ItemType Directory -Force -Path $browserRoot | Out-Null

Write-Host "Consultando versao Stable oficial do Chrome for Testing..."
$metadata = Invoke-RestMethod -Uri $metadataUrl
$chromeDownload = Get-StableDownload -Metadata $metadata -Platform "win64" -Key "chrome"
$driverDownload = Get-StableDownload -Metadata $metadata -Platform "win64" -Key "chromedriver"
$tempDir = Join-Path $browserRoot ("_tmp_" + $chromeDownload.Version.Replace('.', '_'))

$expectedChromeExe = Join-Path $browserOutputDir "chrome-win64\chrome.exe"
$expectedDriverExe = Join-Path $browserOutputDir "chromedriver-win64\chromedriver.exe"
$versionMatches = $false
if ((-not $ForceRefresh) -and (Test-Path $manifestPath) -and (Test-Path $expectedChromeExe) -and (Test-Path $expectedDriverExe)) {
    try {
        $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
        $versionMatches = ($manifest.version -eq $chromeDownload.Version)
    }
    catch {
        $versionMatches = $false
    }
}

if ($versionMatches) {
    Write-Host "Chrome for Testing embutido ja esta atualizado em $($chromeDownload.Version)."
    return
}

Write-Host "Baixando Chrome for Testing Stable $($chromeDownload.Version) para win64..."
Remove-PathIfExists $tempDir
Remove-PathIfExists $browserOutputDir
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
New-Item -ItemType Directory -Force -Path $browserOutputDir | Out-Null

$chromeZip = Join-Path $tempDir "chrome-win64.zip"
$driverZip = Join-Path $tempDir "chromedriver-win64.zip"

Download-File -Url $chromeDownload.Url -DestinationPath $chromeZip
Download-File -Url $driverDownload.Url -DestinationPath $driverZip

Expand-Archive -LiteralPath $chromeZip -DestinationPath $tempDir -Force
Expand-Archive -LiteralPath $driverZip -DestinationPath $tempDir -Force

$chromeDir = Join-Path $tempDir "chrome-win64"
$driverDir = Join-Path $tempDir "chromedriver-win64"
if (-not (Test-Path (Join-Path $chromeDir "chrome.exe"))) {
    throw "chrome.exe nao encontrado apos extracao do Chrome for Testing."
}
if (-not (Test-Path (Join-Path $driverDir "chromedriver.exe"))) {
    throw "chromedriver.exe nao encontrado apos extracao do Chrome for Testing."
}

Move-Item -Path $chromeDir -Destination (Join-Path $browserOutputDir "chrome-win64")
Move-Item -Path $driverDir -Destination (Join-Path $browserOutputDir "chromedriver-win64")

$manifestPayload = [ordered]@{
    version = $chromeDownload.Version
    chrome_url = $chromeDownload.Url
    chromedriver_url = $driverDownload.Url
    platform = "win64"
    fetched_at = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
}
$manifestPayload | ConvertTo-Json -Depth 3 | Set-Content -Path $manifestPath -Encoding UTF8

Remove-PathIfExists $tempDir

Write-Host "Chrome for Testing embutido preparado em installer\\chrome-for-testing\\browser"
