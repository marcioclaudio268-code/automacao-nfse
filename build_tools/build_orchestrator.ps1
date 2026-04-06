param(
    [switch]$InstallMissing
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$backendBuildScript = Join-Path $scriptDir "build_backend.ps1"

if (-not (Test-Path $backendBuildScript)) {
    throw "Script de build do backend nao encontrado em: $backendBuildScript"
}

& $backendBuildScript @PSBoundParameters
