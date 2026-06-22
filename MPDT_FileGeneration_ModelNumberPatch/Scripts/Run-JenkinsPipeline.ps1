[CmdletBinding()]
param(
    [string]$TargetUAID2 = $env:TARGET_UAID2,
    [ValidateSet('AUTO', 'MPDT', 'ACBOS')]
    [string]$FileType = 'AUTO',
    [ValidateSet('all', 'step1', 'step2', 'step3', 'step4', 'step5', 'step6')]
    [string]$Step = 'all',
    [string]$LogLevel = 'INFO',
    [string]$ProjectWiseBin = $env:PROJECTWISE_BIN,
    [switch]$PublishProjectWise,
    [switch]$SkipPwPreflight,
    [switch]$InstallDependencies
)

$ErrorActionPreference = 'Stop'

$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

function Require-Env {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable($Name))) {
        throw "Required environment variable is missing: $Name"
    }
}

Write-Host "Repository: $RepoRoot"
Write-Host "Step: $Step"
Write-Host "FileType: $FileType"
Write-Host "TargetUAID2: $TargetUAID2"
Write-Host "PublishProjectWise: $($PublishProjectWise.IsPresent)"

Require-Env 'SHAREPOINT_CLIENT_ID'
Require-Env 'SHAREPOINT_CLIENT_SECRET'

$env:MPDT_PUBLISH = if ($PublishProjectWise) { 'true' } else { 'false' }

if (-not [string]::IsNullOrWhiteSpace($ProjectWiseBin)) {
    $env:PROJECTWISE_BIN = $ProjectWiseBin
}

if ($PublishProjectWise) {
    Require-Env 'PW_PASSWORD'
    if (-not $SkipPwPreflight) {
        $pwBinArg = if ([string]::IsNullOrWhiteSpace($ProjectWiseBin)) { 'C:\Program Files\Bentley\ProjectWise\bin' } else { $ProjectWiseBin }
        Write-Host 'Running ProjectWise runtime preflight...'
        powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\Test-PWPSRuntime.ps1 -ProjectWiseBin $pwBinArg -FailOnError
        if ($LASTEXITCODE -ne 0) {
            throw "ProjectWise runtime preflight failed with exit code $LASTEXITCODE"
        }
    }
}

python --version
if ($LASTEXITCODE -ne 0) { throw 'python is not available on PATH.' }

if ($InstallDependencies) {
    python -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw 'pip upgrade failed.' }
    python -m pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) { throw 'pip install failed.' }
}

$pythonArgs = @('main.py', '--step', $Step, '--file-type', $FileType, '--log-level', $LogLevel)
if (-not [string]::IsNullOrWhiteSpace($TargetUAID2)) {
    $targets = @($TargetUAID2 -split '[,\s]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($targets.Count -gt 0) {
        $pythonArgs += '--target-uaid2'
        $pythonArgs += $targets
    }
}

Write-Host "Running: python $($pythonArgs -join ' ')"
python @pythonArgs
if ($LASTEXITCODE -ne 0) {
    throw "Pipeline failed with exit code $LASTEXITCODE"
}

Write-Host 'Pipeline completed successfully.'