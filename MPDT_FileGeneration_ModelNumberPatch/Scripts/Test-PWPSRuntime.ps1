[CmdletBinding()]
param(
    [string]$ProjectWiseBin = 'C:\Program Files\Bentley\ProjectWise\bin',
    [string]$DatasourceName = 'arcadis-uk-pw.bentley.com:arcadis-uk-07',
    [string]$UserName = '_asc_user_automation',
    [switch]$FailOnError
)

$ErrorActionPreference = 'Stop'
$failed = $false

function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Host "=== $Title ==="
}

function Mark-Failed {
    param([string]$Message)
    $script:failed = $true
    Write-Host "FAILED: $Message"
}

Write-Section 'Environment'
Write-Host "Process is 64-bit: $([Environment]::Is64BitProcess)"
Write-Host "PowerShell version: $($PSVersionTable.PSVersion)"
Write-Host "ProjectWise bin: $ProjectWiseBin"
Write-Host "PWD bin exists: $(Test-Path $ProjectWiseBin)"
Write-Host "dmscli.dll exists: $(Test-Path (Join-Path $ProjectWiseBin 'dmscli.dll'))"

if (-not [Environment]::Is64BitProcess) {
    Mark-Failed 'ProjectWise pwps_dab requires a 64-bit PowerShell process.'
}

if (-not (Test-Path $ProjectWiseBin)) {
    Mark-Failed "ProjectWise bin folder not found: $ProjectWiseBin"
} else {
    $ProjectWiseBin = (Resolve-Path -LiteralPath $ProjectWiseBin).ProviderPath
    $pathParts = @($env:PATH -split ';' | Where-Object { $_ })
    if ($pathParts -notcontains $ProjectWiseBin) {
        $env:PATH = "$ProjectWiseBin;$env:PATH"
    }
}

Write-Section 'PATH entries'
$env:PATH.Split(';') |
    Where-Object { $_ -match 'Bentley|ProjectWise' } |
    Sort-Object -Unique |
    ForEach-Object { Write-Host $_ }

Write-Section 'Native LoadLibrary test'
$nativeSource = @"
using System;
using System.Runtime.InteropServices;
public static class Kernel32 {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool SetDllDirectory(string lpPathName);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr LoadLibrary(string lpFileName);
}
"@

try {
    if (-not ('Kernel32' -as [type])) {
        Add-Type -TypeDefinition $nativeSource -Language CSharp
    }
    [Kernel32]::SetDllDirectory($ProjectWiseBin) | Out-Null
    $dllPath = Join-Path $ProjectWiseBin 'dmscli.dll'
    $handle = [Kernel32]::LoadLibrary($dllPath)
    if ($handle -eq [IntPtr]::Zero) {
        $win32Error = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Mark-Failed "LoadLibrary failed for '$dllPath' with Win32 error: $win32Error"
        if ($win32Error -eq 126) {
            Write-Host 'Win32 error 126 usually means dmscli.dll exists but one of its dependent native DLLs cannot be found.'
            Write-Host 'Check that the full ProjectWise Explorer/client runtime is installed on this machine and that the ProjectWise bin folder is on PATH.'
        }
    } else {
        Write-Host "LoadLibrary succeeded: $handle"
    }
} catch {
    Mark-Failed "Native runtime test threw: $_"
}

Write-Section 'pwps_dab module'
try {
    Import-Module pwps_dab -ErrorAction Stop
    Write-Host 'Import-Module pwps_dab: OK'
    $module = Get-Module -Name pwps_dab | Select-Object -First 1
    if ($module) {
        Write-Host "pwps_dab loaded from: $($module.ModuleBase)"
        Write-Host "pwps_dab version: $($module.Version)"
    }
    Get-Command -Module pwps_dab New-PWLogin, Get-PWDocumentsBySearch, New-PWDocument -ErrorAction SilentlyContinue | Format-Table -AutoSize
} catch {
    Mark-Failed "Import-Module pwps_dab failed: $_"
}

Write-Section 'Optional login test'
Write-Host 'Run manually only if LoadLibrary and Import-Module succeeded:'
Write-Host "$pw = Read-Host -Prompt 'ProjectWise password' -AsSecureString"
Write-Host "New-PWLogin -DatasourceName '$DatasourceName' -UserName '$UserName' -Password `$pw"

if ($FailOnError -and $failed) {
    exit 1
}

if ($failed) {
    Write-Host ''
    Write-Host 'ProjectWise runtime preflight completed with failures.'
} else {
    Write-Host ''
    Write-Host 'ProjectWise runtime preflight passed.'
}