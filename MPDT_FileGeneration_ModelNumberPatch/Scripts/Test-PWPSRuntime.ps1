[CmdletBinding()]
param(
    [string]$ProjectWiseBin = 'C:\Program Files\Bentley\ProjectWise\bin',
    [string]$DatasourceName = 'arcadis-uk-pw.bentley.com:arcadis-uk-07',
    [string]$UserName = '_asc_user_automation'
)

$ErrorActionPreference = 'Stop'

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "=== $Title ==="
}

Write-Section 'Environment'
Write-Host "Process is 64-bit: $([Environment]::Is64BitProcess)"
Write-Host "PowerShell version: $($PSVersionTable.PSVersion)"
Write-Host "PWD bin exists: $(Test-Path $ProjectWiseBin)"
Write-Host "dmscli.dll exists: $(Test-Path (Join-Path $ProjectWiseBin 'dmscli.dll'))"

Write-Section 'PATH entries'
$env:PATH.Split(';') |
    Where-Object { $_ -match 'Bentley|ProjectWise' } |
    Sort-Object -Unique |
    ForEach-Object { Write-Host $_ }

Write-Section 'Native LoadLibrary test'
$nativeSource = @'
using System;
using System.Runtime.InteropServices;
public static class Kernel32 {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr LoadLibrary(string lpFileName);
}
'@

Add-Type -TypeDefinition $nativeSource -Language CSharp
$dllPath = Join-Path $ProjectWiseBin 'dmscli.dll'
$handle = [Kernel32]::LoadLibrary($dllPath)
if ($handle -eq [IntPtr]::Zero) {
    $win32Error = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    Write-Host "LoadLibrary failed for '$dllPath' with Win32 error: $win32Error"
} else {
    Write-Host "LoadLibrary succeeded: $handle"
}

Write-Section 'pwps_dab module'
try {
    Import-Module pwps_dab -ErrorAction Stop
    Write-Host 'Import-Module pwps_dab: OK'
    Get-Command -Module pwps_dab New-PWLogin, Get-PWDocumentsBySearch, New-PWDocument | Format-Table -AutoSize
} catch {
    Write-Host "Import-Module pwps_dab failed: $_"
}

Write-Section 'Optional login test'
Write-Host "Run manually only if LoadLibrary succeeded:"
Write-Host "$pw = Read-Host -Prompt 'ProjectWise password' -AsSecureString"
Write-Host "New-PWLogin -DatasourceName '$DatasourceName' -UserName '$UserName' -Password $pw"