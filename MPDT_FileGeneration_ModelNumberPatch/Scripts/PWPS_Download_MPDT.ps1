<#
.SYNOPSIS
    Download an MPDT document from ProjectWise by URN or document name.

.PARAMETER DocumentURN
    Full ProjectWise document URN (for example from ACBOS MPDT extract).

.PARAMETER DocumentName
    Document name (with or without extension) to search in ProjectWise when URN is not supplied.

.PARAMETER OutputDir
    Local folder to store the downloaded file.

.PARAMETER DatasourceName
    ProjectWise datasource name.

.PARAMETER UserName
    ProjectWise username.
#>

[CmdletBinding()]
param(
    [string]$DocumentURN,

    [string]$DocumentName,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$DatasourceName = 'arcadis-uk-pw.bentley.com:arcadis-uk-07',
    [string]$UserName       = '_asc_user_automation'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($DocumentURN) -and [string]::IsNullOrWhiteSpace($DocumentName)) {
    Write-Error "Provide either -DocumentURN or -DocumentName"
    exit 1
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

try {
    Import-Module pwps_dab -ErrorAction Stop
} catch {
    Write-Error "Failed to import pwps_dab. Install with: Install-Module pwps_dab -Scope CurrentUser"
    exit 1
}

try {
    $PWPassword = Read-Host -Prompt "ProjectWise password for $UserName" -AsSecureString
    New-PWLogin -DatasourceName $DatasourceName -UserName $UserName -Password $PWPassword
} catch {
    Write-Error "PW login failed: $_"
    exit 1
}

try {
    $docs = @()
    $stem = ''

    if (-not [string]::IsNullOrWhiteSpace($DocumentURN)) {
        $docByUrn = Get-PWDocumentByURN -URN $DocumentURN -ErrorAction Stop
        if ($docByUrn) {
            $docs = @($docByUrn)
            $stem = [System.IO.Path]::GetFileNameWithoutExtension([string]$docByUrn.Name)
        }
    }

    if (-not $docs -and -not [string]::IsNullOrWhiteSpace($DocumentName)) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($DocumentName)
        $docs = Get-PWDocumentsBySearch -DocumentName $stem
    }

    if (-not $docs) {
        if (-not [string]::IsNullOrWhiteSpace($DocumentURN)) {
            throw "No ProjectWise document found for URN '$DocumentURN'"
        }
        throw "No ProjectWise document found for name '$DocumentName'"
    }

    $docs = $docs | Where-Object {
        $n = [string]($_.Name)
        $n -match '\.(xlsm|xlsx|xls)$' -or -not ($n -match '\.[a-zA-Z0-9]+$')
    }

    if (-not $docs) {
        throw "No Excel MPDT document found for '$DocumentName'"
    }

    $latest = $docs | Sort-Object -Property @{Expression = {
        if ($_.FileUpdateDate) { $_.FileUpdateDate }
        elseif ($_.DocumentUpdateDate) { $_.DocumentUpdateDate }
        elseif ($_.CreateDate) { $_.CreateDate }
        else { Get-Date '1900-01-01' }
    }} -Descending | Select-Object -First 1
    if (-not $latest) {
        if (-not [string]::IsNullOrWhiteSpace($DocumentURN)) {
            throw "Could not select document for URN '$DocumentURN'"
        }
        throw "Could not select latest document for '$DocumentName'"
    }

    $downloaded = $false

    if (Get-Command Export-PWDocumentsSimple -ErrorAction SilentlyContinue) {
        try {
            # First try unmanaged export for speed
            $latest | Export-PWDocumentsSimple -TargetFolder $OutputDir -ErrorAction Stop | Out-Null
            $downloaded = $true
        } catch {
            Write-Warning "Unmanaged export failed, retrying managed export: $_"
            # Error 58019 commonly requires managed copy export
            $latest | Export-PWDocumentsSimple -TargetFolder $OutputDir -ExportManagedDocs -ErrorAction Stop | Out-Null
            $downloaded = $true
        }
    }

    if (-not $downloaded -and (Get-Command Export-PWDocuments -ErrorAction SilentlyContinue)) {
        Export-PWDocuments -InputDocuments @($latest) -TargetFolder $OutputDir -ErrorAction Stop | Out-Null
        $downloaded = $true
    }

    if (-not $downloaded) {
        throw "No supported pwps_dab download cmdlet available (Export-PWDocumentsSimple/Export-PWDocuments)."
    }

    if ([string]::IsNullOrWhiteSpace($stem)) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension([string]$latest.Name)
    }

    $cand = Get-ChildItem -Path $OutputDir -File -Recurse |
        Where-Object { $_.Name -like "$stem*" -and $_.Extension -match '^\.xls(m|x)?$' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $cand) {
        throw "Download command completed but file not found in '$OutputDir'"
    }

    Write-Output "DownloadedFile=$($cand.FullName)"
    Write-Output "DocumentGUID=$($latest.DocumentGUID)"
    Write-Output "DocumentURN=$($latest.DocumentURN)"
} catch {
    Write-Error "PW download failed: $_"
    exit 1
} finally {
    try { Remove-PWLogin } catch { }
}
