<#
.SYNOPSIS
    Extract ACBOS and MPDT document metadata from ProjectWise.

.DESCRIPTION
    Connects to ProjectWise using the pwps_dab module, runs a saved search
    (Z_ACBOS_MPDT), and exports document metadata to an Excel file.
    The exported file is used by subsequent Python pipeline steps to:
      - Determine which documents already exist in PW
      - Find the latest revision of each document
      - Decide whether to create MPDT or ACBOS for each UAID_2

.PARAMETER OutputPath
    Full path to the output Excel file.  Defaults to the Input\ folder
    relative to the script's parent directory.

.EXAMPLE
    pwsh -File PWPS_Data_Extract.ps1
    pwsh -File PWPS_Data_Extract.ps1 -OutputPath "C:\MyFolder\ACBOS MPDT.xlsx"
#>

[CmdletBinding()]
param(
    [string]$OutputPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Resolve output path
# ---------------------------------------------------------------------------
if (-not $OutputPath) {
    $ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $WorkspaceDir = Split-Path -Parent $ScriptDir
    $OutputPath = Join-Path $WorkspaceDir "Input\ACBOS MPDT.xlsx"
}

# Ensure output directory exists
$OutputDir = Split-Path -Parent $OutputPath
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

Write-Host "Output will be written to: $OutputPath"

# ---------------------------------------------------------------------------
# Import pwps_dab module
# ---------------------------------------------------------------------------
# Preserve current PSModulePath and append common user module locations.
$modulePaths = @()
foreach ($rawPath in @(
    $env:PSModulePath,
    [System.Environment]::GetEnvironmentVariable('PSModulePath', 'User'),
    [System.Environment]::GetEnvironmentVariable('PSModulePath', 'Machine')
)) {
    if (-not [string]::IsNullOrWhiteSpace($rawPath)) {
        foreach ($p in ($rawPath -split ';')) {
            if (-not [string]::IsNullOrWhiteSpace($p)) {
                $modulePaths += $p.Trim()
            }
        }
    }
}

$myDocs = [System.Environment]::GetFolderPath('MyDocuments')
$candidateModulePaths = @(
    (Join-Path $myDocs 'PowerShell\Modules'),
    (Join-Path $myDocs 'WindowsPowerShell\Modules'),
    (Join-Path $HOME 'Documents\PowerShell\Modules'),
    (Join-Path $HOME 'Documents\WindowsPowerShell\Modules')
)
foreach ($p in $candidateModulePaths) {
    if (Test-Path $p) {
        $modulePaths += $p.Trim()
    }
}
$env:PSModulePath = (
    $modulePaths |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object -Unique
) -join ';'

$importOk = $false
$importError = $null
$compatUsed = $false

$alreadyLoaded = Get-Module -Name pwps_dab | Select-Object -First 1
if ($alreadyLoaded) {
    $importOk = $true
}

if (-not $importOk) {
    try {
        Import-Module pwps_dab -Force -ErrorAction Stop
        $importOk = $true
    } catch {
        $importError = $_
    }
}

if (-not $importOk -and $PSEdition -eq 'Core') {
    try {
        # Some machines only expose pwps_dab through Windows PowerShell compatibility.
        Import-Module pwps_dab -UseWindowsPowerShell -Force -ErrorAction Stop
        $importOk = $true
        $compatUsed = $true
    } catch {
        $importError = $_
    }
}

if (-not $importOk) {
    $mods = Get-Module -ListAvailable -Name pwps_dab | Select-Object Name, Version, ModuleBase
    $installed = (Get-Module -ListAvailable | Where-Object Name -Like '*pw*' | Select-Object -ExpandProperty Name) -join ', '
    $modsText = if ($mods) { ($mods | ForEach-Object { "{0} {1} @ {2}" -f $_.Name, $_.Version, $_.ModuleBase }) -join "`n" } else { '<none>' }
    Write-Error @"
Module 'pwps_dab' could not be imported.
PS Edition: $PSEdition
Visible pwps_dab modules:
$modsText
Modules with 'pw' in name visible now: $installed
Current PSModulePath:
$env:PSModulePath
If needed, run:
  Install-Module pwps_dab -Scope CurrentUser -Force
Then open a NEW PowerShell window and rerun.
"@
    if ($importError) {
        Write-Error "Import error details: $importError"
    }
    return   # use return, not exit, so dot-sourcing doesn't kill your terminal
}

$loaded = Get-Module -Name pwps_dab | Select-Object -First 1
$detected = Get-Module -ListAvailable -Name pwps_dab | Sort-Object Version -Descending | Select-Object -First 1
$verText = if ($detected) { [string]$detected.Version } elseif ($loaded) { [string]$loaded.Version } else { 'unknown' }
if ($compatUsed) {
    Write-Host "pwps_dab $verText loaded via -UseWindowsPowerShell compatibility."
} elseif ($alreadyLoaded) {
    Write-Host "pwps_dab $verText already loaded in current session."
} else {
    Write-Host "pwps_dab $verText loaded OK."
}

# ---------------------------------------------------------------------------
# Connect to ProjectWise
# NOTE: For automated runs, store the password in a secure credential store
#       rather than plain text. Replace the block below with a secure method.
# ---------------------------------------------------------------------------
try {
    $PWPassword = Read-Host -Prompt "ProjectWise password for _asc_user_automation" -AsSecureString
    New-PWLogin `
        -DatasourceName 'arcadis-uk-pw.bentley.com:arcadis-uk-07' `
        -UserName '_asc_user_automation' `
        -Password $PWPassword
    Write-Host "Connected to ProjectWise."
} catch {
    Write-Error "ProjectWise login failed: $_"
    return
}

# ---------------------------------------------------------------------------
# Run saved search to get document GUIDs
# ---------------------------------------------------------------------------
try {
    Write-Host "Running saved search 'Z_ACBOS_MPDT' …"
    $PWDocumentGUIDs = Get-PWDocumentsBySearch -SearchName 'Z_ACBOS_MPDT' |
                       Select-Object -ExpandProperty DocumentGUID
    Write-Host "  Found $($PWDocumentGUIDs.Count) document(s)."
} catch {
    Write-Error "Failed to run PW search: $_"
    return
}

# ---------------------------------------------------------------------------
# Fetch full document metadata
# ---------------------------------------------------------------------------
try {
    $PWDocuments = Get-PWDocumentsByGUIDs -DocumentGUIDs $PWDocumentGUIDs
    Write-Host "  Retrieved metadata for $($PWDocuments.Count) document(s)."
} catch {
    Write-Error "Failed to retrieve document metadata: $_"
    return
}

# ---------------------------------------------------------------------------
# Build DataTable
# ---------------------------------------------------------------------------
$DT = [System.Data.DataTable]::new('PW Documents')
$columns = @(
    'DocumentName',
    'Description',
    'Version',
    'FileName',
    'FullPath',
    'WorkflowState',
    'URN',
    'FileUpdated',
    'PW_PROJECT_NAME',
    'ASSET_ID'
)
foreach ($col in $columns) {
    [void]$DT.Columns.Add($col, [string])
}

function Get-DocPropValue {
    param(
        [Parameter(Mandatory = $true)] $Doc,
        [Parameter(Mandatory = $true)][string] $PropertyName
    )
    $prop = $Doc.PSObject.Properties[$PropertyName]
    if ($null -eq $prop -or $null -eq $prop.Value) {
        return ''
    }
    return [string]$prop.Value
}

function Convert-FileUpdatedValue {
    param(
        [Parameter(Mandatory = $true)] $Doc,
        [Parameter(Mandatory = $true)][string] $PropertyName
    )

    $prop = $Doc.PSObject.Properties[$PropertyName]
    if ($null -eq $prop -or $null -eq $prop.Value) {
        return ''
    }

    $raw = $prop.Value

    # 1) Native DateTime
    if ($raw -is [datetime]) {
        return ([datetime]$raw).ToString('yyyy-MM-dd HH:mm:ss')
    }

    # 2) OA / Excel serial number
    $num = 0.0
    if ([double]::TryParse([string]$raw, [ref]$num)) {
        try {
            return [datetime]::FromOADate($num).ToString('yyyy-MM-dd HH:mm:ss')
        } catch {
            # continue with other parsing
        }
    }

    # 3) Date string
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse([string]$raw, [ref]$dt)) {
        return $dt.ToString('yyyy-MM-dd HH:mm:ss')
    }

    # 4) Fallback as string
    return [string]$raw
}

foreach ($Doc in $PWDocuments) {
    $DR = $DT.NewRow()
    $DR['DocumentName']    = Get-DocPropValue -Doc $Doc -PropertyName 'Name'
    $DR['Description']     = Get-DocPropValue -Doc $Doc -PropertyName 'Description'
    $DR['Version']         = Get-DocPropValue -Doc $Doc -PropertyName 'Version'
    $DR['FileName']        = Get-DocPropValue -Doc $Doc -PropertyName 'FileName'
    $DR['FullPath']        = Get-DocPropValue -Doc $Doc -PropertyName 'FullPath'
    $DR['WorkflowState']   = Get-DocPropValue -Doc $Doc -PropertyName 'WorkflowState'
    $DR['URN']             = Get-DocPropValue -Doc $Doc -PropertyName 'URN'
    $DR['FileUpdated']     = Convert-FileUpdatedValue -Doc $Doc -PropertyName 'FileUpdated'
    $DR['PW_PROJECT_NAME'] = Get-DocPropValue -Doc $Doc -PropertyName 'ProjectName'
    # ASSET_ID is a custom PW attribute and may be absent on some documents.
    $DR['ASSET_ID'] = Get-DocPropValue -Doc $Doc -PropertyName 'ASSET_ID'
    $DT.Rows.Add($DR)
}

# ---------------------------------------------------------------------------
# Export to Excel
# ---------------------------------------------------------------------------
$exportSuccess = $false

# Try native pwps_dab export (wrap DataTable in array for compatibility mode)
try {
    New-XLSXWorkbook -InputTables @($DT) -OutputFileName $OutputPath
    Write-Host "PW extract written to: $OutputPath"
    $exportSuccess = $true
} catch {
    Write-Host "Native export failed ($_). Attempting Python fallback..."
}

# Fallback: use Python to export via pandas
if (-not $exportSuccess) {
    try {
        $csvTemp = [System.IO.Path]::GetTempFileName() + ".csv"
        $DT | Export-Csv -Path $csvTemp -NoTypeInformation -Encoding UTF8
        
        $pythonScript = @"
import pandas as pd
import sys
csv_path = r'$csvTemp'
xlsx_path = r'$OutputPath'
df = pd.read_csv(csv_path, dtype=str)
df.to_excel(xlsx_path, sheet_name='PW Documents', index=False)
print(f'Python export written: {xlsx_path}')
"@
        
        $pythonScript | python
        if ($LASTEXITCODE -eq 0) {
            Write-Host "PW extract written to: $OutputPath (via Python)"
            $exportSuccess = $true
        } else {
            Write-Error "Python export failed"
        }
        
        Remove-Item $csvTemp -ErrorAction SilentlyContinue
    } catch {
        Write-Error "Fallback export failed: $_"
    }
}

if (-not $exportSuccess) {
    Write-Error "Failed to write Excel output via both methods."
    return
}

# ---------------------------------------------------------------------------
# Disconnect
# ---------------------------------------------------------------------------
try { Remove-PWLogin } catch { <# ignore #> }

Write-Host "Done."
