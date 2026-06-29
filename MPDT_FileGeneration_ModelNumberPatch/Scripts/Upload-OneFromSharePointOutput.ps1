<#
.SYNOPSIS
    Download one generated MPDT/ACBOS file from a SharePoint output_pw folder and upload it to ProjectWise.

.DESCRIPTION
    This is intended for the first Jenkins live-upload test. It downloads the selected file
    and PW_Upload_Metadata.xlsx from a SharePoint output_pw_<timestamp> folder, finds the
    metadata row for the selected file, writes an attributes JSON sidecar, and calls
    PWPS_Upload_PW.ps1 with the document name, version, folder path, and metadata.
#>

[CmdletBinding()]
param(
    [string]$SiteUrl = 'https://arcadiso365.sharepoint.com/teams/HS2ASC-RW',

    [Parameter(Mandatory = $true)]
    [string]$SharePointFolder,

    [Parameter(Mandatory = $true)]
    [string]$FileName,

    [string]$MetadataFileName = 'PW_Upload_Metadata.xlsx',
    [string]$LocalWorkDir = '',
    [string]$DatasourceName = 'arcadis-uk-pw.bentley.com:arcadis-uk-07',
    [string]$UserName = $env:PW_USERNAME,
    [string]$Password = $env:PW_PASSWORD,
    [string]$ProjectWiseBin = 'C:\Program Files\Bentley\ProjectWise\bin',
    [string]$UploadScript = '',
    [string]$WorkflowState = 'Work in Progress',
    [switch]$KeepDownloads
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($UserName)) {
    $UserName = '_asc_user_automation'
}
if ([string]::IsNullOrWhiteSpace($Password)) {
    throw 'PW_PASSWORD is required in the Jenkins environment for non-interactive ProjectWise upload.'
}
if ([string]::IsNullOrWhiteSpace($UploadScript)) {
    $UploadScript = Join-Path $PSScriptRoot 'PWPS_Upload_PW.ps1'
}
if (-not (Test-Path -LiteralPath $UploadScript)) {
    throw "Upload script not found: $UploadScript"
}
if ([string]::IsNullOrWhiteSpace($LocalWorkDir)) {
    $root = if ($env:WORKSPACE) { $env:WORKSPACE } else { (Split-Path -Parent $PSScriptRoot) }
    $LocalWorkDir = Join-Path $root 'pw_single_upload'
}

function Get-RequiredEnv {
    param([string]$Name)
    $value = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "Required environment variable is missing: $Name"
    }
    return $value
}

function Convert-ToServerRelativeUrl {
    param([string]$Value)
    $raw = [string]$Value
    $raw = $raw.Trim()
    if (-not $raw) { return '' }

    if ($raw -match '^https?://') {
        $uri = [Uri]$raw
        $query = $uri.Query.TrimStart('?')
        foreach ($part in $query.Split('&')) {
            if ($part -like 'id=*') {
                return [Uri]::UnescapeDataString($part.Substring(3))
            }
        }
        $path = [Uri]::UnescapeDataString($uri.AbsolutePath)
        $lower = $path.ToLowerInvariant()
        foreach ($marker in @('/:f:/r/', '/:x:/r/', '/:w:/r/')) {
            $idx = $lower.IndexOf($marker)
            if ($idx -ge 0) {
                return '/' + $path.Substring($idx + $marker.Length).TrimStart('/')
            }
        }
        return $path
    }

    if (-not $raw.StartsWith('/')) {
        $raw = '/' + $raw.TrimStart('/','\')
    }
    return $raw.TrimEnd('/')
}

function Join-SpPath {
    param([string]$Folder, [string]$Name)
    return ($Folder.TrimEnd('/') + '/' + $Name.TrimStart('/'))
}

function Normalize-Key {
    param([string]$Value)
    return (([string]$Value).ToLowerInvariant() -replace '[\s_]+', ' ').Trim()
}

function Get-RowValue {
    param($Row, [string[]]$Names)
    if (-not $Row) { return '' }
    $props = @{}
    foreach ($prop in $Row.PSObject.Properties) {
        $props[(Normalize-Key $prop.Name)] = $prop.Name
    }
    foreach ($name in $Names) {
        $key = Normalize-Key $name
        if ($props.ContainsKey($key)) {
            $value = $Row.($props[$key])
            if ($null -ne $value) { return ([string]$value).Trim() }
        }
    }
    return ''
}

function Get-DocKey {
    param([string]$Name)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension(([string]$Name).Trim())
    return ($stem -replace '-ACBOS$', '').ToLowerInvariant()
}

function Resolve-FolderPath {
    param($Row, [string]$DocumentName)
    $explicit = Get-RowValue $Row @('PWFolderPath', 'PW Folder Path', 'FolderPath', 'Folder Path')
    if ($explicit) { return $explicit }

    $fullPath = Get-RowValue $Row @('FullPath')
    if (-not $fullPath) { return '' }

    $parts = @($fullPath -split '[\\/]+' | Where-Object { $_ })
    if ($parts.Count -gt 0 -and (Get-DocKey $parts[-1]) -eq (Get-DocKey $DocumentName)) {
        return ($parts[0..($parts.Count - 2)] -join '\')
    }
    return $fullPath
}

function Write-AttributesJson {
    param($Row, [string]$OutPath)
    $core = @(
        'documentname','document name','document','description','document description',
        'version','revision','rev','filename','file name','localfilename',
        'filetype','file type','fullpath','urn','pwfolderpath','pw folder path',
        'folderpath','folder path','localfilepath','local file path','filepath','file path',
        'stagedfile','sourcefile','extension','currentpwrevision','currentrevision',
        'previousversion','nextversion','expectedrevision',
        'documentguid','projectguid','projectguidstring','projectid','documentid',
        'documentstatus','documentoutto','documentouttoname','documentcheckoutdate',
        'workflow','workflowstate','workflowid','stateid','applicationid','application',
        'applicationname','documentupdater','documentupdatername','documentupdatedate',
        'fileupdater','fileupdatername','fileupdated','fileupdatedate',
        'documentcreator','documentcreatorname','createdate','documenturn',
        'oldversion','oldname','oldfilename','olddescription','oldfullpath',
        'checkedoutlocalfilename','copiedoutlocalfilename','isabstract','isset',
        'documentownertype','documentownername','fullpathtostorageareafile',
        'status','mimetype','storagename','storageid','projectwisewebviewlink',
        'projectwiseweblink','webservicesgatewaydownloadlink','attributerecordcount',
        'attributes','nestedreferencecount','customattributes'
    ) | ForEach-Object { Normalize-Key $_ }

    $attrs = [ordered]@{}
    foreach ($prop in $Row.PSObject.Properties) {
        $key = Normalize-Key $prop.Name
        if ($core -contains $key) { continue }
        $value = if ($null -eq $prop.Value) { '' } else { [string]$prop.Value }
        $attrs[$prop.Name] = $value
    }
    $attrs | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $OutPath -Encoding UTF8
    return $OutPath
}

Write-Host '=== Upload one SharePoint output file to ProjectWise ==='
Write-Host "SiteUrl: $SiteUrl"
Write-Host "SharePointFolder: $SharePointFolder"
Write-Host "FileName: $FileName"
Write-Host "MetadataFileName: $MetadataFileName"
Write-Host "LocalWorkDir: $LocalWorkDir"

Import-Module PnP.PowerShell -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

$clientId = Get-RequiredEnv 'SHAREPOINT_CLIENT_ID'
$clientSecret = Get-RequiredEnv 'SHAREPOINT_CLIENT_SECRET'
Connect-PnPOnline -Url $SiteUrl -ClientId $clientId -ClientSecret $clientSecret -WarningAction SilentlyContinue

$spFolder = Convert-ToServerRelativeUrl $SharePointFolder
New-Item -ItemType Directory -Path $LocalWorkDir -Force | Out-Null

$fileUrl = Join-SpPath $spFolder $FileName
$metadataUrl = Join-SpPath $spFolder $MetadataFileName
Write-Host "Downloading file: $fileUrl"
Get-PnPFile -Url $fileUrl -Path $LocalWorkDir -FileName $FileName -AsFile -Force -ErrorAction Stop | Out-Null
Write-Host "Downloading metadata: $metadataUrl"
Get-PnPFile -Url $metadataUrl -Path $LocalWorkDir -FileName $MetadataFileName -AsFile -Force -ErrorAction Stop | Out-Null

$localFile = Join-Path $LocalWorkDir $FileName
$metadataPath = Join-Path $LocalWorkDir $MetadataFileName
if (-not (Test-Path -LiteralPath $localFile)) { throw "Downloaded file not found: $localFile" }
if (-not (Test-Path -LiteralPath $metadataPath)) { throw "Downloaded metadata not found: $metadataPath" }

$rows = Import-Excel -Path $metadataPath
if (-not $rows) { throw "Metadata workbook has no rows: $metadataPath" }

$fileLower = $FileName.ToLowerInvariant()
$metadataRow = $rows | Where-Object {
    $rowFile = (Get-RowValue $_ @('FileName', 'File Name', 'LocalFileName')).ToLowerInvariant()
    $rowLocal = Get-RowValue $_ @('LocalFilePath', 'Local File Path', 'FilePath', 'File Path')
    $rowLocalName = if ($rowLocal) { [System.IO.Path]::GetFileName($rowLocal).ToLowerInvariant() } else { '' }
    $rowFile -eq $fileLower -or $rowLocalName -eq $fileLower
} | Select-Object -First 1

if (-not $metadataRow) {
    throw "No matching row for $FileName in $MetadataFileName"
}

$documentName = Get-RowValue $metadataRow @('DocumentName', 'Document Name', 'Document')
if (-not $documentName) { $documentName = [System.IO.Path]::GetFileNameWithoutExtension($FileName) }
$description = Get-RowValue $metadataRow @('Description', 'Document Description')
if (-not $description) { $description = $documentName }
$version = Get-RowValue $metadataRow @('Version', 'Revision', 'Rev')
$folderPath = Resolve-FolderPath $metadataRow $documentName
if (-not $folderPath) { throw "Could not resolve PW folder path from metadata row for $FileName" }
$application = Get-RowValue $metadataRow @('Application', 'ApplicationName', 'Application Name')
$attrsJson = Write-AttributesJson $metadataRow ($localFile + '.metadata.json')

Write-Host 'Resolved upload metadata:'
Write-Host "  DocumentName: $documentName"
Write-Host "  Description: $description"
Write-Host "  Version: $version"
Write-Host "  PWFolderPath: $folderPath"
Write-Host "  WorkflowState: $WorkflowState"
Write-Host "  AttributesJson: $attrsJson"

$uploadArgs = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $UploadScript,
    '-FilePath', $localFile,
    '-DatasourceName', $DatasourceName,
    '-UserName', $UserName,
    '-Password', $Password,
    '-PWFolderPath', $folderPath,
    '-DocumentName', $documentName,
    '-Description', $description,
    '-Version', $version,
    '-AttributesJson', $attrsJson,
    '-ProjectWiseBin', $ProjectWiseBin,
    '-WorkflowState', $WorkflowState
)
if (-not [string]::IsNullOrWhiteSpace($application)) {
    $uploadArgs += @('-Application', $application)
}

& powershell.exe @uploadArgs

if ($LASTEXITCODE -ne 0) {
    throw "ProjectWise upload failed with exit code $LASTEXITCODE"
}

Write-Host "ProjectWise upload completed for $FileName"

if (-not $KeepDownloads) {
    Remove-Item -LiteralPath $LocalWorkDir -Recurse -Force -ErrorAction SilentlyContinue
}