<#
.SYNOPSIS
    Upload a generated ACBOS or MPDT file to ProjectWise.

.DESCRIPTION
    Connects to ProjectWise using pwps_dab and uploads a single file,
    creating a new document or adding a new version to an existing one.

.PARAMETER FilePath
    Full path to the local file to upload (.xlsm or .ACBOS).

.PARAMETER DatasourceName
    ProjectWise datasource name.

.PARAMETER UserName
    ProjectWise username.

.PARAMETER Password
    Optional ProjectWise password. If omitted, the script prompts securely.

.PARAMETER PWFolderPath
    Target folder path in ProjectWise where the file will be uploaded.
    Defaults to the configured ACBOS/MPDT output folder.

.EXAMPLE
    pwsh -File PWPS_Upload_PW.ps1 -FilePath "C:\Output\HS2-000001234.xlsm" `
         -DatasourceName "arcadis-uk-pw.bentley.com:arcadis-uk-07" `
         -UserName "_asc_user_automation"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FilePath,

    [string]$DatasourceName = 'arcadis-uk-pw.bentley.com:arcadis-uk-07',
    [string]$UserName       = '_asc_user_automation',
    [string]$Password       = '',
    [string]$PWFolderPath   = '',  # Set to the target PW folder path
    [string]$DocumentName   = '',
    [string]$Description    = '',
    [string]$Version        = '',
    [string]$Application    = '',
    [string]$AttributesJson = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $FilePath)) {
    Write-Error "File not found: $FilePath"
    exit 1
}

$FileName = Split-Path -Leaf $FilePath
Write-Host "Uploading: $FileName"

function Normalize-PWFolderPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ''
    }
    $p = $Path.Trim()
    $p = $p.Trim('"')
    $p = $p.TrimStart('\\')
    $p = $p.TrimEnd('\\')
    return $p
}

function Convert-AttributesJsonToHashtable {
    param([string]$Json)

    if ([string]::IsNullOrWhiteSpace($Json)) {
        return $null
    }

    $raw = $Json
    if (Test-Path $Json) {
        $raw = Get-Content -Path $Json -Raw -ErrorAction Stop
    }

    $obj = ConvertFrom-Json -InputObject $raw -ErrorAction Stop
    $hash = @{}
    foreach ($p in $obj.PSObject.Properties) {
        $hash[$p.Name] = if ($null -eq $p.Value) { '' } else { [string]$p.Value }
    }
    return $hash
}

function Write-DocumentSummary {
    param($Doc)

    if (-not $Doc) { return }

    $summary = [ordered]@{
        DocumentGUID = $Doc.DocumentGUID
        ProjectGUID = $Doc.ProjectGUID
        ProjectGUIDString = $Doc.ProjectGUIDString
        ProjectID = $Doc.ProjectID
        DocumentID = $Doc.DocumentID
        Name = $Doc.Name
        Description = $Doc.Description
        FileName = $Doc.FileName
        Version = $Doc.Version
        VersionSequence = $Doc.VersionSequence
        OriginalNumber = $Doc.OriginalNumber
        FileSize = $Doc.FileSize
        DocumentStatus = $Doc.DocumentStatus
        DocumentOutTo = $Doc.DocumentOutTo
        DocumentOutToName = $Doc.DocumentOutToName
        DocumentCheckOutDate = $Doc.DocumentCheckOutDate
        Workflow = $Doc.Workflow
        WorkflowState = $Doc.WorkflowState
        DocumentUpdater = $Doc.DocumentUpdater
        StatusChangeDate = $Doc.StatusChangeDate
        DocumentUpdaterName = $Doc.DocumentUpdaterName
        DocumentUpdateDate = $Doc.DocumentUpdateDate
        FileUpdater = $Doc.FileUpdater
        FileUpdaterName = $Doc.FileUpdaterName
        FileUpdateDate = $Doc.FileUpdateDate
        DocumentCreator = $Doc.DocumentCreator
        DocumentCreatorName = $Doc.DocumentCreatorName
        CreateDate = $Doc.CreateDate
        DocumentGUIDString = $Doc.DocumentGUIDString
        DocumentURN = $Doc.DocumentURN
        FullPath = $Doc.FullPath
        FolderPath = $Doc.FolderPath
        OldVersion = $Doc.OldVersion
        WorkflowId = $Doc.WorkflowId
        StateId = $Doc.StateId
        ApplicationId = $Doc.ApplicationId
        ApplicationName = $Doc.ApplicationName
        CheckedOutLocalFileName = $Doc.CheckedOutLocalFileName
        CopiedOutLocalFileName = $Doc.CopiedOutLocalFileName
        IsAbstract = $Doc.IsAbstract
        IsSet = $Doc.IsSet
        DocumentOwnerType = $Doc.DocumentOwnerType
        DocumentOwnerName = $Doc.DocumentOwnerName
        FullPathToStorageAreaFile = $Doc.FullPathToStorageAreaFile
        Status = $Doc.Status
        MIMEType = $Doc.MIMEType
        StorageName = $Doc.StorageName
        StorageId = $Doc.StorageId
        ProjectWiseWebViewLink = $Doc.ProjectWiseWebViewLink
        ProjectWiseWebLink = $Doc.ProjectWiseWebLink
        WebServicesGatewayDownloadLink = $Doc.WebServicesGatewayDownloadLink
        OldName = $Doc.OldName
        OldFileName = $Doc.OldFileName
        OldDescription = $Doc.OldDescription
        OldFullPath = $Doc.OldFullPath
        AttributeRecordCount = $Doc.AttributeRecordCount
        Attributes = $Doc.Attributes
        NestedReferenceCount = $Doc.NestedReferenceCount
        CustomAttributes = $Doc.CustomAttributes
    }

    [pscustomobject]$summary | Format-List
}

function Test-PWNativeRuntime {
    param(
        [string]$ProjectWiseBin = 'C:\Program Files\Bentley\ProjectWise\bin'
    )

    $dllPath = Join-Path $ProjectWiseBin 'dmscli.dll'

    if (-not (Test-Path $dllPath)) {
        throw "dmscli.dll not found at: $dllPath"
    }

    if (-not [Environment]::Is64BitProcess) {
        throw 'pwps_dab requires a 64-bit PowerShell process for this ProjectWise installation.'
    }

    $nativeSource = @'
using System;
using System.Runtime.InteropServices;
public static class Kernel32 {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr LoadLibrary(string lpFileName);
}
'@

    if (-not ('Kernel32' -as [type])) {
        Add-Type -TypeDefinition $nativeSource -Language CSharp -ErrorAction Stop
    }

    $handle = [Kernel32]::LoadLibrary($dllPath)
    if ($handle -eq [IntPtr]::Zero) {
        $win32Error = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "LoadLibrary failed for '$dllPath' with Win32 error $win32Error"
    }
}

# ---------------------------------------------------------------------------
# Import module
# ---------------------------------------------------------------------------
try {
    Import-Module pwps_dab -ErrorAction Stop
} catch {
    Write-Error "Failed to import pwps_dab: $_. Install with: Install-Module pwps_dab -Scope CurrentUser"
    exit 1
}

try {
    Test-PWNativeRuntime
} catch {
    Write-Error "ProjectWise native runtime check failed: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Connect
# ---------------------------------------------------------------------------
try {
    if ([string]::IsNullOrWhiteSpace($Password)) {
        $PWPassword = Read-Host -Prompt "ProjectWise password for $UserName" -AsSecureString
    } else {
        $PWPassword = ConvertTo-SecureString $Password -AsPlainText -Force
    }
    New-PWLogin -DatasourceName $DatasourceName -UserName $UserName -Password $PWPassword
    Write-Host "Connected to ProjectWise."
} catch {
    Write-Error "PW login failed: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Upload file
# ---------------------------------------------------------------------------
try {
    $ResolvedFolder = $null
    $NormalizedFolderPath = Normalize-PWFolderPath -Path $PWFolderPath
    $AttributesMap = Convert-AttributesJsonToHashtable -Json $AttributesJson
    $docNameToUse = if ([string]::IsNullOrWhiteSpace($DocumentName)) { $FileName } else { $DocumentName }
    $descToUse = if ([string]::IsNullOrWhiteSpace($Description)) { $FileName } else { $Description }

    if ($NormalizedFolderPath) {
        $ResolvedFolder = Get-PWFolders -FolderPath $NormalizedFolderPath -JustOne -PopulatePaths -ErrorAction Stop
        if (-not $ResolvedFolder) {
            throw "Target PW folder not found: '$PWFolderPath'"
        }
        Write-Host "Resolved target folder: $($ResolvedFolder.FullPath)"
    }

    if ($PWFolderPath) {
        # New-PWDocument in pwps_dab uses FilePath + FolderPath
        $UploadParams = @{
            FilePath   = $FilePath
            FolderPath = $NormalizedFolderPath
        }
    } else {
        $UploadParams = @{
            FilePath = $FilePath
        }
    }

    # Check if document already exists (update) or create new
    $ExistingDoc = $null
    try {
        # Use a wildcard search to increase chance of matching existing document names
        $basename = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
        $searchName = "*${basename}*"
        if ($NormalizedFolderPath) {
            $ExistingDoc = Get-PWDocumentsBySearch -DocumentName $searchName -FolderPath $NormalizedFolderPath |
                           Select-Object -First 1
        } else {
            $ExistingDoc = Get-PWDocumentsBySearch -DocumentName $searchName |
                           Select-Object -First 1
        }
    } catch { <# not found or search error #> }

    if ($ExistingDoc) {
        Write-Host "  Updating existing document: $($ExistingDoc.Name)"
        if (Get-Command Send-PWFile -ErrorAction SilentlyContinue) {
            Send-PWFile -DocumentGUID $ExistingDoc.DocumentGUID -LocalFileName $FilePath
        } elseif (Get-Command Update-PWDocumentFile -ErrorAction SilentlyContinue) {
            Update-PWDocumentFile -InputDocuments $ExistingDoc -NewFilePathName $FilePath -KeepExistingFileName
        } else {
            throw 'No supported update cmdlet found (expected Send-PWFile or Update-PWDocumentFile).'
        }

        # Apply editable metadata after file update.
        $ExistingDoc.Name = $docNameToUse
        $ExistingDoc.Description = $descToUse
        if (-not [string]::IsNullOrWhiteSpace($Version)) {
            $ExistingDoc.Version = $Version
        }
        Update-PWDocumentProperties -InputDocument @($ExistingDoc) | Out-Null

        if (-not [string]::IsNullOrWhiteSpace($Application) -and (Get-Command Set-PWDocumentApplication -ErrorAction SilentlyContinue)) {
            Set-PWDocumentApplication -InputDocuments @($ExistingDoc) -Application $Application | Out-Null
        }

        if ($AttributesMap -and (Get-Command Update-PWDocumentAttributes -ErrorAction SilentlyContinue)) {
            Update-PWDocumentAttributes -InputDocuments @($ExistingDoc) -Attributes $AttributesMap | Out-Null
        }

        if ($ExistingDoc.DocumentGUID -and (Get-Command Get-PWDocumentsByGUIDs -ErrorAction SilentlyContinue)) {
            $refreshed = Get-PWDocumentsByGUIDs -DocumentGUIDs @([string]$ExistingDoc.DocumentGUID) -ErrorAction SilentlyContinue | Select-Object -First 1
            Write-DocumentSummary -Doc $refreshed
        }
        Write-Host "  Updated OK."
    } else {
        Write-Host "  Creating new document in PW."
        try {
            if (-not $UploadParams.ContainsKey('FolderPath') -and -not $UploadParams.ContainsKey('FilePath')) {
                throw 'Internal error: upload parameters not populated.'
            }

            # Try the creation call that matches the installed pwps_dab signature.
            # If that signature is not available, fail immediately instead of prompting interactively.
            $createdDoc = $null
            if ($ResolvedFolder) {
                $createdDoc = $ResolvedFolder | New-PWDocument -FilePath $FilePath -DocumentName $docNameToUse -Description $descToUse -Version $Version -Application $Application -ErrorAction Stop
            } else {
                $createdDoc = New-PWDocument @UploadParams -DocumentName $docNameToUse -Description $descToUse -Version $Version -Application $Application -ErrorAction Stop
            }

            if ($createdDoc -and $createdDoc.DocumentGUID -and (Get-Command Get-PWDocumentsByGUIDs -ErrorAction SilentlyContinue)) {
                $verified = Get-PWDocumentsByGUIDs -DocumentGUIDs @([string]$createdDoc.DocumentGUID) -ErrorAction SilentlyContinue
                if ($verified) {
                    $v = $verified | Select-Object -First 1
                    Write-Host "  Created GUID: $($v.DocumentGUID)"
                    Write-Host "  Created Name: $($v.Name)"
                    Write-Host "  Created FolderPath: $($v.FolderPath)"
                    Write-Host "  Created FileSize: $($v.FileSize)"

                    if ($AttributesMap -and (Get-Command Update-PWDocumentAttributes -ErrorAction SilentlyContinue)) {
                        Update-PWDocumentAttributes -InputDocuments @($v) -Attributes $AttributesMap | Out-Null
                        $v = Get-PWDocumentsByGUIDs -DocumentGUIDs @([string]$v.DocumentGUID) -ErrorAction SilentlyContinue | Select-Object -First 1
                    }

                    Write-DocumentSummary -Doc $v
                }
            }
            Write-Host "  Created OK."
        } catch {
            Write-Error "Upload failed creating a new document for $FileName : $_"
            exit 1
        }
    }
} catch {
    Write-Error "Upload failed for $FileName : $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Disconnect
# ---------------------------------------------------------------------------
try { Remove-PWLogin } catch { <# ignore #> }

Write-Host "Upload complete: $FileName"

