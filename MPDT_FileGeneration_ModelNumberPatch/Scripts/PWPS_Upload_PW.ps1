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
    [string]$AttributesJson = '',
    [string]$ProjectWiseBin = 'C:\Program Files\Bentley\ProjectWise\bin'
)

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

function Normalize-PWDocumentIdentity {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $name = (Split-Path -Leaf $Value).Trim()
    try {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($name)
    } catch {
        # Keep the original value if PowerShell cannot parse it as a path.
    }
    return $name.Trim().ToLowerInvariant()
}

function Test-PWDocumentIdentityMatch {
    param(
        $Doc,
        [string[]]$ExpectedNames
    )

    if (-not $Doc) { return $false }

    $expected = @{}
    foreach ($name in $ExpectedNames) {
        $normalized = Normalize-PWDocumentIdentity -Value $name
        if ($normalized) { $expected[$normalized] = $true }
    }

    foreach ($propName in @('Name', 'FileName', 'OldName', 'OldFileName')) {
        if ($Doc.PSObject.Properties.Name -contains $propName) {
            $normalized = Normalize-PWDocumentIdentity -Value ([string]$Doc.$propName)
            if ($normalized -and $expected.ContainsKey($normalized)) {
                return $true
            }
        }
    }

    return $false
}

function Get-UniquePWDocuments {
    param([object[]]$Documents)

    $seen = @{}
    $unique = @()
    foreach ($doc in @($Documents)) {
        if (-not $doc) { continue }
        $key = if ($doc.PSObject.Properties.Name -contains 'DocumentGUID' -and $doc.DocumentGUID) {
            [string]$doc.DocumentGUID
        } else {
            "$(Normalize-PWDocumentIdentity -Value ([string]$doc.Name))|$($doc.Version)"
        }
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $unique += $doc
        }
    }
    return $unique
}

function Find-PWDocumentForUpload {
    param(
        [string]$FolderPath,
        [string[]]$ExpectedNames,
        [string]$TargetVersion
    )

    $allCandidates = @()
    $searchCmd = Get-Command Get-PWDocumentsBySearch -ErrorAction Stop
    $supportsFileName = $searchCmd.Parameters.ContainsKey('FileName')
    $queries = @()

    foreach ($name in $ExpectedNames) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $leaf = (Split-Path -Leaf $name).Trim()
        $base = Normalize-PWDocumentIdentity -Value $leaf
        foreach ($q in @($leaf, $base, "*$base*")) {
            if (-not [string]::IsNullOrWhiteSpace($q) -and $queries -notcontains $q) {
                $queries += $q
            }
        }
    }

    foreach ($query in $queries) {
        try {
            if ($FolderPath) {
                $allCandidates += @(Get-PWDocumentsBySearch -DocumentName $query -FolderPath $FolderPath -ErrorAction SilentlyContinue)
            } else {
                $allCandidates += @(Get-PWDocumentsBySearch -DocumentName $query -ErrorAction SilentlyContinue)
            }
        } catch { }

        if ($supportsFileName) {
            try {
                if ($FolderPath) {
                    $allCandidates += @(Get-PWDocumentsBySearch -FileName $query -FolderPath $FolderPath -ErrorAction SilentlyContinue)
                } else {
                    $allCandidates += @(Get-PWDocumentsBySearch -FileName $query -ErrorAction SilentlyContinue)
                }
            } catch { }
        }
    }

    $exactMatches = @(Get-UniquePWDocuments -Documents $allCandidates | Where-Object {
        Test-PWDocumentIdentityMatch -Doc $_ -ExpectedNames $ExpectedNames
    })

    if (-not $exactMatches -or $exactMatches.Count -eq 0) {
        return $null
    }

    if (-not [string]::IsNullOrWhiteSpace($TargetVersion)) {
        $sameVersion = @($exactMatches | Where-Object { [string]$_.Version -eq $TargetVersion })
        if ($sameVersion.Count -gt 0) {
            return ($sameVersion | Sort-Object VersionSequence -Descending | Select-Object -First 1)
        }
    }

    return ($exactMatches | Sort-Object VersionSequence -Descending | Select-Object -First 1)
}

function Refresh-PWDocumentByGuid {
    param($Doc)

    if ($Doc -and $Doc.DocumentGUID -and (Get-Command Get-PWDocumentsByGUIDs -ErrorAction SilentlyContinue)) {
        $refreshed = Get-PWDocumentsByGUIDs -DocumentGUIDs @([string]$Doc.DocumentGUID) -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($refreshed) { return $refreshed }
    }
    return $Doc
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

    $resolvedProjectWiseBin = (Resolve-Path -LiteralPath $ProjectWiseBin -ErrorAction Stop).ProviderPath
    $ProjectWiseBin = $resolvedProjectWiseBin
    $pathParts = @($env:Path -split ';' | Where-Object { $_ })
    if ($pathParts -notcontains $ProjectWiseBin) {
        $env:Path = "$ProjectWiseBin;$env:Path"
    }

    $nativeSource = @'
using System;
using System.Runtime.InteropServices;
public static class Kernel32 {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool SetDllDirectory(string lpPathName);

    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr LoadLibrary(string lpFileName);
}
'@

    if (-not ('Kernel32' -as [type])) {
        Add-Type -TypeDefinition $nativeSource -Language CSharp -ErrorAction Stop
    }

    [Kernel32]::SetDllDirectory($ProjectWiseBin) | Out-Null
    $handle = [Kernel32]::LoadLibrary($dllPath)
    if ($handle -eq [IntPtr]::Zero) {
        $win32Error = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "LoadLibrary failed for '$dllPath' with Win32 error $win32Error"
    }
}

try {
    Test-PWNativeRuntime -ProjectWiseBin $ProjectWiseBin
} catch {
    Write-Error "ProjectWise native runtime check failed: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Import module
# ---------------------------------------------------------------------------
try {
    # pwps_dab can emit non-terminating import-time errors on some ProjectWise
    # installs, especially under strict/Stop preference. Import it softly, then
    # explicitly verify the module is available before continuing.
    $PreviousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    Import-Module pwps_dab -ErrorAction Continue -WarningAction SilentlyContinue
    $ErrorActionPreference = $PreviousErrorActionPreference

    if (-not (Get-Module -Name pwps_dab)) {
        throw 'pwps_dab did not load into the current PowerShell session.'
    }
} catch {
    $ErrorActionPreference = 'Stop'
    Write-Error "Failed to import pwps_dab: $_. Install with: Install-Module pwps_dab -Scope CurrentUser"
    exit 1
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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


    # Check if document already exists (version/update) or create new.
    $ExpectedNames = @($docNameToUse, $FileName, [System.IO.Path]::GetFileNameWithoutExtension($FileName))
    $ExistingDoc = Find-PWDocumentForUpload -FolderPath $NormalizedFolderPath -ExpectedNames $ExpectedNames -TargetVersion $Version

    if ($ExistingDoc) {
        Write-Host "  Existing document found: $($ExistingDoc.Name) version $($ExistingDoc.Version)"
        $TargetDoc = $ExistingDoc

        if (-not [string]::IsNullOrWhiteSpace($Version) -and [string]$ExistingDoc.Version -eq $Version) {
            throw "Target ProjectWise version '$Version' already exists for '$docNameToUse'. Refusing to replace an existing revision."
        }

        if (-not [string]::IsNullOrWhiteSpace($Version) -and [string]$ExistingDoc.Version -ne $Version) {
            $newVersionCmd = Get-Command New-PWDocumentVersion -ErrorAction SilentlyContinue
            if (-not $newVersionCmd) {
                throw "Existing document version is $($ExistingDoc.Version), target version is $Version, but New-PWDocumentVersion is not available. Refusing to overwrite the existing version."
            }

            Write-Host "  Creating ProjectWise document version: $($ExistingDoc.Version) -> $Version"
            $versionedDocs = @(New-PWDocumentVersion -InputDocument @($ExistingDoc) -VersionString $Version -ErrorAction Stop)
            if ($versionedDocs.Count -gt 0) {
                $TargetDoc = $versionedDocs | Select-Object -First 1
            } else {
                $TargetDoc = Find-PWDocumentForUpload -FolderPath $NormalizedFolderPath -ExpectedNames $ExpectedNames -TargetVersion $Version
            }

            if (-not $TargetDoc) {
                throw "Created ProjectWise version $Version, but could not resolve the new version document."
            }
        }

        $TargetDoc = Refresh-PWDocumentByGuid -Doc $TargetDoc
        Write-Host "  Uploading file to document: $($TargetDoc.Name) version $($TargetDoc.Version)"
        if (Get-Command Send-PWFile -ErrorAction SilentlyContinue) {
            Send-PWFile -DocumentGUID $TargetDoc.DocumentGUID -LocalFileName $FilePath
        } elseif (Get-Command Update-PWDocumentFile -ErrorAction SilentlyContinue) {
            Update-PWDocumentFile -InputDocuments $TargetDoc -NewFilePathName $FilePath -KeepExistingFileName
        } else {
            throw 'No supported update cmdlet found (expected Send-PWFile or Update-PWDocumentFile).'
        }

        # Apply editable metadata after file update.
        $TargetDoc = Refresh-PWDocumentByGuid -Doc $TargetDoc
        $TargetDoc.Name = $docNameToUse
        $TargetDoc.Description = $descToUse
        if (-not [string]::IsNullOrWhiteSpace($Version)) {
            $TargetDoc.Version = $Version
        }
        Update-PWDocumentProperties -InputDocument @($TargetDoc) | Out-Null

        if (-not [string]::IsNullOrWhiteSpace($Application) -and (Get-Command Set-PWDocumentApplication -ErrorAction SilentlyContinue)) {
            Set-PWDocumentApplication -InputDocuments @($TargetDoc) -Application $Application | Out-Null
        }

        if ($AttributesMap -and (Get-Command Update-PWDocumentAttributes -ErrorAction SilentlyContinue)) {
            Update-PWDocumentAttributes -InputDocuments @($TargetDoc) -Attributes $AttributesMap | Out-Null
        }

        $refreshed = Refresh-PWDocumentByGuid -Doc $TargetDoc
        if (-not [string]::IsNullOrWhiteSpace($Version) -and [string]$refreshed.Version -ne $Version) {
            throw "Upload completed but ProjectWise version is '$($refreshed.Version)' instead of expected '$Version'."
        }
        Write-DocumentSummary -Doc $refreshed
        Write-Host "  Updated OK."
    } else {
        throw "No existing ProjectWise document found in '$NormalizedFolderPath' for '$docNameToUse'. Refusing to create a new/duplicate document."
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


