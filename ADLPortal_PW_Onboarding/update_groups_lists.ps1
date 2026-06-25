Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Ensure-PwCommandsAvailable {
    [CmdletBinding()]
    param()

    if (Get-Command -Name 'New-PWLogin' -ErrorAction SilentlyContinue) {
        return
    }

    # PWPS_DAB module import is not fully compatible with StrictMode Latest.
    # Temporarily relax strict mode only for import, then restore it.
    Set-StrictMode -Off
    try {
        $moduleCandidates = @('pwps_dab', 'PWPS_DAB')
        foreach ($moduleName in $moduleCandidates) {
            try {
                Import-Module -Name $moduleName -ErrorAction Stop *> $null
                if (Get-Command -Name 'New-PWLogin' -ErrorAction SilentlyContinue) {
                    return
                }
            }
            catch {
            }
        }

        $modulePathCandidates = @(
            "$env:USERPROFILE\Documents\PowerShell\Modules\pwps_dab\pwps_dab.psd1",
            "$env:USERPROFILE\OneDrive - ARCADIS\Documents\PowerShell\Modules\pwps_dab\pwps_dab.psd1"
        )

        foreach ($modulePath in $modulePathCandidates) {
            try {
                if (Test-Path -Path $modulePath -PathType Leaf) {
                    Import-Module -Name $modulePath -ErrorAction Stop *> $null
                    if (Get-Command -Name 'New-PWLogin' -ErrorAction SilentlyContinue) {
                        return
                    }
                }
            }
            catch {
            }
        }
    }
    finally {
        Set-StrictMode -Version Latest
    }
}

function Ensure-PwRuntimePath {
    [CmdletBinding()]
    param()

    $candidates = @(
        'C:\Program Files\Bentley\ProjectWise\bin',
        'C:\Program Files (x86)\Bentley\ProjectWise\bin'
    )

    $pathItems = @($env:Path -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    foreach ($candidate in $candidates) {
        if (-not (Test-Path -Path $candidate -PathType Container)) {
            continue
        }

        if (-not (Test-Path -Path (Join-Path $candidate 'dmscli.dll') -PathType Leaf)) {
            continue
        }

        if ($pathItems -notcontains $candidate) {
            $env:Path = "$candidate;$env:Path"
            $pathItems = @($env:Path -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }
    }
}

function Get-FirstAvailablePwCommand {
    param(
        [string[]]$Candidates
    )

    foreach ($cmd in $Candidates) {
        if (Get-Command -Name $cmd -ErrorAction SilentlyContinue) {
            return $cmd
        }
    }

    return $null
}

function Get-NamesFromPwObjects {
    param(
        [object[]]$Items
    )

    if (-not $Items) {
        return @()
    }

    $names = @()
    foreach ($item in $Items) {
        if ($null -eq $item) {
            continue
        }

        $nameProp = $item.PSObject.Properties['Name']
        if ($nameProp -and -not [string]::IsNullOrWhiteSpace([string]$nameProp.Value)) {
            $names += [string]$nameProp.Value
            continue
        }

        $groupNameProp = $item.PSObject.Properties['GroupName']
        if ($groupNameProp -and -not [string]::IsNullOrWhiteSpace([string]$groupNameProp.Value)) {
            $names += [string]$groupNameProp.Value
            continue
        }

        $userListNameProp = $item.PSObject.Properties['UserListName']
        if ($userListNameProp -and -not [string]::IsNullOrWhiteSpace([string]$userListNameProp.Value)) {
            $names += [string]$userListNameProp.Value
            continue
        }

        $asString = [string]$item
        if (-not [string]::IsNullOrWhiteSpace($asString)) {
            $names += $asString
        }
    }

    return @($names | Sort-Object -Unique)
}

function Get-ProjectConfigFromDictionary {
    param(
        [string]$DictionaryPath,
        [string]$ProjectCode
    )

    if (-not (Test-Path -Path $DictionaryPath -PathType Leaf)) {
        throw "Dictionary file not found: $DictionaryPath"
    }

    $dict = (Get-Content -Path $DictionaryPath -Raw) | ConvertFrom-Json
    if (-not $dict.projects) {
        throw "Dictionary is invalid. Expected top-level 'projects' array."
    }

    $project = $dict.projects | Where-Object { $_.projectCode -eq $ProjectCode } | Select-Object -First 1
    if (-not $project) {
        throw "Project '$ProjectCode' was not found in dictionary."
    }

    return @{ Dictionary = $dict; Project = $project }
}

function Update-GroupsLists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Project,

        [Parameter(Mandatory = $false)]
        [string]$DictionaryPath = ".\projectwise-project-dictionary.json",

        # When set, the discovered groups and lists are also written to this JSON file
        # so they can be inspected before user creation runs.
        [Parameter(Mandatory = $false)]
        [string]$OutputFile = ".\pw-groups-lists-$($Project.ToUpper()).json"
    )

    $state = Get-ProjectConfigFromDictionary -DictionaryPath $DictionaryPath -ProjectCode $Project
    $dict = $state.Dictionary
    $projectConfig = $state.Project

    $datasource = $projectConfig.datasource
    $pwUserNameEnvVar = $projectConfig.credentials.pwUserNameEnvVar
    $pwPasswordEnvVar = $projectConfig.credentials.pwPasswordEnvVar

    if ([string]::IsNullOrWhiteSpace($pwUserNameEnvVar) -or [string]::IsNullOrWhiteSpace($pwPasswordEnvVar)) {
        throw "Dictionary credentials for project '$Project' are not configured."
    }

    $pwUserName = [Environment]::GetEnvironmentVariable($pwUserNameEnvVar)
    $pwPasswordPlain = [Environment]::GetEnvironmentVariable($pwPasswordEnvVar)
    if ([string]::IsNullOrWhiteSpace($pwUserName) -or [string]::IsNullOrWhiteSpace($pwPasswordPlain)) {
        throw "Environment variables '$pwUserNameEnvVar' and/or '$pwPasswordEnvVar' are not set."
    }

    $pwPassword = ConvertTo-SecureString -String $pwPasswordPlain -AsPlainText -Force

    Ensure-PwRuntimePath
    Ensure-PwCommandsAvailable

    $groupCandidates = @('Get-PWUserGroups', 'Get-PWGroups', 'Get-PWUserGroup', 'Get-PWGroupNames')
    $listCandidates = @('Get-PWUserLists', 'Get-PWLists', 'Get-PWUserList', 'Get-PWUserListNames')

    $groupCmd = Get-FirstAvailablePwCommand -Candidates $groupCandidates
    $listCmd = Get-FirstAvailablePwCommand -Candidates $listCandidates

    if (-not $groupCmd) {
        $availableGroupCommands = @(
            Get-Command -Name 'Get-PW*Group*' -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty Name -Unique
        )
        throw "No ProjectWise group query command found. Checked: $($groupCandidates -join ', '). Available Get-PW*Group* commands: $($availableGroupCommands -join ', ')"
    }

    if (-not $listCmd) {
        $availableListCommands = @(
            Get-Command -Name 'Get-PW*List*' -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty Name -Unique
        )
        throw "No ProjectWise list query command found. Checked: $($listCandidates -join ', '). Available Get-PW*List* commands: $($availableListCommands -join ', ')"
    }

    try {
        New-PWLogin -DatasourceName $datasource -UserName $pwUserName -Password $pwPassword -Verbose | Out-Null

        $groupItems = & $groupCmd
        $listItems = & $listCmd

        $groups = Get-NamesFromPwObjects -Items $groupItems
        $lists = Get-NamesFromPwObjects -Items $listItems

        $target = $dict.projects | Where-Object { $_.projectCode -eq $Project } | Select-Object -First 1
        $target.allowedUserGroups = @($groups)
        $target.allowedUserLists = @($lists)

        $dict | ConvertTo-Json -Depth 8 | Set-Content -Path $DictionaryPath -Encoding UTF8

        $result = [ordered]@{
            discoveredAt   = (Get-Date).ToUniversalTime().ToString('o')
            project        = $Project
            datasource     = $datasource
            groupCount     = $groups.Count
            listCount      = $lists.Count
            groups         = $groups
            lists          = $lists
            dictionaryPath = $DictionaryPath
        }

        if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
            $result | ConvertTo-Json -Depth 4 | Set-Content -Path $OutputFile -Encoding UTF8
        }

        return $result
    }
    finally {
        if (Get-Command -Name Undo-PWLogin -ErrorAction SilentlyContinue) {
            try {
                Undo-PWLogin
            }
            catch {
            }
        }
    }
}
