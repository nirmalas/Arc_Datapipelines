# HOW TO RUN
# -----------
# Step 1: Open a PowerShell terminal and navigate to the script folder:
#   cd "C:\path\to\auk-pmo-prod-automationScripts\python_scripts\ADL\ADLPortal_PW_Onboarding"
#
# Step 2: Set credentials (only needed for local testing; Jenkins injects these automatically):
#   $env:PW_WWHD_USER = "_ADL_Automation"
#   $env:PW_WWHD_PASSWORD = "your-actual-password"
#
# Step 3: Run the script using one of the examples below.
#
# Dry-run only (validate without touching ProjectWise):
#   pwsh -NoProfile -ExecutionPolicy Bypass -File .\Invoke-PWUserOnboarding.ps1 `
#     -Project WWHD -UserName first.last -Email first.last@arcadis.com -DryRun
#
# Dry-run with automatic group/list refresh from ProjectWise:
#   pwsh -NoProfile -ExecutionPolicy Bypass -File .\Invoke-PWUserOnboarding.ps1 `
#     -Project WWHD -UserName first.last -Email first.last@arcadis.com -UpdateGroupsLists -DryRun
#
# Create user for real (remove -DryRun):
#   pwsh -NoProfile -ExecutionPolicy Bypass -File .\Invoke-PWUserOnboarding.ps1 `
#     -Project WWHD -UserName first.last -Email first.last@arcadis.com `
#     -UserGroups "WWHD-Admin,WWHD-Design" -UserLists "WWHD-Team"
#
# NOTE: Always run from inside the ADLPortal_PW_Onboarding folder so that
#       .\projectwise-project-dictionary.json and .\update_groups_lists.ps1 are found.
# -----------


param(
    [Parameter(Mandatory = $true)]
    [string]$Project,

    [Parameter(Mandatory = $true)]
    [string]$UserName,

    [Parameter(Mandatory = $true)]
    [string]$Email,

    [Parameter(Mandatory = $false)]
    [string]$Description = "",

    [Parameter(Mandatory = $false)]
    [string]$Password = "",

    [Parameter(Mandatory = $false)]
    [bool]$IMSUser = $true,

    [Parameter(Mandatory = $false)]
    [string[]]$UserGroups = @(),

    [Parameter(Mandatory = $false)]
    [string[]]$UserLists = @(),

    [Parameter(Mandatory = $false)]
    [string]$SecurityProvider = "",

    [Parameter(Mandatory = $false)]
    [string]$Initial = "",

    [Parameter(Mandatory = $false)]
    [string]$TBName = "",

    [Parameter(Mandatory = $false)]
    [string]$OriginCode = "",

    [Parameter(Mandatory = $false)]
    [string]$Originator = "",

    [Parameter(Mandatory = $false)]
    [string]$Discipline = "",

    [Parameter(Mandatory = $false)]
    [string]$Grade = "",

    [Parameter(Mandatory = $false)]
    [string]$DisciplineCode = "",

    [Parameter(Mandatory = $false)]
    [string]$GradeLevel = "",

    [Parameter(Mandatory = $false)]
    [string]$DictionaryPath = ".\projectwise-project-dictionary.json",

    [Parameter(Mandatory = $false)]
    [string]$CallbackUrl = "",

    [Parameter(Mandatory = $false)]
    [string]$CallbackToken = "",

    [Parameter(Mandatory = $false)]
    [string]$CorrelationId = "",

    [Parameter(Mandatory = $false)]
    [switch]$UpdateGroupsLists,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"



function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$ts] [$Level] $Message"
}

function Post-Status {
    param(
        [hashtable]$Payload,
        [string]$TargetUrl,
        [string]$Token
    )

    if ([string]::IsNullOrWhiteSpace($TargetUrl)) {
        Write-Log "CallbackUrl was not provided. Skipping callback POST." "WARN"
        return
    }

    try {
        $headers = @{
            "Content-Type" = "application/json"
        }
        if (-not [string]::IsNullOrWhiteSpace($Token)) {
            $headers["Authorization"] = "Bearer $Token"
        }

        $body = $Payload | ConvertTo-Json -Depth 8
        Invoke-RestMethod -Uri $TargetUrl -Method Post -Headers $headers -Body $body | Out-Null
        Write-Log "Status callback posted successfully to $TargetUrl"
    }
    catch {
        Write-Log "Failed to post callback: $($_.Exception.Message)" "WARN"
    }
}

function Get-ProjectConfig {
    param(
        [string]$ConfigPath,
        [string]$ProjectCode
    )

    if (-not (Test-Path -Path $ConfigPath -PathType Leaf)) {
        throw "Dictionary file not found: $ConfigPath"
    }

    $raw = Get-Content -Path $ConfigPath -Raw
    $dict = $raw | ConvertFrom-Json

    if (-not $dict.projects) {
        throw "Dictionary is invalid. Expected top-level 'projects' array."
    }

    $match = $dict.projects | Where-Object { $_.projectCode -eq $ProjectCode } | Select-Object -First 1
    if (-not $match) {
        throw "Project '$ProjectCode' was not found in dictionary."
    }

    return $match
}

function Get-EnvValueOrThrow {
    param(
        [string]$EnvName,
        [string]$FriendlyName
    )

    if ([string]::IsNullOrWhiteSpace($EnvName)) {
        throw "$FriendlyName environment variable name is not configured in dictionary."
    }

    $value = [Environment]::GetEnvironmentVariable($EnvName)
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "$FriendlyName environment variable '$EnvName' is empty or not set."
    }

    return $value
}

function Ensure-PwRuntimePath {
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

function Test-ObjectHasProperty {
    param(
        [object]$Obj,
        [string]$PropertyName
    )

    if ($null -eq $Obj -or [string]::IsNullOrWhiteSpace($PropertyName)) {
        return $false
    }

    return ($null -ne $Obj.PSObject -and $null -ne $Obj.PSObject.Properties[$PropertyName])
}

function Get-AllowedMemberships {
    param(
        [object[]]$Requested,
        [object[]]$Allowed,
        [string]$MembershipType
    )

    if (-not $Requested) {
        return @()
    }

    $requestedClean = @($Requested | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })
    $allowedClean = @($Allowed | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })

    foreach ($item in $requestedClean) {
        if ($allowedClean -notcontains $item) {
            throw "$MembershipType '$item' is not allowed for this project."
        }
    }

    return $requestedClean
}

function Add-UserToGroups {
    param(
        [string]$TargetUserName,
        [string[]]$Groups
    )

    $added = @()
    foreach ($g in $Groups) {
        Add-PWUserToGroup -GroupName $g -UserName $TargetUserName -Verbose
        $added += $g
    }
    return $added
}

function Add-UserToLists {
    param(
        [string]$TargetUserName,
        [string[]]$Lists
    )

    if (-not $Lists -or $Lists.Count -eq 0) {
        return @()
    }

    $candidateCommands = @(
        "Add-PWUserToList",
        "Add-PWUserToUserList",
        "Add-PWUserListMember"
    )

    $commandName = $null
    foreach ($cmd in $candidateCommands) {
        if (Get-Command -Name $cmd -ErrorAction SilentlyContinue) {
            $commandName = $cmd
            break
        }
    }

    if (-not $commandName) {
        throw "No ProjectWise user-list command found. Checked: $($candidateCommands -join ', ')"
    }

    $added = @()
    foreach ($lst in $Lists) {
        switch ($commandName) {
            "Add-PWUserToList" {
                Add-PWUserToList -ListName $lst -UserName $TargetUserName -Verbose
            }
            "Add-PWUserToUserList" {
                Add-PWUserToUserList -UserListName $lst -UserName $TargetUserName -Verbose
            }
            "Add-PWUserListMember" {
                Add-PWUserListMember -UserListName $lst -UserName $TargetUserName -Verbose
            }
        }
        $added += $lst
    }

    return $added
}

function Build-ResultPayload {
    param(
        [string]$Result,
        [string]$Message,
        [string]$ProjectCode,
        [string]$RepositoryId,
        [string]$TargetUserName,
        [string]$TargetEmail,
        [bool]$Created,
        [string[]]$GroupsAdded,
        [string[]]$ListsAdded,
        [string]$Cid
    )

    return @{
        timestampUtc  = (Get-Date).ToUniversalTime().ToString("o")
        correlationId = $Cid
        result        = $Result
        message       = $Message
        project       = $ProjectCode
        repositoryId  = $RepositoryId
        user = @{
            userName = $TargetUserName
            email    = $TargetEmail
            created  = $Created
            groups   = $GroupsAdded
            lists    = $ListsAdded
        }
    }
}

$resultPayload = $null

try {
    . ".\update_groups_lists.ps1"

    if ($UpdateGroupsLists.IsPresent) {
        Write-Log "Optional pre-step enabled: updating allowed groups and user lists for project '$Project'"
        $refreshResult = Update-GroupsLists -Project $Project -DictionaryPath $DictionaryPath

        $refreshObject = $refreshResult
        if ($refreshResult -is [System.Array]) {
            $refreshObject = $refreshResult |
                Where-Object {
                    ($_ -is [System.Collections.IDictionary]) -or
                    (Test-ObjectHasProperty -Obj $_ -PropertyName 'groupCount')
                } |
                Select-Object -Last 1
        }

        $groupCount = $null
        $listCount = $null
        if ($refreshObject -is [System.Collections.IDictionary]) {
            $groupCount = $refreshObject['groupCount']
            $listCount = $refreshObject['listCount']
        }
        elseif (Test-ObjectHasProperty -Obj $refreshObject -PropertyName 'groupCount') {
            $groupCount = $refreshObject.groupCount
            $listCount = $refreshObject.listCount
        }
        else {
            throw "Update-GroupsLists completed but did not return group/list counts."
        }

        Write-Log "Dictionary updated for '$Project'. Groups: $groupCount, Lists: $listCount"
    }

    Write-Log "Loading project dictionary from '$DictionaryPath'"
    $projectConfig = Get-ProjectConfig -ConfigPath $DictionaryPath -ProjectCode $Project

    $repoId = $projectConfig.repositoryId
    $datasource = $projectConfig.datasource
    $pwUserNameEnvVar = $projectConfig.credentials.pwUserNameEnvVar
    $pwPasswordEnvVar = $projectConfig.credentials.pwPasswordEnvVar

    if ([string]::IsNullOrWhiteSpace($CorrelationId)) {
        $CorrelationId = [guid]::NewGuid().ToString()
    }

    Write-Log "Validating requested groups and user lists for project '$Project'"
    $validatedGroups = Get-AllowedMemberships -Requested $UserGroups -Allowed $projectConfig.allowedUserGroups -MembershipType "Group"
    $validatedLists = Get-AllowedMemberships -Requested $UserLists -Allowed $projectConfig.allowedUserLists -MembershipType "User list"

    $groupCount = @($validatedGroups).Count
    $listCount = @($validatedLists).Count
    Write-Log "Validated groups: $groupCount group(s), lists: $listCount list(s)"

    if ($DryRun.IsPresent) {
        Write-Log "Dry run enabled. No ProjectWise changes will be performed." "WARN"
        Write-Log "Dry run summary: Project=$Project, User=$UserName, Email=$Email, Groups=$groupCount, Lists=$listCount"
        $resultPayload = Build-ResultPayload -Result "success" -Message "Dry run successful. Validation passed." -ProjectCode $Project -RepositoryId $repoId -TargetUserName $UserName -TargetEmail $Email -Created $false -GroupsAdded $validatedGroups -ListsAdded $validatedLists -Cid $CorrelationId
        Post-Status -Payload $resultPayload -TargetUrl $CallbackUrl -Token $CallbackToken
        $resultPayload | ConvertTo-Json -Depth 8
        exit 0
    }

    $pwUserName = Get-EnvValueOrThrow -EnvName $pwUserNameEnvVar -FriendlyName "ProjectWise username"
    $pwPasswordPlain = Get-EnvValueOrThrow -EnvName $pwPasswordEnvVar -FriendlyName "ProjectWise password"
    $pwPassword = ConvertTo-SecureString -String $pwPasswordPlain -AsPlainText -Force


    Ensure-PwRuntimePath

    Write-Log "Connecting to ProjectWise datasource '$datasource'"
    Ensure-PwCommandsAvailable
    New-PWLogin -DatasourceName $datasource -UserName $pwUserName -Password $pwPassword -Verbose
    Write-Log "Successfully connected to ProjectWise datasource '$datasource'"

    $created = $false
    $groupsAdded = @()
    $listsAdded = @()

    Write-Log "Checking if user '$UserName' already exists in repository"
    $userExists = Get-PWUsersByMatch -UserName $UserName -ErrorAction SilentlyContinue
    if ($userExists) {
        $userDetails = "Disabled: $($userExists.IsDisabled)"
        if ($userExists.PSObject.Properties['Email']) {
            $userDetails += ", Email: $($userExists.Email)"
        }
        if ($userExists.PSObject.Properties['FullName']) {
            $userDetails += ", FullName: $($userExists.FullName)"
        }
        Write-Log "User '$UserName' already exists in repository. Skipping user creation." "WARN"
        Write-Log "User details - $userDetails"
    }
    else {
        Write-Log "User '$UserName' does not exist. Creating new user in ProjectWise"

        if ($IMSUser) {
            Write-Log "Creating IMS user: $UserName with email: $Email"
            New-PWUserSimple -UserNames $UserName -Description $Description -Email $Email -IMSUser -Verbose
        }
        else {
            if ([string]::IsNullOrWhiteSpace($Password)) {
                throw "Password must be provided when IMSUser is false."
            }
            Write-Log "Creating non-IMS user: $UserName with email: $Email"
            New-PWUserSimple -UserNames $UserName -Description $Description -Email $Email -Password $Password -Verbose
        }

        $created = $true
        Write-Log "User '$UserName' created successfully in ProjectWise"
    }

    if ($groupCount -gt 0) {
        Write-Log "Adding user '$UserName' to $groupCount group(s): $($validatedGroups -join ', ')"
        $groupsAdded = Add-UserToGroups -TargetUserName $UserName -Groups $validatedGroups
        Write-Log "Successfully added user '$UserName' to groups. Count: $(@($groupsAdded).Count)"
    }
    else {
        Write-Log "No groups specified for user '$UserName'. Skipping group assignment."
    }

    if ($listCount -gt 0) {
        Write-Log "Adding user '$UserName' to $listCount user list(s): $($validatedLists -join ', ')"
        $listsAdded = Add-UserToLists -TargetUserName $UserName -Lists $validatedLists
        Write-Log "Successfully added user '$UserName' to lists. Count: $(@($listsAdded).Count)"
    }
    else {
        Write-Log "No user lists specified for user '$UserName'. Skipping list assignment."
    }

    Write-Log "Verifying user '$UserName' in repository after processing"
    $pwFinalUser = Get-PWUsersByMatch -UserName $UserName -ErrorAction SilentlyContinue
    if (-not $pwFinalUser) {
        throw "Verification failed. User '$UserName' was not found after create/update workflow."
    }
    Write-Log "User '$UserName' verification successful. Disabled: $($pwFinalUser.IsDisabled), Email: $($pwFinalUser.Email)"

    $successMessage = if ($created) { "User created successfully and processed." } else { "User already existed and has been processed." }
    Write-Log "Workflow completed successfully. Message: $successMessage. Groups added: $($groupsAdded.Count), Lists added: $($listsAdded.Count)"

    $resultPayload = Build-ResultPayload -Result "success" -Message $successMessage -ProjectCode $Project -RepositoryId $repoId -TargetUserName $UserName -TargetEmail $Email -Created $created -GroupsAdded $groupsAdded -ListsAdded $listsAdded -Cid $CorrelationId
    Post-Status -Payload $resultPayload -TargetUrl $CallbackUrl -Token $CallbackToken

    $resultPayload | ConvertTo-Json -Depth 8
    exit 0
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Log "Workflow failed: $errorMessage" "ERROR"

    if (-not $CorrelationId) {
        $CorrelationId = [guid]::NewGuid().ToString()
    }

    $safeRepoId = ""
    try {
        if ($projectConfig -and $projectConfig.repositoryId) {
            $safeRepoId = $projectConfig.repositoryId
        }
    }
    catch {
    }

    $resultPayload = Build-ResultPayload -Result "error" -Message $errorMessage -ProjectCode $Project -RepositoryId $safeRepoId -TargetUserName $UserName -TargetEmail $Email -Created $false -GroupsAdded @() -ListsAdded @() -Cid $CorrelationId
    Post-Status -Payload $resultPayload -TargetUrl $CallbackUrl -Token $CallbackToken

    $resultPayload | ConvertTo-Json -Depth 8
    exit 1
}
finally {
    if (Get-Command -Name Undo-PWLogin -ErrorAction SilentlyContinue) {
        try {
            Undo-PWLogin
        }
        catch {
            Write-Log "Undo-PWLogin failed: $($_.Exception.Message)" "WARN"
        }
    }
}
