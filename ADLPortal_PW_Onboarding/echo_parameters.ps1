param(
    [Parameter(Mandatory = $true)]
    [string]$Project,

    [Parameter(Mandatory = $true)]
    [string]$UserName,

    [Parameter(Mandatory = $true)]
    [string]$Email,

    [string]$RequestId = "",
    [string]$Source = "manual",
    [int]$SimulateSeconds = 0,
    [switch]$AsJson
)

if ($SimulateSeconds -gt 0) {
    Start-Sleep -Seconds $SimulateSeconds
}

$result = [ordered]@{
    timestampUtc = (Get-Date).ToUniversalTime().ToString("o")
    result       = "success"
    message      = "Parameters echoed successfully"
    parameters   = [ordered]@{
        Project         = $Project
        UserName        = $UserName
        Email           = $Email
        RequestId       = $RequestId
        Source          = $Source
        SimulateSeconds = $SimulateSeconds
        AsJson          = [bool]$AsJson
    }
}

if ($AsJson) {
    $result | ConvertTo-Json -Depth 4
} else {
    Write-Host "timestampUtc: $($result.timestampUtc)"
    Write-Host "result: $($result.result)"
    Write-Host "message: $($result.message)"
    Write-Host "parameters:"
    $result.parameters.GetEnumerator() | ForEach-Object {
        Write-Host "  $($_.Key): $($_.Value)"
    }
}
