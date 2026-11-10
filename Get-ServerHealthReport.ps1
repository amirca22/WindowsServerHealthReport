$ErrorActionPreference = 'Stop'

$serversFile = Join-Path -Path $PSScriptRoot -ChildPath 'servers.txt'

if (-not (Test-Path -LiteralPath $serversFile -PathType Leaf)) {
    throw "The server list was not found: $serversFile"
}

$servers = @(
    Get-Content -LiteralPath $serversFile |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith('#') } |
        Select-Object -Unique
)

if ($servers.Count -eq 0) {
    throw "No server names were found in $serversFile. Add one server name per line."
}

while ($true) {
    $daysInput = Read-Host 'How many days back should installed updates be checked?'
    $parsedDays = 0
    if ([int]::TryParse($daysInput, [ref]$parsedDays) -and $parsedDays -ge 1) {
        $days = $parsedDays
        break
    }

    Write-Warning 'Enter a whole number greater than zero.'
}

$username = Read-Host 'Enter an Active Directory account (DOMAIN\username or username@domain)'
if ([string]::IsNullOrWhiteSpace($username)) {
    throw 'No username was entered. No report was generated.'
}
$password = Read-Host 'Enter the account password' -AsSecureString
$credential = [System.Management.Automation.PSCredential]::new($username, $password)

$remoteQuery = {
    param([int]$Days)

    $cutoff = (Get-Date).AddDays(-$Days)
    $updates = @(
        foreach ($hotfix in (Get-HotFix -ErrorAction Stop)) {
            if ($hotfix.InstalledOn) {
                $installedOn = [datetime]::MinValue
                if ([datetime]::TryParse(
                        [string]$hotfix.InstalledOn,
                        [Globalization.CultureInfo]::CurrentCulture,
                        [Globalization.DateTimeStyles]::AssumeLocal,
                        [ref]$installedOn
                    ) -and $installedOn -ge $cutoff) {
                    [pscustomobject]@{
                        HotFixId    = $hotfix.HotFixID
                        Description = $hotfix.Description
                        InstalledOn = $installedOn.ToString('yyyy-MM-dd HH:mm:ss')
                        InstalledBy = $hotfix.InstalledBy
                    }
                }
            }
        }
    )

    $disks = @(
        Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType = 3' -ErrorAction Stop |
            ForEach-Object {
                [pscustomobject]@{
                    Drive       = $_.DeviceID
                    VolumeName  = $_.VolumeName
                    SizeGB      = if ($_.Size) { [math]::Round($_.Size / 1GB, 2) } else { 0 }
                    FreeGB      = if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 2) } else { 0 }
                    FreePercent = if ($_.Size -gt 0) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 1) } else { 0 }
                }
            }
    )

    [pscustomobject]@{
        Server      = $env:COMPUTERNAME
        CollectedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Updates     = $updates
        Disks       = $disks
    }
}

$results = foreach ($server in $servers) {
    Write-Host "Checking $server ..."
    try {
        $remoteResult = Invoke-Command -ComputerName $server -Credential $credential `
            -ScriptBlock $remoteQuery -ArgumentList $days -ErrorAction Stop
        [pscustomobject]@{
            Server      = $server
            Success     = $true
            CollectedAt = $remoteResult.CollectedAt
            Updates     = @($remoteResult.Updates)
            Disks       = @($remoteResult.Disks)
            Error       = $null
        }
    }
    catch {
        $record = $_
        $errorDetails = @(
            "Message: $($record.Exception.Message)"
            "Exception type: $($record.Exception.GetType().FullName)"
            "Category: $($record.CategoryInfo.Category)"
            "Fully qualified error ID: $($record.FullyQualifiedErrorId)"
            "Target: $($record.CategoryInfo.TargetName)"
            "Script stack: $($record.ScriptStackTrace)"
            'Check that the server name resolves, WinRM is enabled and reachable (usually TCP 5985/5986), and this account is authorized for PowerShell remoting.'
        ) -join "`r`n"

        [pscustomobject]@{
            Server      = $server
            Success     = $false
            CollectedAt = $null
            Updates     = @()
            Disks       = @()
            Error       = $errorDetails
        }
        Write-Warning "Could not query $server. Details will be included in the HTML report."
    }
}

function ConvertTo-HtmlText {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

$successCount = @($results | Where-Object Success).Count
$failureCount = $results.Count - $successCount
$updateCount = @($results | Where-Object Success | ForEach-Object { $_.Updates }).Count
$diskCount = @($results | Where-Object Success | ForEach-Object { $_.Disks }).Count
$serverCount = $results.Count
$serverSections = foreach ($result in $results) {
    $serverName = ConvertTo-HtmlText $result.Server
    if (-not $result.Success) {
        $errorText = ConvertTo-HtmlText $result.Error
        @"
<section class="server">
  <h2>$serverName <span class="badge failed">Connection failed</span></h2>
  <pre class="error">$errorText</pre>
</section>
"@
        continue
    }

    $updateRows = foreach ($update in $result.Updates) {
        '<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>' -f `
            (ConvertTo-HtmlText $update.HotFixId),
            (ConvertTo-HtmlText $update.Description),
            (ConvertTo-HtmlText $update.InstalledOn),
            (ConvertTo-HtmlText $update.InstalledBy)
    }
    if (@($result.Updates).Count -eq 0) {
        $updateRows = @('<tr><td colspan="4" class="empty">No updates found in the selected period.</td></tr>')
    }

    $diskRows = foreach ($disk in $result.Disks) {
        '<tr><td>{0}</td><td>{1}</td><td>{2:N2} GB</td><td>{3:N2} GB</td><td>{4:N1}%</td></tr>' -f `
            (ConvertTo-HtmlText $disk.Drive),
            (ConvertTo-HtmlText $disk.VolumeName),
            $disk.SizeGB,
            $disk.FreeGB,
            $disk.FreePercent
    }
    if (@($result.Disks).Count -eq 0) {
        $diskRows = @('<tr><td colspan="5" class="empty">No fixed disks were reported.</td></tr>')
    }

    $collectedAt = ConvertTo-HtmlText $result.CollectedAt
    @"
<section class="server">
  <h2>$serverName <span class="badge success">Connected</span></h2>
  <p class="meta">Collected: $collectedAt</p>
  <h3>Installed updates in the last $days days</h3>
  <table>
    <thead><tr><th>Hotfix ID</th><th>Description</th><th>Installed</th><th>Installed by</th></tr></thead>
    <tbody>$($updateRows -join "`n")</tbody>
  </table>
  <h3>Fixed disks</h3>
  <table>
    <thead><tr><th>Drive</th><th>Volume</th><th>Total</th><th>Free</th><th>Free %</th></tr></thead>
    <tbody>$($diskRows -join "`n")</tbody>
  </table>
</section>
"@
}

$generatedAt = Get-Date
$reportPath = Join-Path -Path $PSScriptRoot -ChildPath (
    'ServerHealthReport_{0}.html' -f $generatedAt.ToString('yyyyMMdd_HHmmss')
)
$html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Server health report</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; color: #24292f; margin: 2rem auto; max-width: 1100px; padding: 0 1rem; }
    h1 { margin-bottom: .25rem; }
    .meta { color: #57606a; }
    .stats { display: flex; flex-wrap: wrap; gap: 1rem; margin: 1.5rem 0; }
    .stat { background: #f6f8fa; border: 1px solid #d0d7de; border-radius: 6px; min-width: 140px; padding: 1rem; }
    .stat strong { display: block; font-size: 1.5rem; }
    .server { border-top: 2px solid #d0d7de; margin-top: 2rem; padding-top: 1rem; }
    .badge { border-radius: 1rem; font-size: .75rem; padding: .2rem .6rem; vertical-align: middle; }
    .success { background: #dafbe1; color: #116329; }
    .failed { background: #ffebe9; color: #82071e; }
    table { border-collapse: collapse; margin: .75rem 0 1.5rem; width: 100%; }
    th, td { border: 1px solid #d0d7de; padding: .55rem; text-align: left; }
    th { background: #f6f8fa; }
    .empty { color: #57606a; font-style: italic; }
    .error { background: #ffebe9; border: 1px solid #ff8182; border-radius: 4px; overflow-wrap: anywhere; padding: 1rem; white-space: pre-wrap; }
    @media (max-width: 700px) { table { font-size: .85rem; } th, td { padding: .35rem; } }
  </style>
</head>
<body>
  <h1>Server health report</h1>
  <p class="meta">Generated $($generatedAt.ToString('yyyy-MM-dd HH:mm:ss')) | Update lookback: $days days | Servers: $($results.Count)</p>
  <div class="stats">
    <div class="stat"><strong>$serverCount</strong>Servers checked</div>
    <div class="stat"><strong>$successCount</strong>Connected servers</div>
    <div class="stat"><strong>$failureCount</strong>Connection failures</div>
    <div class="stat"><strong>$updateCount</strong>Recent updates</div>
    <div class="stat"><strong>$diskCount</strong>Fixed disks checked</div>
  </div>
  $($serverSections -join "`n")
</body>
</html>
"@

Set-Content -LiteralPath $reportPath -Value $html -Encoding UTF8
Write-Host "Report saved to: $reportPath"
Write-Host "Summary: $successCount connected, $failureCount failed, $updateCount updates, $diskCount disks."
