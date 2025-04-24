<#
    Exchange Server 2019 Backup & Health Report Script (Final Revised)
    -------------------------------------------------------------
    Description:
      This script collects Exchange Server mailbox database backup statuses,
      evaluates key health details for each database, and builds a modern, responsive
      HTML report. The report includes backup status/time information, detailed "DB Status",
      and a Server-DAG section that explicitly separates the Active and Passive server names.
      In the Server-DAG section, active servers are shown with an "A:" label and passive
      servers with a "P:" label. Disk details (total, used, free in GB and free percentage)
      are obtained from Get-MailboxDatabaseCopyStatus properties: DiskTotalSpace, DiskFreeSpace,
      DiskFreeSpacePercent, and DatabaseVolumeMountPoint. If the disk free percentage is less than
      or equal to the defined threshold, the disk info is highlighted.
      
      The final table is sorted as follows:
         1. Databases with no backup ("None") first,
         2. then "In Progress",
         3. then backed up databases.
      Within each group, the results are sorted by ActiveServer then by database Name.
      
      The HTML report is responsive (includes viewport meta tag) for mobile devices.
      The TO email field supports multiple addresses separated by semicolons.
      
      Service status checks have been removed.
      
    Author:          Cengiz YILMAZ
    Date:            4/15/2025
    Title:           Microsoft MVP - MCT
    WebSite:         https://cengizyilmaz.net
#>

###############################################
# Function: Modern-WriteProgress
# A wrapper for Write-Progress to display progress updates.
###############################################
function Modern-WriteProgress {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Activity,
        [Parameter(Mandatory = $true)]
        [int]$PercentComplete,
        [string]$Status = ""
    )
    Write-Progress -Activity $Activity -PercentComplete $PercentComplete -Status $Status
}

###############################################
# Commit: Configuration and Central Variables
###############################################
$Config = @{
    # Email Settings
    From           = "no-reply@fixcloud.com.tr"
    To             = "cengiz.yilmaz@fixcloud.com.tr;kerem.asci@fixcloud.com.tr"
    SMTPServer     = "mail.fixcloud.com.tr"
    Port           = 587
    Priority       = "High"
    Subject        = "Exchange Server Backup & Health Report"
    CredentialFile = "C:\Backup Report\credentials.backup"
    
    # Time Thresholds
    BackupThreshold          = (Get-Date).AddHours(-24)
    CopyQueueLengthThreshold = 5

    # Free disk space threshold (in percent)
    FreeSpaceThreshold       = 10

    # Backup Report Directory
    BackupReportDir          = "C:\Backup Report"

    # CSS Color Classes for HTML Report – used as table cell background classes.
    CSSClasses = @{
        InProgress      = "inprogress"
        Success         = "success"
        Fail            = "fail"
        Mounted         = "mounted"
        Mounting        = "mounting"
        Dismounted      = "dismounted"
        Dismounting     = "dismounting"
        Healthy         = "healthy"
        Failed          = "failed"
        Suspended       = "suspended"
        Seeding         = "seeding"
        Initializing    = "initializing"
        Resynchronizing = "resynchronizing"
        LowDisk         = "lowDisk"
        Default         = "default"
    }
}

###############################################
# Commit: Load Exchange Snap-In and Setup Environment
###############################################
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
if (!(Test-Path $Config.BackupReportDir)) {
    New-Item -ItemType Directory -Path $Config.BackupReportDir | Out-Null
}

###############################################
# Commit: Credential Check and Save
###############################################
if (!(Test-Path $Config.CredentialFile)) {
    Get-Credential -Message "Please verify the credentials for account $($Config.From)" | Export-Clixml -Path $Config.CredentialFile
}
$Credential = Import-Clixml -Path $Config.CredentialFile

###############################################
# Commit: Fetch DAG, Mailbox Database and Health Information with Progress
###############################################
Modern-WriteProgress -Activity "Fetching Mailbox Databases" -PercentComplete 0 -Status "Initializing..."
$allDAGs = Get-DatabaseAvailabilityGroup
$mailboxDatabases = Get-MailboxDatabase -Status
$totalDBs = $mailboxDatabases.Count

$Databases = @()
$currentDBIndex = 0

foreach ($db in $mailboxDatabases) {
    $currentDBIndex++
    $percentComplete = [math]::Round(($currentDBIndex / $totalDBs * 100), 0)
    Modern-WriteProgress -Activity "Processing Mailbox Databases" -PercentComplete $percentComplete -Status "Processing database: $($db.Name)"
    
    $dbName = $db.Name
    $activeServer = $db.Server

    # Retrieve all copy statuses for the database (active and passive copies)
    $dbCopies = Get-MailboxDatabaseCopyStatus -Identity "$dbName\*" -ErrorAction SilentlyContinue

    ###############################################
    # Server-DAG:
    # Build server list by splitting the Identity property (format: "DatabaseName\ServerName").
    # Separate active and passive servers. Active servers are listed first.
    ###############################################
    $allServers = ($dbCopies | ForEach-Object { ($_."Identity" -split '\\')[1] } | Sort-Object -Unique)
    $activeServers = $allServers | Where-Object { $_ -eq $activeServer }
    $passiveServers = $allServers | Where-Object { $_ -ne $activeServer } | Sort-Object
    $activeLabel = if ($activeServers) { "A: " + ($activeServers -join ", ") } else { "A: $activeServer" }
    $passiveLabel = if ($passiveServers) { "P: " + ($passiveServers -join ", ") } else { "" }
    if ($passiveLabel) {
        $serverDag = "$activeLabel | $passiveLabel"
    }
    else {
        $serverDag = $activeLabel
    }
    if ($db.MasterServerOrAvailabilityGroup) {
        $serverDag += " | DAG: $($db.MasterServerOrAvailabilityGroup)"
    }

    ###############################################
    # Health Details: Evaluate each copy's status and CopyQueueLength.
    ###############################################
    $healthDetails = ($dbCopies | ForEach-Object {
        $status = $_.Status
        $healthClass = $Config.CSSClasses.Default
        switch ($status) {
            "Dismounted"       { $healthClass = $Config.CSSClasses.Dismounted }
            "Suspended"        { $healthClass = $Config.CSSClasses.Suspended }
            "Failed"           { $healthClass = $Config.CSSClasses.Fail }
            "Resync"           { $healthClass = $Config.CSSClasses.Resynchronizing }
            "Resynchronizing"  { $healthClass = $Config.CSSClasses.Resynchronizing }
            default {
                if ($_.CopyQueueLength -gt $Config.CopyQueueLengthThreshold) {
                    $healthClass = $Config.CSSClasses.Fail
                }
            }
        }
        "<span class='$healthClass'>$($status): CQ=$($_.CopyQueueLength)</span>"
    }) -join "; "

    ###############################################
    # Backup Status: Determine backup type.
    ###############################################
    $backupTime = if ($db.LastFullBackup) { $db.LastFullBackup.ToString("yyyy-MM-dd HH:mm:ss") } else { "N/A" }
    if ($db.BackupInProgress) {
        $backupStatusText = "In Progress"
        $backupStatusClass = $Config.CSSClasses.InProgress
    }
    elseif ($db.LastFullBackup -and ($db.LastFullBackup -gt $Config.BackupThreshold)) {
        $backupStatusText = "Full Backup"
        $backupStatusClass = $Config.CSSClasses.Success
    }
    elseif (($db.PSObject.Properties.Name -contains "LastIncrementalBackup") -and $db.LastIncrementalBackup -and ($db.LastIncrementalBackup -gt $Config.BackupThreshold)) {
        $backupStatusText = "Incremental Backup"
        $backupStatusClass = $Config.CSSClasses.Success
    }
    elseif ($dbCopies | Where-Object { $_.CopyQueueLength -gt $Config.CopyQueueLengthThreshold }) {
        $backupStatusText = "Warning: CopyQueue Exceeded"
        $backupStatusClass = $Config.CSSClasses.Fail
    }
    else {
        $backupStatusText = "None"
        $backupStatusClass = $Config.CSSClasses.Fail
    }
    
    ###############################################
    # DB Status: Determine database status based on active copy.
    # Use the server portion of the Identity property to match the active server.
    ###############################################
    $activeCopy = $dbCopies | Where-Object { (($_."Identity" -split '\\')[1]) -eq $activeServer } | Select-Object -First 1
    if (-not $activeCopy -and $dbCopies.Count -gt 0) {
         $activeCopy = $dbCopies[0]
    }
    if ($activeCopy) {
         $statusValue = $activeCopy.Status
    }
    else {
         $statusValue = "Unknown"
    }
    switch ($statusValue) {
        "Mounted"           { $dbStatusClass = $Config.CSSClasses.Mounted;      $dbStatusText = "Mounted" }
        "Mounting"          { $dbStatusClass = $Config.CSSClasses.Mounting;     $dbStatusText = "Mounting" }
        "Dismounted"        { $dbStatusClass = $Config.CSSClasses.Dismounted;   $dbStatusText = "Dismounted" }
        "Dismounting"       { $dbStatusClass = $Config.CSSClasses.Dismounting;  $dbStatusText = "Dismounting" }
        "Healthy"           { $dbStatusClass = $Config.CSSClasses.Healthy;      $dbStatusText = "Healthy" }
        "Failed"            { $dbStatusClass = $Config.CSSClasses.Fail;         $dbStatusText = "Failed" }
        "Suspended"         { $dbStatusClass = $Config.CSSClasses.Suspended;    $dbStatusText = "Suspended" }
        "Seeding"           { $dbStatusClass = $Config.CSSClasses.Seeding;      $dbStatusText = "Seeding" }
        "Initializing"      { $dbStatusClass = $Config.CSSClasses.Initializing; $dbStatusText = "Initializing" }
        "Resynchronizing"   { $dbStatusClass = $Config.CSSClasses.Resynchronizing; $dbStatusText = "Resynchronizing" }
        default             { $dbStatusClass = $Config.CSSClasses.Default;      $dbStatusText = $statusValue }
    }
    
    ###############################################
    # WhiteSpace/Mailboxes: Merge free space info and mailbox count.
    ###############################################
    $freeSpaceText = if ($db.AvailableNewMailboxSpace -gt 1GB) {
        "{0:N2} GB" -f ($db.AvailableNewMailboxSpace.ToGB())
    }
    else {
        "{0:N2} MB" -f ($db.AvailableNewMailboxSpace.ToMB())
    }
    $mailboxCount = 0
    try {
        $mailboxCount = (Get-Mailbox -Database $dbName -ErrorAction SilentlyContinue).Count
    }
    catch {
        $mailboxCount = 0
    }
    $whiteSpaceMailboxes = "WhiteSpace: $freeSpaceText / Mailboxes: $mailboxCount"

    ###############################################
    # Disk Info: For each copy, use disk properties to build disk info.
    # Properties: DiskTotalSpace, DiskFreeSpace, DiskFreeSpacePercent, DatabaseVolumeMountPoint.
    # If DiskFreeSpacePercent is below or equal to threshold, wrap info in a CSS "lowDisk" div.
    ###############################################
    $diskInfoLines = @()
    $freePctValues = @()
    foreach ($copy in $dbCopies) {
        if ($copy.PSObject.Properties.Name -contains "DiskTotalSpace" -and $copy.DiskTotalSpace) {
            # Extract server name from Identity.
            $serverNameFromCopy = ($copy.Identity -split '\\')[1]
            $role = if ($serverNameFromCopy -eq $activeServer) { "Active" } else { "Passive" }
            $totalSpaceGB = [math]::Round($copy.DiskTotalSpace / 1GB, 2)
            $freeSpaceGB = [math]::Round($copy.DiskFreeSpace / 1GB, 2)
            $freePercent = $copy.DiskFreeSpacePercent  # expected numeric value
            $mountPoint = $copy.DatabaseVolumeMountPoint
            $diskDetails = "Drive: $mountPoint | Total: ${totalSpaceGB} GB, Free: ${freeSpaceGB} GB ($freePercent`%)"
            if ($freePercent -le $Config.FreeSpaceThreshold) {
                $diskDetails = "<div class='$($Config.CSSClasses.LowDisk)'>$diskDetails</div>"
            }
            $diskInfoLines += "${role}: $diskDetails"
            $freePctValues += $freePercent
        }
    }
    if ($diskInfoLines.Count -gt 0) {
        $diskInfoText = $diskInfoLines -join "<br/>"
    }
    else {
        $diskInfoText = "N/A"
    }
    # Determine minimum free percentage among copies (worst-case scenario)
    if ($freePctValues.Count -gt 0) {
        $minFreePercent = [math]::Round(($freePctValues | Measure-Object -Minimum).Minimum, 2)
    }
    else {
        $minFreePercent = 100
    }

    ###############################################
    # Create Custom Object for the Database
    ###############################################
    $Databases += [PSCustomObject]@{
        Name                = $dbName
        ServerDAG           = $serverDag
        DBStatusText        = $dbStatusText
        DBStatusClass       = $dbStatusClass
        Health              = $healthDetails
        BackupStatusText    = $backupStatusText
        BackupStatusClass   = $backupStatusClass
        BackupTime          = $backupTime
        DatabaseSize        = if ($db.DatabaseSize) { $db.DatabaseSize.ToString() } else { "N/A" }
        WhiteSpaceMailboxes = $whiteSpaceMailboxes
        DiskInfo            = $diskInfoText
        DiskFreePercentage  = $minFreePercent
        ActiveServer        = $activeServer
    }
}

Modern-WriteProgress -Activity "Processing Completed" -PercentComplete 100 -Status "Finished processing databases"

###############################################
# Commit: Sort Databases into groups and by server and DB name.
# Sort order:
#   1. BackupStatusText "None" (no backup) first,
#   2. then "In Progress",
#   3. then others (backed up).
# Within each group, sort by ActiveServer then by database Name.
###############################################
$sortedDatabases = $Databases | Sort-Object `
    @{ Expression = { if ($_.BackupStatusText -eq "None") { 1 } elseif ($_.BackupStatusText -eq "In Progress") { 2 } else { 3 } } }, `
    @{ Expression = { $_.ActiveServer } }, `
    @{ Expression = { $_.Name } }

###############################################
# Commit: Dynamic Alarm Messages for Email Body
###############################################
$dbAlarm = ""

# Alarm: Dismounted databases
$dismountedDatabases = $sortedDatabases | Where-Object { $_.Health -match "Dismounted" }
if ($dismountedDatabases.Count -gt 0) {
    $dbAlarm += "<p style='font-size:14px;color:red;'><strong>Alarm: Dismounted databases detected!</strong></p><ul style='font-size:13px;color:#555;'>"
    foreach ($dDB in $dismountedDatabases) {
         $dbAlarm += "<li>$($dDB.Name) on $($dDB.ServerDAG) - Health: $($dDB.Health)</li>"
    }
    $dbAlarm += "</ul>"
}

# Alarm: Failed databases – list each failed copy
$failedDatabases = $sortedDatabases | Where-Object { $_.Health -match "Failed" }
if ($failedDatabases.Count -gt 0) {
    $dbAlarm += "<p style='font-size:14px;color:red;'><strong>Alarm: Failed databases detected!</strong></p><ul style='font-size:13px;color:#555;'>"
    foreach ($fDB in $failedDatabases) {
         $dbAlarm += "<li>$($fDB.Name) - Health: $($fDB.Health)</li>"
    }
    $dbAlarm += "</ul>"
}

# Alarm: Low Disk Space detected (using DiskFreePercentage from our custom object)
$lowDiskDatabases = $sortedDatabases | Where-Object { $_.DiskFreePercentage -lt $Config.FreeSpaceThreshold }
if ($lowDiskDatabases.Count -gt 0) {
    $dbAlarm += "<p style='font-size:14px;color:red;'><strong>Alarm: Low Disk Space detected on the following databases!</strong></p><ul style='font-size:13px;color:#555;'>"
    foreach ($ldDB in $lowDiskDatabases) {
         $dbAlarm += "<li>$($ldDB.Name) on $($ldDB.ServerDAG) - Disk Free: $($ldDB.DiskFreePercentage)%</li>"
    }
    $dbAlarm += "</ul>"
}

###############################################
# Commit: Generate Modern HTML Report
###############################################
$backedUpCount    = ($sortedDatabases | Where-Object { $_.BackupStatusText -eq "Full Backup" -or $_.BackupStatusText -eq "Incremental Backup" }).Count
$inProgressCount  = ($sortedDatabases | Where-Object { $_.BackupStatusText -eq "In Progress" }).Count
$notBackedUpCount = $sortedDatabases.Count - $backedUpCount - $inProgressCount

$HTMLReport = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Exchange Server Backup & Health Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f4f4f4;
            color: #333;
        }
        h2, h3 {
            color: #333;
        }
        p {
            font-size: 14px;
            line-height: 1.6;
        }
        .table-responsive {
            width: 100%;
            overflow-x: auto;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        th, td {
            padding: 12px 15px;
            border: 1px solid #ddd;
            text-align: left;
            white-space: nowrap;
        }
        th {
            background-color: #4CAF50;
            color: #ffffff;
            font-weight: bold;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        /* CSS classes for cell background coloring */
        .inprogress { background-color: blue; color: white; }
        .success { background-color: green; color: white; }
        .fail { background-color: red; color: white; }
        .mounted { background-color: green; color: white; }
        .mounting { background-color: lightgreen; color: black; }
        .dismounted { background-color: purple; color: white; }
        .dismounting { background-color: mediumpurple; color: white; }
        .healthy { background-color: limegreen; color: white; }
        .failed { background-color: maroon; color: white; }
        .suspended { background-color: darkorange; color: white; }
        .seeding { background-color: blueviolet; color: white; }
        .initializing { background-color: orange; color: white; }
        .resynchronizing { background-color: teal; color: white; }
        .lowDisk { background-color: red; color: white; }
        .default { background-color: lightgray; color: black; }
    </style>
</head>
<body>
    <h2>Exchange Server Backup & Health Report</h2>
    <p>Total Databases: $($sortedDatabases.Count). In the last 24 hours, 
       <strong>$backedUpCount</strong> databases were fully backed up, 
       <strong>$notBackedUpCount</strong> were not backed up, and 
       <strong>$inProgressCount</strong> are in progress.</p>
    $dbAlarm
    <h3>Database Details</h3>
    <div class="table-responsive">
    <table>
        <tr>
            <th>Name</th>
            <th>Server - DAG</th>
            <th>DB Status</th>
            <th>Health</th>
            <th>Backup Status</th>
            <th>Database Size</th>
            <th>WhiteSpace/Mailboxes</th>
            <th>Disk Info</th>
        </tr>
"@

foreach ($db in $sortedDatabases) {
    $HTMLReport += @"
        <tr>
            <td>$($db.Name)</td>
            <td>$($db.ServerDAG)</td>
            <td class='$($db.DBStatusClass)'>$($db.DBStatusText)</td>
            <td>$($db.Health)</td>
            <td class='$($db.BackupStatusClass)'>$($db.BackupStatusText)<br/><small>$($db.BackupTime)</small></td>
            <td>$($db.DatabaseSize)</td>
            <td>$($db.WhiteSpaceMailboxes)</td>
            <td>$($db.DiskInfo)</td>
        </tr>
"@
}

$HTMLReport += @"
    </table>
    </div>
</body>
</html>
"@

###############################################
# Commit: Send Email Report with Progress Update
###############################################
Modern-WriteProgress -Activity "Sending Email" -PercentComplete 0 -Status "Initializing email sending process"
$MessageParameters = @{
    From       = $Config.From
    # TO addresses are split by semicolon to allow multiple recipients.
    To         = $Config.To -split ';'
    Subject    = $Config.Subject
    Priority   = $Config.Priority
    Body       = $HTMLReport
    BodyAsHtml = $true
    SmtpServer = $Config.SMTPServer
    Port       = $Config.Port
    Credential = $Credential
}
Send-MailMessage @MessageParameters
Modern-WriteProgress -Activity "Email Sent" -PercentComplete 100 -Status "Email report successfully sent"
