<#
=============================================================================================
Name = Cengiz YILMAZ
Microsoft Certified Trainer (MCT)
Date = 23.03.2023
www.cengizyilmaz.net
www.cozumpark.com/author/cengizyilmaz
============================================================================================
#>

# Mail, Subject and File
$From = "from@contoso.com"
$To = "to@contoso.com"
$SMTPServer = "mail@contoso.com"
$Port = 587
$Priority = "High"
$Subject = "Exchange Server Backup Report"
$CredentialFile = "C:\Backup Report\credentials.backup"

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

# Backup Report Folder
if (!(Test-Path "C:\Backup Report")) {
    New-Item -ItemType Directory -Path "C:\Backup Report"
}

# Credential Check and Save
if (!(Test-Path $CredentialFile)) {
    Get-Credential -Message "Lütfen $From hesabı ile hesap bilgilerini doğrulayın."| Export-Clixml -Path $CredentialFile
}
$Credential = Import-Clixml -Path $CredentialFile

# Database Info
$Databases = Get-MailboxDatabase -Status | Select Name, Server, Mounted, LastFullBackup, LastIncrementalBackup, DatabaseSize, MasterType, MasterServerOrAvailabilityGroup, @{Name='Mailboxes';Expression={(Get-Mailbox -Database $_.Name).Count}}

# Number of unsupported databases
$UnbackedUpCount = ($Databases | Where-Object {-not $_.LastFullBackup -and -not $_.LastIncrementalBackup}).Count
$BackedUpCount = ($Databases.Count) - $UnbackedUpCount

#
$AllDbCount = ($Databases).Count
$BackedUpCount = ($Databases | Where-Object { $_.LastFullBackup }).Count
$UnbackedUpCount = $AllDbCount - $BackedUpCount

# Creating an HTML report
$HTMLReport = @"
<!DOCTYPE html>
<html>
<head>
<style>
table {
    width: 100%;
    border-collapse: collapse;
}
table, th, td {
    border: 1px solid black;
}
th, td {
    padding: 15px;
    text-align: left;
}
th {
    background-color: #f2f2f2;
}
.fail {
    color: red;
    background-color: #fdd;
}
.success {
    background-color: #dfd;
}
.incremental {
    background-color: #ffea7f;
}
</style>
</head>
<body>

<h2>Exchange Server Backup Report</h2>
<p>Toplamda $AllDbCount DB bulunmaktadir. Bunlardan $BackedUpCount tanesi yedeklenmistir ve $UnbackedUpCount tanesi yedeklenmemistir.</p>

<h3>Yedeklenmemis Databases</h3>
<table>
<tr>
    <th>Name</th>
    <th>Server</th>
    <th>DAG</th>
    <th>Health</th>
    <th>Backup Type</th>
    <th>Backup Time</th>
    <th>Database Size</th>
    <th>Mailboxes</th>
</tr>

$($Databases | Where-Object {-not $_.LastFullBackup -and -not $_.LastIncrementalBackup} | ForEach-Object {
    $backupType = "<td class='fail'>Fail</td>"
    $backupTime = 'N/A'
    $health = "$($_.Mounted) / $($_.MasterType)"
    @"
    <tr>
        <td>$($_.Name)</td>
        <td>$($_.Server)</td>
        <td>$($_.MasterServerOrAvailabilityGroup)</td>
        <td>$health</td>
        $backupType
        <td>$backupTime</td>
        <td>$($_.DatabaseSize)</td>
        <td>$($_.Mailboxes)</td>
    </tr>
"@})
</table>

<h3>Yedeklenmis Databases</h3>
<table>
<tr>
    <th>Name</th>
    <th>Server</th>
    <th>DAG</th>
    <th>Health</th>
    <th>Backup Type</th>
    <th>Backup Time</th>
    <th>Database Size</th>
    <th>Mailboxes</th>
</tr>
$($Databases | Where-Object { $_.LastFullBackup -or $_.LastIncrementalBackup } | ForEach-Object {
    $backupType = if ($_.LastFullBackup) { "<td class='success'>Full</td>" } elseif ($_.LastIncrementalBackup) { "<td class='incremental'>Incremental</td>" } else { "<td class='fail'>Fail</td>" }
    $backupTime = if ($_.LastFullBackup) { $_.LastFullBackup } elseif ($_.LastIncrementalBackup) { $_.LastIncrementalBackup } else { 'N/A' }
    $health = "$($_.Mounted) / $($_.MasterType)"
    @"
    <tr>
        <td>$($_.Name)</td>
        <td>$($_.Server)</td>
        <td>$($_.MasterServerOrAvailabilityGroup)</td>
        <td>$health</td>
        $backupType
        <td>$backupTime</td>
        <td>$($_.DatabaseSize)</td>
        <td>$($_.Mailboxes)</td>
    </tr>
"@})
</table>
</body>
</html>
"@

# Sending the report by e-mail
$MessageParameters = @{
    From = $From
    To = $To
    Subject = $Subject
    Priority = $Priority
    Body = $HTMLReport
    BodyAsHtml = $true
    SmtpServer = $SMTPServer
    Port = $Port
    Credential = $Credential
}

Send-MailMessage @MessageParameters