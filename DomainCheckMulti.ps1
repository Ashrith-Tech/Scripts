# Path to the server list text file (one server name per line)
$serverListPath = "D:\Scripts\Windows\DomainCheck\ServersList.txt"

# Read server names from the text file
$servers = Get-Content -Path $serverListPath

# Initialize variables
$problemServers = @()
$logPath = "D:\Scripts\Windows\DomainCheck\Tier2_DomainStatus.log"

# Create Logs folder if it doesn't exist
if (!(Test-Path "C:\Logs")) {
    New-Item -Path "C:\Logs" -ItemType Directory -Force | Out-Null
}

# Clear previous log content or create new file
Clear-Content $logPath -ErrorAction SilentlyContinue
if (!(Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType File -Force | Out-Null
}

foreach ($server in $servers) {
    try {
        # Get domain info remotely
        $domain = Invoke-Command -ComputerName $server -ScriptBlock {
            (Get-WmiObject Win32_ComputerSystem).Domain
        } -ErrorAction Stop

        if ($domain -eq $server) {
            $status = "$server is NOT domain joined."
            $problemServers += $server
        } else {
            $status = "$server is domain joined to $domain."
        }

        # Write status to log
        "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) - $status" | Out-File $logPath -Append
    }
    catch {
        $errorMsg = "$server is unreachable or error occurred: $_"
        $problemServers += $server
        "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) - $errorMsg" | Out-File $logPath -Append
    }
}

# Email parameters
$smtpServer = "smtp.brunswick.com"   # Replace with your SMTP server
$from = "Tier2.Domain_Status@brunswick.com"
$to = "BC.IT.TCS.WinServer@brunswick.com","bc.it.tcs.commandcenter@brunswick.com"
$subject = "Tier2 Domain Join Status Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

if ($problemServers.Count -gt 0) {
    $body = "Warning! The following servers have issues with domain trust or are unreachable:`n"
    $body += $problemServers -join "`n"
} else {
    $body = "All servers are properly joined to the domain."
}

# Send email with the log attached, named Tier2_DomainStatus.log
Send-MailMessage -From $from -To $to -Subject $subject -Body $body -SmtpServer $smtpServer -Attachments $logPath -BodyAsHtml:$false