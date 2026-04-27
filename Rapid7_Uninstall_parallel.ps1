# ============================
# CONFIGURATION
# ============================
$serversFile = "C:\Temp\R7\servers.txt"
$csvPath     = "C:\Temp\R7\UninstallResults3.csv"

$servers = Get-Content $serversFile

# Parallel settings
$minThreads = 1
$maxThreads = 10   # Run 10 servers in parallel

# Output collection
$results = [System.Collections.Generic.List[object]]::new()

# ============================
# CREATE RUNSPACE POOL
# ============================
Write-Host "Starting parallel execution (Max Threads = $maxThreads)..." -ForegroundColor Cyan

$pool = [runspacefactory]::CreateRunspacePool($minThreads, $maxThreads)
$pool.Open()

$jobs = @()

# ============================
# SCRIPT BLOCK (PER SERVER)
# ============================
$script = {
    param($server)

    try {
        Write-Host "Processing $server" -ForegroundColor Cyan

        # ---------------------------
        # Step 1: Check WinRM
        # ---------------------------
        try {
            Test-WSMan -ComputerName $server -ErrorAction Stop | Out-Null
        }
        catch {
            return [PSCustomObject]@{
                Server  = $server
                Status  = "WinRM Failed"
                Message = $_.Exception.Message
            }
        }

        # ---------------------------
        # Step 2: Execute uninstall remotely
        # ---------------------------
        $output = Invoke-Command -ComputerName $server -ScriptBlock {

            try {
                $softwareKey = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" |
                    Get-ItemProperty |
                    Where-Object { $_.DisplayName -like "*Rapid7 Insight Agent*" } |
                    Select-Object -First 1

                if ($softwareKey) {
                    $uninstallString = $softwareKey.UninstallString
                    $uninstallString = $uninstallString.Replace("MsiExec.exe /I","/X ")
                    $uninstallString = $uninstallString.Replace("MsiExec.exe /X","/X ")

                    $process = Start-Process -FilePath "MsiExec.exe" `
                        -ArgumentList "$uninstallString /qn /norestart" `
                        -Wait -PassThru

                    if ($process.ExitCode -eq 0) {
                        return @{
                            Status  = "Success"
                            Message = "Uninstalled successfully"
                        }
                    }
                    else {
                        return @{
                            Status  = "Failed"
                            Message = "ExitCode: $($process.ExitCode)"
                        }
                    }
                }
                else {
                    return @{
                        Status  = "NotFound"
                        Message = "Software not installed"
                    }
                }
            }
            catch {
                return @{
                    Status  = "Error"
                    Message = $_.Exception.Message
                }
            }

        } -ErrorAction Stop

        # ---------------------------
        # Return structured result
        # ---------------------------
        [PSCustomObject]@{
            Server  = $server
            Status  = $output.Status
            Message = $output.Message
        }

    }
    catch {
        [PSCustomObject]@{
            Server  = $server
            Status  = "Execution Failed"
            Message = $_.Exception.Message
        }
    }
}

# ============================
# START PARALLEL EXECUTION
# ============================
foreach ($server in $servers) {

    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool

    $ps.AddScript($script).AddArgument($server) | Out-Null

    $handle = $ps.BeginInvoke()

    $jobs += [PSCustomObject]@{
        PowerShell = $ps
        Handle     = $handle
        Server     = $server
    }
}

# ============================
# COLLECT RESULTS
# ============================
Write-Host "Waiting for all jobs to complete..." -ForegroundColor Yellow

foreach ($job in $jobs) {

    try {
        $output = $job.PowerShell.EndInvoke($job.Handle)

        foreach ($item in $output) {
            $results.Add($item) | Out-Null
        }
    }
    catch {
        $results.Add([PSCustomObject]@{
            Server  = $job.Server
            Status  = "Runspace Error"
            Message = $_.Exception.Message
        }) | Out-Null
    }

    $job.PowerShell.Dispose()
}

# Cleanup
$pool.Close()
$pool.Dispose()

# ============================
# EXPORT RESULTS
# ============================
$results | Export-Csv -Path $csvPath -NoTypeInformation -Force

Write-Host "`nExecution completed." -ForegroundColor Green
Write-Host "Results exported to: $csvPath" -ForegroundColor Green