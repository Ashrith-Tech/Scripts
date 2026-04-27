$servers = Get-Content "C:\Temp\R7\servers.txt"
$results = @()

foreach ($server in $servers) {

    Write-Host "Processing $server..." -ForegroundColor Cyan

    try {
        # Test WinRM connectivity first
        if (-not (Test-WSMan -ComputerName $server -ErrorAction Stop)) {
            throw "WinRM not reachable"
        }

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
                            Status = "Success"
                            Message = "Uninstalled successfully"
                        }
                    }
                    else {
                        return @{
                            Status = "Failed"
                            Message = "Uninstall failed with exit code $($process.ExitCode)"
                        }
                    }
                }
                else {
                    return @{
                        Status = "NotFound"
                        Message = "Software not installed"
                    }
                }
            }
            catch {
                return @{
                    Status = "Error"
                    Message = $_.Exception.Message
                }
            }
        }

        $results += [PSCustomObject]@{
            Server  = $server
            Status  = $output.Status
            Message = $output.Message
        }

    }
    catch {
        # This block catches WinRM / connection issues
        $results += [PSCustomObject]@{
            Server  = $server
            Status  = "WinRM Failed"
            Message = $_.Exception.Message
        }
    }
}

# Export results
$results | Export-Csv "C:\Temp\R7\UninstallResults.csv" -NoTypeInformation

Write-Host "Execution completed. Results exported to CSV." -ForegroundColor Green