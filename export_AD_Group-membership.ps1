$groupName = "ServerSoftwareUpdatesMfgTier_Sun-Morning_Ent_AppGG"

$outputCsvPath = "C:\Users\gkota\OneDrive - Brunswick Corporation\Desktop\Patching Sheet\MfgTier_Sun-Morn.csv"

Import-Module ActiveDirectory

$group = Get-ADGroup -Identity $groupName
$groupMembers = Get-ADGroupMember -Identity $group -Recursive | Where-Object { $_.objectClass -eq 'computer' }

$computerDetails = @()

foreach ($computer in $groupMembers) {
    $computerInfo = Get-ADComputer -Identity $computer.DistinguishedName -Properties CanonicalName
    $computerDetails += [PSCustomObject]@{
        Name         = $computerInfo.Name
        CanonicalName = $computerInfo.CanonicalName
    }
}

$computerDetails | Export-Csv -Path $outputCsvPath -NoTypeInformation

Write-Output "Computer details exported to $outputCsvPath"
