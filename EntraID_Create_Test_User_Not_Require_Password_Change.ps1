#  If creating multiple account then comment out last line of 
#  script "Disconnect-MgGraph" which disconnect your session from Azure
#
#  Make sure you run that line after completing creation of all accounts!!

##############################################################################
###################### For install of Microsoft Graph  #######################
Get-InstalledModule Microsoft.Graph
#  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
#  Install-Module Microsoft.Graph -Scope AllUsers -Repository PSGallery -Force
#  break
##############################################################################

#connects to Azure with login required
Connect-MgGraph -scope User.ReadWrite.All


# Asks for username and creates email and displayname
cls
write-host "Enter test account username" -ForegroundColor Yellow
$displayName = (Read-Host -Prompt "Username").ToLower() 
$userName = ($displayName + '@bconline.onmicrosoft.com')
cls

#Asks for employeeId retrieved from Controls & Compliance and adds employee type
write-host "Enter provided employee ID (non-human identity) from Controls & Compliance" -ForegroundColor Yellow
$employeeId = (Read-host "employee ID")
$employeeType = "T"

cls

#Asks for managers email
write-host "Enter managers email address." -ForegroundColor Yellow
$manageremail = (Read-host "Manager's Email Address")


cls
#Asks to enter password in secured popup
$password = Read-Host -Prompt "Enter 20 character random password"

#Creates password profile to not require password change on next logon along with password defined
$passwordProfile = @{
    
    
    forceChangePasswordNextSignIn = $false
    forceChangePasswordNextSignInWithMfa = $false
    Password = $password
}


#Create the user
New-MgUser -DisplayName $displayName -UserPrincipalName $userName -PasswordProfile $passwordProfile -AccountEnabled -MailNickname $displayname -EmployeeId $employeeId -EmployeeType $employeeType #-PasswordPolicies "DisablePasswordExpiration"


#define user and manager Azure object IDs to add users manager below
$managerId = (Get-MgUser -Filter "userPrincipalName eq '$manageremail'").Id
$userId = (Get-MgUser -Filter "userPrincipalName eq '$userName'").Id
$Manager = @{
                    '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($managerId)"
                }

#Add users manager to account
Set-MgUserManagerByRef -UserId $userId -OdataId "https://graph.microsoft.com/v1.0/users/$managerId"


#disconnect user from Azure
#Disconnect-MgGraph

    
    