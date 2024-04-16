Get-PSSession | Remove-PSSession

#Set the title of the window.

$host.ui.RawUI.WindowTitle = "Leavers Script"

#Logging into 365

Write-Host
Write-Host -ForegroundColor Yellow "Enter your Office 365 Credentials"
Write-Host
$CloudCredential = Get-Credential $null

$ulist = Import-Csv C:\Location\Starters-Leavers\leavers.csv
$PermLeaversOU = 'OU'

Write-Host
Write-Host -ForegroundColor Green "Connecting to Office 365 and on-prem Exchange"
Write-Host

Connect-MsolService -Credential $CloudCredential

#Connect to Exchange Online using Modern Authentication

$CloudSessionParameters = @{
    Credential        = $CloudCredential
    WarningAction     = 'SilentlyContinue'
}
Connect-ExchangeOnline @CloudSessionParameters -ShowBanner:$false


#Connect to Local Exchange

$LocalExchangeParameters = @{
    ConfigurationName = 'Microsoft.Exchange'
    ConnectionUri     = 'http://servername/Powershell/'
    Authentication    = 'Kerberos'
    WarningAction     = 'SilentlyContinue'
}
$LocalSession = New-PSSession @LocalExchangeParameters
Import-PSSession $LocalSession -Prefix Onprem -DisableNameChecking | Out-Null
 
###### PART 1 ######
####################

#Importing the CSV file and going through steps to remove the user from all AD groups except 'Domain Users', disabling the user account, changing the AD password to a random password, disabling the remote mailbox and sending an email to IT
 
$ulist | ForEach-Object {
 
    try {
        $adacct = Get-ADUser $_.user -Properties Name, SamAccountname, UserPrincipalName, CanonicalName, Enabled, EmailAddress, PasswordExpired, Modified -ErrorAction Stop
    } catch {
        Write-Error "User $($_.user) does not exist, cannot disable"
        Add-Content -Path C:\Location\Starters-Leavers\UsersNotProcessed.log -Value $_.user
        #Skips to the next user in $ulist, does not disable anything
        continue
    }
 
 
    $body = Get-EXOMailbox -Identity $adacct.UserPrincipalName | Select-Object Name, Alias, EmailAddresses -ExpandProperty EmailAddresses
    
    $report = $adacct | Select-Object Name, SamAccountname, UserPrincipalName, CanonicalName, EmailAddress, PasswordExpired, Modified | Out-String

    Write-Host
    Write-Host -ForegroundColor Yellow "Taking note of all AD groups to email into the ticket"
    $adgroups = Get-AdPrincipalGroupMembership -Identity $_.user | Where-Object -Property Name -Ne -Value 'Domain Users' | Select-Object name | Out-File 'C:\Location\Starters-Leavers\adgroups.txt'
    Write-Host
    
    Write-Host -ForegroundColor Yellow "Removing the leaver from all AD groups except Domain Users"
    Get-AdPrincipalGroupMembership -Identity $_.user | Where-Object -Property Name -Ne -Value 'Domain Users' | Remove-AdGroupMember -Members $_.user -Confirm:$false
    Write-Host
    Write-Host -ForegroundColor Green "Removed from all AD groups except Domain Users"
    Write-Host

    Write-Host -ForegroundColor Yellow "Disabling user account on AD"
    Write-Host
    Disable-ADAccount -Identity $adacct.SamAccountName
    Write-Host -ForegroundColor Green "Disabled AD account"
    Write-Host
 
    Write-Host -ForegroundColor Yellow "Changing AD Password to Random Password"
    Write-Host
    $Password = -join ((48..122) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
    $PwdSecStr = ConvertTo-SecureString $Password -AsPlainText -Force
    Set-ADAccountPassword -Identity $adacct.SamAccountName -NewPassword $PwdSecStr -Reset
    Write-Host -ForegroundColor Green "Password changed for $($adacct.Name)"
    Write-Host
 
    ###### PART 2 ######
    #################### 

    #Get AD user details again as the user has moved OU

    $adacct = Get-ADUser $_.user
    $ticket = $_.ticket
 
    #Disable mailbox, move user to Leavers OU
 
    Write-Host -ForegroundColor Yellow "Disabling Remote Mailbox"
    Write-Host
    Disable-OnpremRemoteMailbox -Identity $adacct.SamAccountName -Confirm:$false
    Write-Host -ForegroundColor Green "Remote Mailbox disabled"
    Write-Host
    Write-Host -ForegroundColor Yellow "Now moving user to Leavers AD OU"
    Write-Host
    Move-ADObject -Identity $adacct.DistinguishedName -TargetPath $PermLeaversOU
    Write-Host -ForegroundColor Green "Moved to Leavers OU"
    Write-Host

    $report1 = $adacct | Select-Object Enabled | Out-String
 
    Write-Host -ForegroundColor Yellow "Generating and sending user status report to IT Support"
 
    #Sends SMTP email via o365 smtp relay

    $sendMailMessageSplat = @{

        Subject    = "[#INC-$($_.ticket)]"
        From       = 'LeaverPSScriptreport@domain.com'
        To         = 'support@domain.com'
        SmtpServer = 'relay server'
        Body       = $report + $report1 + $body
        Attachment = 'C:\Location\Starters-Leavers\adgroups.txt'
    }
    Send-MailMessage @sendMailMessageSplat
 
}

Remove-Item -Path 'C:\Location\Starters-Leavers\adgroups.txt'

Write-Host
Write-Host 'Press any key to exit.';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');