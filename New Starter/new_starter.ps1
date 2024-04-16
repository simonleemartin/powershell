Get-PSSession | Remove-PSSession

#Set the title of the window.

$host.ui.RawUI.WindowTitle = "New User Creation Script"

#Starts a transcript of the new user creation script

$transcript = 'C:\location\transcript.txt'
Start-Transcript -Path $transcript | Out-Null

#Prompts for Office 365 login credentials and logs into Office 365

Write-Host
Write-Host -ForegroundColor Yellow "Enter your Office 365 Credentials"
Write-Host
$CloudCredential = Get-Credential $null

Write-Host
Write-Host -ForegroundColor Yellow "Connecting to Office 365 and on-prem Exchange"
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

#Give the user a quick description

Write-Host 
Write-Host 
Write-Host "This script will create a new AD user and move them to the correct OU."
Write-Host 
Write-Host

#Ask which company the new user works for to determine several attributes

function Show-Menu
{
param (
[string]$Title = 'What company does the user work for?'
)
    Write-Host
    Write-Host "$Title"
    Write-Host
    Write-Host "Press '1' for Company 1."
    Write-Host "Press '2' for Company 2."
    Write-Host "Press '3' for Company 3."
    Write-Host "Press '4' for Company 4."
    Write-Host "Press '5' for Company 5."
    Write-Host "Press 'Q' to quit."
    Write-Host
}

Do {
Show-Menu
$input = Read-Host "Please make a selection"
switch ($input) {
      '1' {
        $domain = 'domain1.com'
        $website = 'www.company1.com'
        $company = 'Company 1'
        $emailcompany = 'Company 1'
        $sso = "sso"
        Write-Host
        Write-Host -ForegroundColor Green "You have selected Company 1"
        Write-Host
    } '2' {
        $domain = 'domain2.com'
        $company = 'Company 2'
        $emailcompany = 'Company 2'
        $website = 'www.company2.com'
        Write-Host
        Write-Host -ForegroundColor Green "You have selected Company 2" 
        Write-Host
    } '3' {
        $domain = 'domain3.com'
        $company = 'Company 3'
        $emailcompany = 'Company 3'
        $website = 'www.company3.com'
        Write-Host
        Write-Host -ForegroundColor Green "You have selected Company 3"
        Write-Host
    } '4' {
        $domain = 'domain4.co.uk'
        $company = 'Company 4'
        $emailcompany = 'Company 4'
        $website = 'www.company4.co.uk'
        Write-Host
        Write-Host -ForegroundColor Green "You have selected Company 4"
        Write-Host
    } '5' {
        $domain = 'domain5.com'
        $company = 'Company 5'
        $emailcompany = 'Company 5'
        $website = 'www.company5.com'
        Write-Host
        Write-Host -ForegroundColor Green "You have selected Company 5"
        Write-Host
    } 'Q' {
    Write-Host
    Write-Host 'Exiting Script' -ForegroundColor Yellow
    Start-Sleep -Seconds 1
    Exit
    }
    default {
        Write-Host ('{0}Please enter a valid option{0}' -f [environment]::NewLine) -ForegroundColor Red
        }    
}
} While ($domain -eq $null)

#Create new remote mailbox

$firstname = Read-Host "What's the user's first name?"
Write-Host
$lastname = Read-Host "What's the user's last name?"
Write-Host
$upn = Read-Host "What will the username be?"
Write-Host
$password = Read-Host "Enter the user's temporary password" -AsSecureString
Write-Host
$username="$($firstname) $($Lastname)"
$alias="$($firstname.ToLower()).$($lastname.ToLower())"
$smtpaddress = "$alias@$domain"

#Takes attributes from an existing user to add to the new user

$Properties = @(
    'company'
    'department'
    'facsimileTelephoneNumber'
    'physicalDeliveryOfficeName'
    'telephoneNumber'
    'postalCode'
    'streetAddress'
    'st'
    'c'
    'co'
    'l'
    'scriptpath'
    )


#Prepares the parameters for the new remote mailbox and creates the new remote mailbox

$newMbxParams = @{
    Name = $username
    Password = $password 
    UserPrincipalName = "$upn@$domain"
    Alias = $alias 
    DisplayName = $username
    FirstName = $firstname 
    LastName = $lastname 
    OnPremisesOrganizationalUnit = "AD/Azure/Organization/$Company"
    PrimarySmtpAddress = $smtpaddress
    RemoteRoutingAddress = $smtpaddress
}

try{
    
    New-OnpremRemoteMailbox @newMbxParams -ErrorAction Stop  | Out-Null
}
 catch [System.Management.Automation.RemoteException]
{
    Write-Host -ForegroundColor Red "An error has occurred. Please see the error below:"
    Write-Host
    $PSItem
    Write-Host
    Start-Sleep 3
    Write-Host -ForegroundColor Red "Exiting script"
    Start-Sleep 5
    Exit
}

Write-Host
Write-Host -ForegroundColor Green "Remote mailbox for $username has been created"
Write-Host

#Prepares the parameters needed to amend the new AD user

$ADUserParams = @{
    Country = "GB"
    HomePage = $website
    ChangePasswordAtLogon = $true
}

Start-Sleep 1

Get-ADUser -Identity $upn | Set-ADUser @ADUserParams | Out-Null -ErrorAction SilentlyContinue

Write-Host
Write-Host -ForegroundColor Green "$username has been created in AD"
Write-Host


#Prompts for the username of the existing user

Write-Host
Write-Host -ForegroundColor Yellow  "Do you want to copy attributes from another user?"
Write-Host

    $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost)
     { 
       Y {Write-Host
       $copyuser = Read-Host "What's the username of the person you want to copy?"

       try{
    
    Get-ADUser -Identity $copyuser -ErrorAction Stop  | Out-Null
}
 catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
{
    Write-Host -ForegroundColor Red "An error has occurred. Please see the error below:"
    Write-Host
    $PSItem
    Write-Host
    Start-Sleep 3
    Write-Host -ForegroundColor Red "Exiting script"
    Start-Sleep 5
    Exit
}

$UserProps = Get-ADUser -Identity $copyuser -Properties $Properties |
    Select-Object -Property $Properties

$Replace = @{}

foreach( $Prop in $Properties ){
    if( $UserProps.$Prop ){
        $Replace.Add( $Prop, $UserProps.$Prop )
        }
    }
       Get-ADUser -Identity $upn | Set-ADUser -Replace $Replace} 
       N {
       Write-Host
       Write-Host -ForegroundColor Yellow "Creating the user without copying attributes from another user"
       } 
     } 

#Adding the new user to standard AD user groups

Write-Host
Write-Host -ForegroundColor Yellow "Setting default group membership for the user"
Write-Host

Add-ADGroupMember -Identity "Default AD Group 1" -Members $upn
Add-ADGroupMember -Identity "Default AD Group 2" -Members $upn
Add-ADGroupMember -Identity "Default AD Group 3" -Members $upn
Add-ADGroupMember -Identity "Default AD Group 4" -Members $upn

#Uses the previously specified user to copy attributes to the new user



#Adds the new user to the SSO group but only if they're a Company 1 user

try {

Add-ADGroupMember -Identity $sso -Members $upn

}

catch {}

Write-Host -ForegroundColor Green "Default group membership set"
Write-Host

#A choice is provided asking what licese the user needs. This is set via AD security groups

function Show-Menu
{
param (
[string]$Title = 'What license does the user need?'
)

    Write-Host "$Title"
    Write-Host
    Write-Host "Press '1' for Exchange Online Plan 1."
    Write-Host "Press '2' for Office 365 E3."
    Write-Host "Press 'Q' to quit."
    Write-Host
}


Do {
Show-Menu
$input = Read-Host "Please make a selection"
switch ($input) {
      '1' {
        $exchangeplan1 = 'Exchange_Online_Plan_1_License'
Add-ADGroupMember -Identity $exchangeplan1 -Members $upn
        Write-Host
        Write-Host -ForegroundColor Green "You have selected Exchange Online Plan 1"
        Write-Host
    } '2' {
        $office365e3 = 'Office_365_E3_License'
Add-ADGroupMember -Identity $office365e3 -Members $upn
        Write-Host
        Write-Host -ForegroundColor Green "You have selected Office 365 E3"
        Write-Host
    } 'Q' {
    Write-Host 'Exiting Script' -ForegroundColor Yellow
    Start-Sleep -Seconds 1
    Exit
    }
    default {
        Write-Host ('{0}Please enter a valid option{0}' -f [environment]::NewLine) -ForegroundColor Red
        }    
}
} While ($domain -eq $null)

Write-Host
Write-Host -ForegroundColor Green "User license set"
Write-Host

#Synchronises all changes with Azure

Write-Host -ForegroundColor Yellow "Syncing Azure with AD"
Write-Host

Start-Sleep 2

$sazure = New-PSSession -ComputerName ServerAzure
Invoke-Command -Session $sazure -ScriptBlock {C:\location\AzureADDeltaSync.ps1} | Out-Null

#Checks to see if the user exists in 365 before proceeding to the next step

Do
{

  $Checkif365UserExists = Get-MsolUser -UserPrincipalName $upn@$domain -ErrorAction SilentlyContinue
  Write-Host
  Write-Host "Waiting for the user to be created in 365"

  Start-Sleep 10

}

While ($Checkif365UserExists -eq $null)

Write-Host
Write-Host -ForegroundColor Green "Sync Complete"
Write-Host

#Sets the usage location in 365 to the UK. This is important as it determines the data centre in which this users data is held

Set-MsolUser -UserPrincipalName $upn@$domain -UsageLocation "GB"

#Now checking to see if the mailbox is created and continues to check until it is created

Write-Host -ForegroundColor Yellow "Waiting for mailbox creation"
Write-Host

Do
{

  $CheckifMailboxCreated = Get-EXOMailbox -Identity $alias -ErrorAction SilentlyContinue
  Write-Host
  Write-Host "Waiting for the mailbox to be created"

  Start-Sleep 10

}

While ($CheckifMailboxCreated -eq $null)

Write-Host
Write-Host -ForegroundColor Green "Mailbox created"
Write-Host

#Disables the junk mail folder for users as we're using Mimecast

Write-Host -ForegroundColor Yellow "Disabling Junk Mail folder for user"
Write-Host

Get-EXOMailbox -Identity $alias | Set-MailboxJunkEmailConfiguration -Enabled $false
Write-Host
Write-Host -ForegroundColor Green "Junk Mail folder disabled"
Write-Host

#As the mailbox is created, a welcome email can now be sent to the user

Write-Host -ForegroundColor Yellow "Sending a welcome email to the new starter. This could take a while to send before the mailbox is provisioned. Don't close this window until it completes."
Write-Host

Start-Sleep 15

#This takes the current HTML file, copies it to a new file and changes attributes in the file so the welcome email is personalised

$htmltemplate = Get-Content C:\location\Starters-Leavers\welcome.html
$htmltemplate1 = ($htmltemplate) | ForEach-Object {
    $_ -replace 'firstname',"$firstname" `
       -replace 'lastname', "$lastname" `
       -replace 'emailaddress', "$alias@$domain" `
       -replace 'company', "$emailcompany" `
       -replace 'alias', "$alias" `
} | Set-Content -Path C:\location\Starters-Leavers\welcome1.html
$htmltemplate3 = Get-Content C:\location\Starters-Leavers\welcome1.html -Raw


#Sends SMTP email to the new user via Office 365 SMTP relay

    $sendMailMessageSplat = @{

        Subject    = "Welcome, $firstname $lastname!"
        From       = 'welcome@domain.com'
        To         = $smtpaddress
        SmtpServer = 'relay server'
        Body       = $htmltemplate3
    }

Do {
    Send-MailMessage @sendMailMessageSplat -BodyAsHtml -ErrorAction SilentlyContinue -ErrorVariable ev
    } Until (-not($ev))


Stop-Transcript | Out-Null

$newname = "$upn"+"_"+"_transcript.txt"

Rename-Item -Path $transcript -NewName $newname

$attachment = "C:\location\Starters-Leavers\$newname"

#Sends SMTP email to IT Support via Office 365 SMTP relay

$sendMailMessageSplatITSupport = @{

    Subject    = "Welcome, $firstname $lastname!"
    From       = 'welcome@domain.com'
    To         = 'support@domain.com'
    SmtpServer = 'relay server'
    Body       = $htmltemplate3
    Attachment = $attachment
}

Send-MailMessage @sendMailMessageSplatITSupport -BodyAsHtml

#Removes the amended HTML file

Remove-Item -Path C:\location\Starters-Leavers\welcome1.html

Remove-Item $attachment

Start-Sleep 2

#Sends SMTP email to IT Support asking for MFA to be enabled

$sendMailMessageSplatMFA = @{

    Subject    = "Enable MFA for $firstname $lastname"
    From       = 'mfa@domain.com'
    To         = 'support@domain.com'
    SmtpServer = 'relay server'
    Body       = "Enable MFA for $firstname $lastname. Once enabled, go through the steps to enable it with the user."
}

Send-MailMessage @sendMailMessageSplatMFA -BodyAsHtml

Write-Host
Write-Host -ForegroundColor Green "The welcome email has been sent"
Start-Sleep 2
Write-Host
Write-Host 'Press any key to exit.';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');