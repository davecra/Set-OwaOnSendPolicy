<#
.SYNOPSIS
This script will setup the OnSend policy in Exchange for OWA users or groups so 
that add-ins which use the OnSend event can be installed. It simplifies the steps
outlined in the following article:

    https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins?product=outlook#installing-outlook-add-ins-that-use-on-send 

NOTE: This is only supported in Exchange 2016.
NOTE: Before you can run this script, you must have performed these steps first:

1) Exchange 2016 CU6 must be installed. See: https://www.microsoft.com/en-us/download/details.aspx?id=55520

2) After setup is complete you must upgrade Active Directory using this command: 

    Setup.exe /PrepareAD /IAcceptExchangeServerLicenseTerms

    NOTE: See: https://technet.microsoft.com/en-us/library/bb125224(v=exchg.160).aspx 

3) Finally, if it has not already been configured, you must configure PowerShell to work against the server remotely:

    Get-PowerShellVirtualDirectory | Set-PowerShellVirtualDirectory -BasicAuthentication $true

    NOTE: The above sample applied to BasicAuthentication. And the script below uses BasicAuthentication. If your
          authentication method is different you will need to change this script.
    NOTE: See: https://technet.microsoft.com/en-us/library/dd298108(v=exchg.160).aspx

.DESCRIPTION
Command Line Help:

Usage: Set-OwaOnSendPolicy [-user/-group] [-name:<user/group>]
                           [-policyName:<name>] [-server:<server>]

[-user/-group]: You can specify where to set a policy for a user or a group, but 
                you cannot specify not both
[-name]:        This is the email address of the user or the name of the group 
                you want to set the policy for
[-policyName]:  This is the name of the policy. If you do not include it will 
                default to: OWAOnSendAddinAllUserPolicy
[-server]:      (ONLY ON-PREM) The name of your server. If this is not specified, 
                Office 365 is assumed.

NOTE: If you do not specify anything on the command line it will run in step-by-step 
mode asking you for each value.

.PARAMETER name
This is the email address of the user or the name of the group you want to set the policy for

.PARAMETER policyName
This is the name of the policy. If you do not include it will default to: OWAOnSendAddinAllUserPolicy

.PARAMETER server
(ONLY ON-PREM) The name of your server. If this is not specified, Office 365 is assumed.

.PARAMETER user
You can specify where to set a policy for a user or a group, but  you cannot specify not both

.PARAMETER group 
You can specify where to set a policy for a user or a group, but you cannot specify not both

.LINK
Policy: https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins?product=outlook#installing-outlook-add-ins-that-use-on-send 
CU6 Download: https://www.microsoft.com/en-us/download/details.aspx?id=55520 
PrepareAD: https://technet.microsoft.com/en-us/library/bb125224(v=exchg.160).aspx
PowerShell: https://technet.microsoft.com/en-us/library/dd298108(v=exchg.160).aspx 

.EXAMPLE
Set-OwaOnSendPolicy -name AllOWAUsers -group -server:on-prem-server1.exchange.contoso.com 

To setup an on-prem server for only OWA users.
NOTE: This assumes there is a group called AllOWAUsers.
NOTE: Defaults to the policy name of OWAOnSendAddinAllUserPolicy

.EXAMPLE
Set-OwaOnSendPolicy -name:user@contoso.com -user

To setup Office365 for a single user.
NOTE: Defaults to the policy name of OWAOnSendAddinAllUserPolicy

.EXAMPLE
Set-OwaOnSendPolicy

To run in fully interactive/wizard mode, where you will be prompted to enter each piece of required information:
1) User or group
2) User or group name
3) Policy Name
4) Log-in info
5) Office 365 or On-Prem
6) If on-prem, the server name

.INPUTS
None. You cannot pipe content into this script.

.OUTPUTS
System.String. Displayed output for each completed step with success or failure with error.
#>
Param (
    [string]$name,
    [string]$policyName = 'OWAOnSendAddinAllUserPolicy',
    [string]$server,
    [switch]$user,
    [switch]$group
)
#Clear-Host
#23456789~123456789~123456789~123456789~123456789~123456789~123456789~1234567890
Write-Host -ForegroundColor:DarkGreen @" 
################################################################################
##                       Set-OwaOnSendPolicy                                  ##
##                                                                            ##
## Created by:                                                                ##
##      - David E. Craig                                                      ##
##      - Blaine Mathena                                                      ##
## - Version 1.0.2a                                                            ##
## - September 5, 2017 12:49PM EST                                            ##
## - http://theofficecontext.com                                              ##
##                                                                            ##
## Usage:                                                                     ##
##                                                                            ## 
##     Set-OwaOnSendPolicy [-user/-group] [-name:<user/group>]                ## 
##                         [-policyName:<name>] [-server:<server>]            ## 
##                                                                            ##
## Or, for topic help:                                                        ##
##                                                                            ##
##     Get-Help Set-OwaOnSendPolicy -full                                     ##
##                                                                            ##
################################################################################
"@ 
# validate
if((($user -eq $True) -or ($group -eq $True)) -and ((-not $name) -or (-not $policyName))) {
    throw 'Paramaters are invalid. Please see [Get-Help Set-OwaOnSendPolicy.ps1] for more assistance.'
}

if(($user -eq $True) -and ($group -eq $True)) {
    throw 'Parameters are invalid. You can only specify a user or group, not both. Please see [Get-Help Set-OwaOnSendPolicy.ps1] for more assistance.'
}

# if the user did not specify any parameters then we enter the wizard mode
if(($user -eq $False) -and ($group -eq $False)) {
    Write-Host 'No parameters defined. This shell script will run in step-by-step mode to set '
    Write-Host 'the OWA OnSend Policy for a specific account or it will install it for a particular group.'
    # Ask the user to proceed
    $Answer = Read-Host -Prompt 'Proceed? (Y/n)'
    # If the user answered Y, then proceed
    if (($Answer -eq 'n') -or ($Answer -eq 'N')) {
        Write-Host 'Exited. No changes made.'
        exit
    }
    # Ask the user if they want to install from
    # Ask the user if it is a group or a user
    $Answer = Read-Host -Prompt 'Install for user or group? (U/g)'
    if (($Answer -eq 'u') -or ($Answer -eq 'U')) {
        $user = $True
        $name = Read-Host -Prompt 'What is the user name?'
    } else {
        $group = $True
        $name = Read-Host -Prompt 'What is the group name?'
    }
    # Ask the user for the policy name
    $policyName = Read-Host -Prompt 'What do you want to name the policy, for example: OWAOnSendAddinAllUserPolicy. To choose default, just press enter.'
    if(-not $policyName) {
        $policyName = 'OWAOnSendAddinAllUserPolicy'
    }
}
# Ask use of this is on-prem of O365
if(-not $server)
{
    # No server specified in the command line so we will prompt the user
    $URLRequest = Read-Host -Prompt 'Is this Exchange Online (Office365) (y/n)?'
    If ($URLRequest -eq 'y')
    {
        # Office 365
        $URL = "https://outlook.office365.com/powershell-liveid/"
    }
    else {
        # On-Prem
        $URL = Read-Host -Prompt 'What is the name (or IP) of your exchange server? ex. exserver01.domain.com, or exserver01 or 192.168.12.11'
        $URL = 'https://' + $URL + '/powershell/'
    } 
} else {
    $URL = 'https://' + $server + '/powershell/'
}
Write-Host 'Authentication Required:' -ForegroundColor:Yellow
Write-Host 'You must sign into the Exchange server with an administrator account. Please supply your credentials.' -ForegroundColor:Yellow
$UserCredential = Get-Credential
# connect to the Exchange server
Write-Host 'Connecting to PowerShell session...'
$session_options = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Credential $UserCredential -Authentication Basic -AllowRedirection -ErrorVariable ProcessError -SessionOption $session_options 
if($ProcessError) {
    Write-Host -ForegroundColor:DarkYellow @"
There was an error trying to get the PowerShell sesssion started. This might be because you have not
enabled PowerShell scripting for the server. Run this command on your server:

    Get-PowerShellVirtualDirectory | Set-PowerShellVirtualDirectory -BasicAuthentication `$true

Please see: 
    https://technet.microsoft.com/en-us/library/dd298108(v=exchg.160).aspx
"@
    exit
}
Import-PSSession $Session -ErrorAction Stop 
$exists = Get-OWAMailboxPolicy $policyName
if(-not $exists) {
    Write-Host 'Creating the ' + $policyName + 'policy...'
    New-OWAMailboxPolicy $policyName -ErrorAction Stop
} else {
    Write-Host 'The' + $policyName + ' policy already exists.'
}
Write-Host 'Configuring the ' + $policyName + ' policy...'
Get-OWAMailboxPolicy $policyName | Set-OWAMailboxPolicy -OnSendAddinsEnabled:$true -ErrorVariable ProcessError
if($ProcessError) {
    Write-Host -ForegroundColor:DarkYellow @"
There was an error trying to set the Exchange policy. This might have occurred because
you have not updated Active Directory. You will need to run CU6 setup again, but with
the PrepareAD switch:

    Setup.exe /PrepareAD /IAcceptExchangeServerLicenseTerms

Please see: 
    https://technet.microsoft.com/en-us/library/bb125224(v=exchg.160).aspx
"@
    exit
}
Write-Host 'Creating and configuration of ' + $policyName + ' policy complete.'
if($user -eq $True) {
    Write-Host 'Setting policy for user: ' + $name
    Get-User $name -Filter {RecipientTypeDetails -eq $'UserMailBox'}|Set-CASMailbox -OwaMailboxPolicy $policyName -ErrorAction Stop
    Write-Host 'Completed. The ' + $policyName + ' policy has been set for the user ' + $name
} else {
    Write-Host 'Setting policy for group ' + $name
    $targetUsers = Get-Group $name -ErrorAction Stop | Select-Object -ExpandProperty members -ErrorAction Stop
    Write-Host 'Affected users: ' + $targetUsers
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'} -ErrorAction Stop |Set-CASMailbox -OwaMailboxPolicy $policyName -ErrorAction Stop
    Write-Host 'Completed. The ' + $policyName + ' policy has been set for the group ' + $name
}
Write-Host 'Exiting'
exit
