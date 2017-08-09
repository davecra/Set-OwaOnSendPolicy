# CONFIGURING EXCHANGE 2016 (ON-PREM) FOR ONSENDADDINSENABLED POLICY
For On Premises installations of Exchange 2016, you will need to configure then per this guide in order to take full advantage of the Outlook Web Add-ins OnSend event, which is fully documented here:

  (https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins?product=outlook)

## Requirements
This is only supported in Exchange 2016 and also only supported in Cumulative Update 6 (or later). To get CU6, please follow this link:
  
  (https://www.microsoft.com/en-us/download/details.aspx?id=55520)
  
Next, you will need to update Active Directory so that the ONSENDADDINSENABLED Policy is available to be set. Here are the steps to do this.

1)	After you have completed the install of CU6, reboot
2)	After rebooting, run setup.exe from the CU6 folder/ISO or drive with the following command:

  Setup.exe /PrepareAD /IAcceptExchangeServerLicenseTerms

3)	After this is complete, reboot the server.

Finally, you need to enable scripting to access PowerShell remotely with the desired authentication scheme. For example, to enable it for Basic Authentication, you can use the following PowerShell command:

```powershell
Get-PowerShellVirtualDirectory | Set-PowerShellVirtualDirectory -BasicAuthentication $true
```

## Configuring the Policy
Once you have met the initial requirements, you can begin to apply the policy to your users/group accounts. A script has been made available in order to make this process much easier. If you prefer to to the steps manually, please see the following article:

   (https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins?product=outlook#installing-outlook-add-ins-that-use-on-send) 

Otherwise you can use the PowerShell script provided in this repository:

```powershell
Set-OwaOnSendPolicy [-user/-group] [-name:<user/group>]
                    [-policyName:<name>] [-server:<server>]
```

## Running the Script
For more information on running the script, open a PowerShell command prompt, and type the following command:

```powershell
Get-Help .\Set-OwaOnSendPolicy.ps1 -full
```

## Issues
If you have any issues, please leave a comment.
