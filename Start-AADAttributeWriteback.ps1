######################################################################
## (C) 2018 Michael Miklis (michaelmiklis.de)
##
##
## Filename:      Start-AADAttributeWriteback.ps1
##
## Version:       1.0
##
## Release:       Final
##
## Requirements:  -none-
##
## Description:   Synchronizes Exchange attributes from Exchange Online
##                user account to the on-premises AD account.
##                
##                That will enable the use of Microsoft Identity Manager
##                GALSync module.
##
## This script is provided 'AS-IS'.  The author does not provide
## any guarantee or warranty, stated or implied.  Use at your own
## risk. You are free to reproduce, copy & modify the code, but
## please give the author credit.
##
####################################################################

param (
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()][object]$Credential
)


Set-PSDebug -Strict
Set-StrictMode -Version latest
  
 
function Start-AADAttributeWriteback
{
    <#
        .SYNOPSIS
        Synchronize targetaddress, msExchHomeServerName, legacyExchangeDN 
        and mailNickName from AAD to AD
  
        .DESCRIPTION
        The Start-AADAttributeWriteback CMDlet gets the attributes targetaddress, 
        msExchHomeServerName, legacyExchangeDN and mailNickName and writes the
        attribute's values to the on-premises Active Directory account
  
        .PARAMETER Credential
        Credential to log on to Office 365 / Exchange online
  
        .EXAMPLE
        Start-AADAttributeWriteback -Credential (Get-Credential)
 
    #>

    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()][object]$Credential
    )

    # Connect to Office 365 
    Connect-MsolService -Credential $Credential

    # Connect to Exchange online PowerShell interface
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication  Basic -AllowRedirection

    # Import Exchange online PSSession
    Import-PSSession $Session

    # Loop for each Exchange online Mailbox / User
    foreach ($TenantUser in $(Get-Mailbox | Select LegacyExchangeDN,ServerLegacyDN,PrimarySMTPAddress, UserPrincipalName, Alias))
    {

        # Get AzureAD Account for ImmutableID
        try
        {
            $AzureADUser = Get-MSOLUser -UserPrincipalName $TenantUser.UserPrincipalName
        }
        catch
        {
            $AzureADUser = $null
            Continue
        }

        # Only process synchronized accounts
        if ($AzureADUser.ImmutableID -ne $null)
        {
            # Convert ImmutableID into GUID format
            [GUID]$decodedGUID = [system.convert]::frombase64string($AzureADUser.ImmutableID)
        
            # Search Active Directory user based on GUID
            try
            {
                $ADUser = Get-AdObject -Identity $decodedGUID
            }
            catch
            {
                write-warning ("No on-premise account found for Azure-AD user {0}" -f $TenantUser.UserPrincipalName)
                $ADUser = $null
            }

            # if ADUser was found continue
            if ($ADUser -ne $null)
            {
               Set-ADObject -Identity $decodedGUID -Replace @{targetaddress=("SMTP:"+$TenantUser.PrimarySmtpAddress)}
           
               Set-ADObject -Identity $decodedGUID -Replace @{msExchHomeServerName =$TenantUser.ServerLegacyDN}

               Set-ADObject -Identity $decodedGUID -Replace @{legacyExchangeDN =$TenantUser.LegacyExchangeDN}

               Set-ADObject -Identity $decodedGUID -Replace @{mailNickName =$TenantUser.Alias}
            }

        }
    }

    # Disconnect Exchange online PSSession
    Remove-PSSession $Session
}


Start-AADAttributeWriteback -Credential $Credential