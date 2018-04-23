# Start-AADAttributeWriteback.ps1

This PowerShell script can be used for attribute writeback from Azure Active Directory (AAD) to the on-premises Active Directory (AD).
Only accounts which are synchronized by Azure AD Connect (ImmutabledID attribute present) are being processed. The following AAD attributes
will be written to the corosponding on-premises AD account:

- targetaddress
- msExchHomeServerName
- legacyExchangeDN 
- mailNickName

Once these attributes are present in the on-premises AD user accounts, GALSync can be used to synchronize these users to another
Active Directory as contact objects.

This will be very usefull if you have multiple Active Directory / Exchange Environments and some ADs are using Exchange online only.


Usage
========

	Start-AADAttributeWriteback.ps1 -Credential (Get-Credential)
