# Cleanup-ExchangeProxyAddresses.ps1
PowerShell Script to remove undesired/obsolete Proxy Address types (CC:MAIL, MSMAIL, etc...) from ALL mail-enabled objects types in an Exchange Organization.

Author       : Cary GARVIN  
Credit       : Aaron Guilmette  
Contact      : cary(at)garvin.tech  
LinkedIn     : [https://www.linkedin.com/in/cary-garvin](https://www.linkedin.com/in/cary-garvin)  
GitHub       : [https://github.com/carygarvin/](https://github.com/carygarvin/)  


Script Name  : [Cleanup-ExchangeProxyAddresses.ps1](https://github.com/carygarvin/Cleanup-ExchangeProxyAddresses.ps1)  
Version      : 1.0  
Release date : 28/04/2018 (CET)  

History      : The present script is inspired from the 'Remove-ExchangeProxyAddresses.ps1' ([https://www.powershellgallery.com/packages/Remove-ExchangeProxyAddresses/4.1](https://www.powershellgallery.com/packages/Remove-ExchangeProxyAddresses/4.1) or [https://www.undocumented-features.com/2017/02/10/removing-proxy-addresses-from-exchange-recipients/](https://www.undocumented-features.com/2017/02/10/removing-proxy-addresses-from-exchange-recipients/)) written by Aaron Guilmette, but on steroids...  In addition to removing these CCMAIL, MSMAIL addresses from _Mailbox_, _MailUser_ and _MailContact_ objects, this present version will remove these addresses from ALL possible mail-enabled objects in any Exchange Organization. Additional objects are _RemoteUserMailbox_, _DistributionGroups_ and _MailPublicFolders_.  
				 
Purpose      : The present script can be used for mail objects cleanup prior to any migration from any Exchange version above Exchange 2010 to a higher version or to Office 365/Exchange Online. It will cleanup obsolete CC:MAIL and MSMAIL, or more if you wish (specify which address types in **$UndesiredAddressTypes** array) from ALL mail-enabled object types.  


# Script usage
To run the present Script, you need Exchange Admin privileges. Edit the **$UndesiredAddressTypes** array within the Script to specify which types of Proxy Addresses need to be removed.
