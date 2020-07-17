# ***************************************************************************************************
# ***************************************************************************************************
#
#  Author       : Cary GARVIN
#  Credit       : Aaron Guilmette
#  Contact      : cary(at)garvin.tech
#  LinkedIn     : https://www.linkedin.com/in/cary-garvin
#  GitHub       : https://github.com/carygarvin/
#
#
#  Script Name  : Cleanup-ExchangeProxyAddresses.ps1
#  Version      : 1.0
#  Release date : 28/04/2018 (CET)
#  History      : The present script is inspired from the 'Remove-ExchangeProxyAddresses.ps1' (https://www.powershellgallery.com/packages/Remove-ExchangeProxyAddresses/4.1 or https://www.undocumented-features.com/2017/02/10/removing-proxy-addresses-from-exchange-recipients/) written by Aaron Guilmette, but on steroids...
#                 In addition to removing these CCMAIL, MSMAIL addresses from Mailbox, MailUser and MailContact objects, this present version will remove these addresses from ALL possible mail-enabled objects in any Exchange Organization.
#                 Additional objects are RemoteUserMailbox, DistributionGroups and MailPublicFolders.
#  Purpose      : The present script can be used for mail objects cleanup prior to any migration from any Exchange version above Exchange 2010 to a higher version or to Office 365/Exchange Online. It will cleanup obsolete CC:MAIL and MSMAIL, or more if you wish (specify which address types in '$UndesiredAddressTypes' array) from ALL mail-enabled object types.
#
#
#
#  To run the present Script, you need Exchange Admin privileges.Edit the **$UndesiredAddressTypes** array within the script to specify which types of Proxy Addresses need to be removed.
#


# Configurable parameters by IT Administrator as referred in the script's instructions. Specify in the '$UndesiredAddressTypes' array variable which Address types need to be removed..
$UndesiredAddressTypes = @("MS","CCMAIL")
# $UndesiredAddressTypes = @("MS","CCMAIL","X400","FAX","SIP")    # Further riskier examples


####################################################################################################
#                                       Script function                                            #
####################################################################################################



Function RemoveUndesiredAddresses
	{
	param
		(
		$ObjectType,
		$AddressTypesToRemove
		)
	
	Switch ($ObjectType)
		{
		Mailbox
			{
			Write-Host "Processing Mailbox objects" -foregroundcolor "yellow"
			$ObjectCollection = Get-Mailbox -Resultsize Unlimited | Where {$_.EmailAddresses -like "MS:*" -or $_.EmailAddresses -like "CCMAIL:*"}
			}
		RemoteUserMailbox
			{
			Write-Host "Processing Remote Mailbox objects" -foregroundcolor "yellow"
			$ObjectCollection = Get-RemoteMailbox -Resultsize Unlimited | Where {$_.EmailAddresses -like "MS:*" -or $_.EmailAddresses -like "CCMAIL:*"}
			}
		MailUser
			{
			Write-Host "Processing MailUser objects" -foregroundcolor "yellow"
			$ObjectCollection = Get-MailUser -Resultsize Unlimited | Where {$_.EmailAddresses -like "MS:*" -or $_.EmailAddresses -like "CCMAIL:*"}
			}
		MailContact
			{
			Write-Host "Processing Contact objects" -foregroundcolor "yellow"
			$ObjectCollection = Get-MailContact -Resultsize Unlimited | Where {$_.EmailAddresses -like "MS:*" -or $_.EmailAddresses -like "CCMAIL:*"}
			}
		DistributionGroups # (The scope will be "Distribution Groups" and "Mail-Enabled Security Groups")
			{
			Write-Host "Processing Distribution Groups objects" -foregroundcolor "yellow"
			$ObjectCollection = Get-DistributionGroup -Resultsize Unlimited | Where {$_.EmailAddresses -like "MS:*" -or $_.EmailAddresses -like "CCMAIL:*"}
			}
		MailPublicFolders
			{
			Write-Host "Processing Mail enabled Public Folders objects" -foregroundcolor "yellow"
			$ObjectCollection = Get-MailPublicFolder -Resultsize Unlimited | Where {$_.EmailAddresses -like "MS:*" -or $_.EmailAddresses -like "CCMAIL:*"}
			}
		}
		
	If ($ObjectCollection)
		{
		write-host "A total of $($($ObjectCollection.length)) $ObjectType objects have been identified with MS MAIL or cc:MAIL Addresses." -foregroundcolor "magenta"
		ForEach ($object in $ObjectCollection)
			{
			$Error.Clear()
			Write-Host "`tProcessing $ObjectType '$($object.DisplayName)' ($($object.PrimarySMTPAddress))..."
			
			For ($i = ($object.EmailAddresses.count) - 1; $i -ge 0; $i--)
				{
				Foreach ($AddressTypeToRemove in $AddressTypesToRemove)
					{
					$Address = $object.EmailAddresses[$i]

					If ($Address.Prefix.PrimaryPrefix -eq $AddressTypeToRemove)
						{
						Write-Host "`t`tRemoving Proxy Address '$Address'" -ForegroundColor "white"
						try {
							$object.EmailAddresses.removeat($i)
							Switch ($ObjectType)
								{
								Mailbox
									{$object | Set-Mailbox -EmailAddresses $object.EmailAddresses}
								RemoteMailbox
									{$object | Set-RemoteMailbox -EmailAddresses $object.EmailAddresses}
								MailUser
									{$object | Set-Mailuser -EmailAddresses $object.EmailAddresses}
								MailContact
									{$object | Set-MailContact -EmailAddresses $object.EmailAddresses}
								DistributionGroups
									{$object | Set-DistributionGroup -EmailAddresses $object.EmailAddresses}
								MailPublicFolders
									{$object | Set-MailPublicFolder -EmailAddresses $object.EmailAddresses}
								}
							Write-Host "`t`tProxy Address '$Address' sucessfully removed!" -foregroundcolor "green"
							}
						catch {Write-Host "`t`tProxy Address '$Address' could not be removed!" -foregroundcolor "red"}
						}
					}
				}
			If ($Error)	{$Error | out-file "$($script:ScriptPath)\$($script:ExecutionTimeStamp)_$($script:ScriptName)_$($ObjectType)_errors.txt" -Append -NoClobber}
			}
		}
	Else {write-host "`tNone found!" -foregroundcolor "green"}
	write-host
	}





####################################################################################################
#                                          Script Main                                             #
####################################################################################################



write-host "Proxy Addresses of the following types will be removed:`r`n" -foregroundcolor "gray"
$UndesiredAddressTypes | ForEach {Write-Output $_}
$ProxyAddressesToRemoveAreOK = "init"
While("yes","no" -notcontains $ProxyAddressesToRemoveAreOK)
	{
	write-host "`r`nAre the above Proxy Address types OK to remove by executing the present script - [yes] or [no]? " -NoNewLine -foregroundcolor "gray"
	$ProxyAddressesToRemoveAreOK = read-host
	}
If ($ProxyAddressesToRemoveAreOK -eq "no")
	{
	write-host "Exiting script to allow list of Proxy Address types at the top of the script to be amended."
	Break
	}


$ExchangeSnapin = Get-PSSnapin -Registered Microsoft.Exchange.Management.PowerShell.E* -ErrorAction 'SilentlyContinue'
If ($ExchangeSnapin -eq $null)
	{
	Write-Warning "Exchange Snapin not available. Aborting script..." -foregroundcolor "red"
	Break
	}
Else
	{
	If ((Get-Command "Get-Mailbox*") -eq $null)
		{
		Write-Host "Loading Exchange Spnapin..."
		Add-PSSnapin $ExchangeSnapin -ErrorAction 'SilentlyContinue'
		}
	}


$script:ScriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$script:ScriptName = (Get-Item $MyInvocation.MyCommand).Basename
$script:ExecutionTimeStamp = get-date -format "yyyy-MM-dd_hh-mm-ss"

Start-Transcript -Path "$($script:ScriptPath)\$($script:ExecutionTimeStamp)_$($script:ScriptName).log" -NoClobber | out-null

[array]$RecipientTypes = @('Mailbox', 'MailUser', 'MailContact', 'DistributionGroups', 'MailPublicFolders')
Foreach ($RecipientType in $RecipientTypes) {RemoveUndesiredAddresses $RecipientType $UndesiredAddressTypes}

Stop-Transcript  | out-null




