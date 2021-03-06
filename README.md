# Manage Out-Of-Office (OOF) settings

The following two scripts have been developed as part of a solution to implement an compliant OOF conifguration.

The scripts are supposed to be executed in order.

* Remove-OOFRule.ps1

Find and delete any existing OOF rules in user mailboxes. This script is supposed to be executed once in preperation for the Set-ExternalOOF.ps1 script.

* Set-ExternalOOF.ps1

Allow External OOF only for members of a dedicated security group. This script is supposed to executed as a scheduled task.

# Remove-OOFRule.ps1

This script searches for OOF rules created by users using the Outlook rule-tab in the OOF assistant and delete exisiting rules.

## Description

In preparation to configure Out-Of-Office (OFF) settings for users, any existing rule needs to be deleted.

The script will use either an exisiting Exchange Server EWS library or the Managed EWS library installed using the default file path.

## Requirements

* Exchange Management Shell (EMS) 2013+
* GlobalFunctions library as described here: [http://scripts.granikos.eu](http://scripts.granikos.eu)
* Locally installed Exchange Web Services (EWS) Library: [https://www.microsoft.com/en-us/download/details.aspx?id=42951](https://www.microsoft.com/en-us/download/details.aspx?id=42951)

## Parameters

### Mailbox

User mailbox alias, when removing OOF rule from a single mailbox

### Delete

Switch to finally delete any exisiting OOF rules in the user mailbox

### DebugLog

Switch to write each processed mailbox to log file. Using this swith will blow up your log file.

## Examples

``` PowerShell
.\Remove-OOFRule
```

Find any existing OOF rule and write results to log file

``` PowerShell
.\Remove-OOFRule -Mailbox SomeUser@varunagroup.de -Delete
```

Find and delete any existing OOF rules for user SomeUser@varunagroup.de and write delete actions to log file

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery

Find the script at TechNet Gallery

* [https://gallery.technet.microsoft.com/Remove-Out-Of-Office-rules-3f740964](https://gallery.technet.microsoft.com/Remove-Out-Of-Office-rules-3f740964)

# Set-ExternalOOF.ps1

Set mailbox ExternalOofOptions to _External_ for members of a dedicated security group.

## Description

This script sets the mailbox ExternalOofOptions to 'External' for members of a given security group. ExternalOofOptions for users that are NOT a member of the security group will be set to 'InternalOnly'. Controlling the ExternalOofOptions has been implemented follow compliance rules.

## Requirements

* Exchange Management Shell (EMS) 2013+
* GlobalFunctions library as described here: http://scripts.granikos.eu

## Parameters

### ADGroup

Defines the Active Directory security group which holds the users allowed for external OOF. If user is part of the group ExternalOofOptions would be set to 'External'

### OrganizationalUnit

OU for filtering user objects

### RemoveRights

Switch to control, if ExternalOofOptions rights should be removed

## Examples

``` PowerShell
.\Set-ExternalOOF.ps1
```

Run script with default settings

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki, TSC

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)

## Additional Credits

- [Rhoderick Milne](https://blogs.technet.microsoft.com/mspfe/2015/07/22/using-exchange-ews-to-delete-corrupt-oof/)