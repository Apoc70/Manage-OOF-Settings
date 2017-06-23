<#
  .SYNOPSIS 
  This script searches for OOF rules created by users using the Outlook rule-tab in the OOF assistant and deletes exisiting OOF rules.

  Thomas Stensitzki, TSC
	
  THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
  RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
  Version 1.0  

  .DESCRIPTION 
  In preparation to configure Out-Of-Office (OFF) settings for users, any existing rule needs to be deleted.

  The script will use either an exisiting Exchange Server EWS library or the Managed EWS library installed using the default file path.

  The script is based on Rhoderick Milne's script
  - Blog: https://blogs.technet.microsoft.com/rmilne/
  - Script: https://blogs.technet.microsoft.com/mspfe/2015/07/22/using-exchange-ews-to-delete-corrupt-oof/
    
  .NOTES 
  Requirements 
  - Exchange Management Shell (EMS) 2013+
  - GlobalFunctions library as described here: http://scripts.granikos.eu
  - Locally installed Exchange Web Services (EWS) Library, https://www.microsoft.com/en-us/download/details.aspx?id=42951
    
  Revision History 
  -------------------------------------------------------------------------------- 
  1.0      Initial release 

  This PowerShell script has been developed using ISESteroids - www.powertheshell.com 

  .PARAMETER Mailbox
  User mailbox alias, when removing OOF rule from a single mailbox

  .PARAMETER Delete
  Switch to finally delete any exisiting OOF rules in the user mailbox

  .PARAMETER DebugLog
  Switch to write each processed mailbox to log file. Using this swith will blow up your log file.

  .EXAMPLE
  Remove-OOFRule 

  Find any existing OOF rule and write results to log file

  .EXAMPLE
  Remove-OOFRule -Delete

  Find and delete any existing OOF rules in all user mailboxes and write delete actions to log file

  .EXAMPLE
  Remove-OOFRule -Mailbox SomeUser@varunagroup.de -Delete

  Find and delete any existing OOF rules for user SomeUser@varunagroup.de and write delete actions to log file

#>

[CmdletBinding()]
Param(
  [string]$Mailbox,
  [switch]$Delete,
  [switch]$DebugLog
)

# Import global modules
Import-Module GlobalFunctions

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name

# Create a log folder
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started')

# Variables
[string]$ErrorActionPreference = 'SilentlyContinue'

try {
  if($env:ExchangeInstallPath -ne '') {
    # Use local Exchange install path, if available
    $dllpath = "$($env:ExchangeInstallPath)\bin\Microsoft.Exchange.WebServices.dll"
  }
  else {
    # Use EWS managed API install path
    $dllpath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
  }
  
  [void][Reflection.Assembly]::LoadFile($dllpath)
}
catch {
  # Ooops, we could not load the Managed EWS DLL
  $logger.Write('Error on loading the EWS dll. Please check the path or install the EWS Managed API!',1)
  $logger.Write('Script aborted')
  exit(1)
}


function Create-Service
{
  [CmdletBinding()]
  Param(
    $IdentityForService
  )
  try {
    #Create a service reference
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    $mailAddress = $IdentityForService.PrimarySmtpAddress.ToString()
    $Service.AutodiscoverUrl($mailAddress)
    $enumSmtpAddress = [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress

    #Set Impersonation
    $Service.ImpersonatedUserId =  New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId($enumSmtpAddress,$mailAddress) 
  }
  catch {
    # Oops, something went wrong
    $logger.Write("Failed creating the exchange web service for $($IdentityForService). Check mailbox name and impersonation rights.",1)
    return
  }
  return $Service
}

function New-RuleSearch
{
  [CmdletBinding()]
  Param(
    $Service,
    $MailboxToSearch
  )
  
  [bool]$RuleFound = $false
  
  # Limit the page size. If there are more than 100 results returned from the search, we'll page through them
  $pageSize=100
  $pageLimitOffset=0
  $getMoreItems=$true
  $itemCount=0
  $itemsProcessed = 0
  
  # Iterate to rule set
  while($getMoreItems) { 

    # Setup Basic EWS Properties for Message Search - Used to locate Hidden Forwarding Rule 
    $SearchFilterForwardRule         = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM", [Microsoft.Exchange.WebServices.Data.ContainmentMode]::Prefixed, [Microsoft.Exchange.WebServices.Data.ComparisonMode]::Exact) 
    $itemViewForwardRule             = New-object Microsoft.Exchange.WebServices.Data.ItemView($pageSize,$pageLimitOffset,[Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
    $itemViewForwardRule.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties, [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject) 
    $itemViewForwardRule.Traversal   = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated 

    # Properties for hidden OOF Rule
    $PID_TAG_RULE_MSG_PROVIDER1    = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65EB,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String) 
    $PID_TAG_RULE_ACTION = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E99,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)    

    # Property Set for OOF Rule
    $propertySetForwardRule1 = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties, $PID_TAG_RULE_MSG_PROVIDER1) 
    $rfRootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
    $rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$rfRootFolderID) 
    $findResults = $rfRootFolder.FindItems($SearchFilterForwardRule, $itemViewForwardRule) 
  
    If ($findResults.TotalCount -gt 1) { 
      Foreach ($item in $findResults.Items) { 
        $item.Load($propertySetForwardRule1) 

        # Check for Exchange OOF Ghost Rules only
        if (($item.Itemclass -eq 'IPM.Rule.Version2.Message') -and ($item.ExtendedProperties[0].Value -eq 'MSFT:TDX OOF Rules')) {
          $propertySetForwardRule1.Add($PID_TAG_RULE_ACTION)
          $item.Load($propertySetForwardRule1)
          if (($item.ExtendedProperties[1].Value -ne $null ) -and ($item.ExtendedProperties[1].Value[14] -eq "3")) {
          
            $RuleFound = $true
            $logger.Write(('Found a rule for Mailbox {0}.' -f ($MailboxToSearch)))
            
            if ($DebugLog) {
              $ActionString = [System.Text.Encoding]::ASCII.GetString($item.ExtendedProperties[1].Value)
              $logger.Write(('As String: {0}' -f ($ActionString)))
              $logger.Write(('Binary: {0}' -f $item.ExtendedProperties[1].Value))
            }
            
            <#            switch ($item.ExtendedProperties[1].Value[14]) {
                3 {$logger.Write("Rule type: OP_REPLY")}
                7 {$logger.Write("Rule type: OP_FORWARD")}
            }#>         
          }

          if ($Delete) {
            # Hard deleted mailbox rule
            try {
              $Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
              $logger.Write('Rule successfully DELETED')
            }
            catch {
              $logger.Write("Deletion for mailbox $($MailboxToSearch) failed",1)
            }
          }
        }
        $itemsProcessed++  
      }
    }
    
    if ($findResults.MoreAvailable -eq $false){$getMoreItems = $false}

    # If there are more items to get, update the page offset
    if ($getMoreItems){$pageLimitOffset += $pageSize}
  }

  # uncommented. this only for checking, because on large environments it would blow up the log unneccesary
  
  if ($DebugLog) {
    if (-not ($RuleFound)) {
    
      $logger.Write("No rule for mailbox $($MailboxToSearch) found")
      
    }
  }
}


## MAIN ########################################

if ($Mailbox.Length -gt 0) {

  # Only a single mailbox is given
  try {
    $Identity = Get-Mailbox $Mailbox -ErrorAction $ErrorActionPreference -WarningAction $ErrorActionPreference
  }
  catch {
    # Oops
    $logger.Write("Failed to load mailbox $($Mailbox)",1)
    $logger.Write('Script aborted')
    exit(1)
  }

  # Create the web service for mailbox access
  $Service = Create-Service -Identity $Identity
  
  if ($Service) {
    New-RuleSearch -Service $Service -MailboxToSearch $Identity
  }
  else {
    # Oops
    $logger.Write('Something got wrong with the web service',1)
    $logger.Write('Script aborted')
    exit(1)
  }
}
else {
  # We need to fetch all user mailboxes
  
  $Mailboxes = Get-Mailbox -ResultSize Unlimited -WarningAction $ErrorActionPreference -ErrorAction $ErrorActionPreference
  
  ForEach ($UserMailbox in $Mailboxes) {
    try {
      $Service = Create-Service -Identity $UserMailbox
    }
    catch {
      $logger.Write("Failed to load mailbox $($UserMailbox)",1)
      $logger.Write('More than one mailbox to load - CONTINUE')
    }
    
    if ($Service) {
      New-RuleSearch -Service $Service -MailboxToSearch $UserMailbox
    }
    else {
      $logger.Write('Something got wrong with the service',1)
      $logger.Write('More than one mailbox to load - CONTINUE')
    }
    
  }
}

$logger.Write('Script ended')
exit(0)