<#
    .SYNOPSIS
    Set mailbox ExternalOofOptions to 'External' for members of a dedicated security group.
       
    Thomas Stensitzki, TSC
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.2
	
    .DESCRIPTION
    This script sets the mailbox ExternalOofOptions to 'External' for members of a given security group.
    ExternalOofOptions for users that are NOT a member of the security group will be set to 'InternalOnly'.
    Controlling the ExternalOofOptions has been implemented follow compliance rules.

    Based an GRP2CAS 1.2

    .NOTES 
    Requirements 
    - Exchange Management Shell (EMS) 2013+
    - GlobalFunctions library as described here: http://scripts.granikos.eu
 
    Revision History 
    -------------------------------------------------------------------------------- 
    1.2 Inital community release
   
    .PARAMETER ADGroup
    Defines the Active Directory security group which holds the users allowed for external OOF. If user is part of the group ExternalOofOptions would be set to 'External'

    .PARAMETER OrganizationalUnit
    OU for filtering user objects
    
    .PARAMETER RemoveRights
    Switch to control, if ExternalOofOptions rights should be removed

    .EXAMPLE
    .\Set-ExternalOOF.ps1
    Run script with default settings
#>

[CmdletBinding()]
Param(
    [string]$ADGroup = 'External_OOF_Users',
    [string]$OrganizationalUnit = 'mcsmemail.de/DE/Users',
    [switch]$RemoveRights
)

# Import global functions
Import-Module -Name GlobalFunctions
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name

# Create a log folder
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started')

# Variables
[string[]]$GroupUsers = @()
$NeedToExecute = $false
$SetAudience = $false

# Load all group members
$group = Get-Group $ADGroup

if ($group -eq $null) {
  # Group not found
  $logger.Write(('Error on loading group {0}. Script aborted.' -f ($ADGroup)))
  $logger.Write('Script finished') 
  Exit(1)
}
else {
  # Fetch all group members
  foreach ($user in $group.members) {
    $GroupUsers += $($user.DistinguishedName)
  }
}

# Getting all mailboxes
$logger.Write('Loading mailboxes...')

# Fetch all users from $OrganizationalUnit, but use some proprietary filter, adjst as needed

$mblist = Get-Mailbox  -OrganizationalUnit $OrganizationalUnit -ResultSize unlimited | Where-Object{(($_.OrganizationalUnit -eq 'mcsmemail.de/DE/Users/SpecialUsers') -or ($_.OrganizationalUnit -like 'mcsmemail.de/DE/Users/Windows10Users*'))}

foreach ($mb in $mblist ) {
  # Check mailbox
  [string]$dn = $mb.DistinguishedName.ToString()
				
  # User is a group member, to be set to 'External'
  if ($GroupUsers.Contains($dn)) {

    if ($mb.ExternalOofOptions -ne 'External') {
      $logger.Write(("{0} ExternalOofOptions is not set to External. Setting ExternalOofOptions to 'External'." -f $mb.DisplayName))
      $NeedToExecute = $true
      $SetAudience = $false  
    }
    else {
      # Nothing to do
      $NeedToExecute = $false
      $SetAudience = $false
    }

    # Set Invoke-Command command string
    $Command = "Set-Mailbox -Identity $($mb.PrimarySmtpAddress) -ExternalOofOptions 'External'"
  }
  else {
    if ($mb.ExternalOofOptions -ne 'InternalOnly')
    {
      If ($RemoveRights) {
        $logger.Write("$($mb.DisplayName) is not set to 'InternalOnly'. RemoveRights is enabled, ExternalOofOptions will be set To 'InternalOnly'. External OOF message will be cleared.")
        $NeedToExecute = $true
        $SetAudience = $true
      }
    }
    else {
      # Nothing to do
      $NeedToExecute = $false
      $SetAudience = $false
    }

    # Set Invoke-Command command string 
    $Command = "Set-Mailbox -Identity $($mb.PrimarySmtpAddress) -ExternalOofOptions 'InternalOnly'"
    $AudienceCommand = "Set-MailboxAutoReplyConfiguration -Identity $($mb.PrimarySmtpAddress) -ExternalAudience 'None' -ExternalMessage " +'$null'
  }
  
  if ($NeedToExecute) {
    # Set ExternalOofOptions
    Invoke-Expression -Command $Command
  }

  if ($SetAudience) {
    # Set audience
    Invoke-Expression -Command $AudienceCommand
  }
}
# Done
$logger.Write('Script finished')