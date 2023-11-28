## This script will authenticate with the Exchange Compliance Center, start a search based on your specifications, and then delete the emails associated.
## Take extra care when running this script. 
## Reed Sutherland
## August 4, 2021


Set-ExecutionPolicy RemoteSigned -Scope Process # Enables execution of scripts in this Powershell session
cmd /c "winrm get winrm/config/client/auth" # Checks to see if WinRM is enabled // No error checking as of now
cmd /c "winrm set winrm/config/client/auth @{Basic="true"}" # Turns on WinRM
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Confirm:$false # Installs the Exchange Online module for compliance connections. Automatically accepts prompts.
Import-Module ExchangeOnlineManagement # Imports the Exchange Online module into the Powershell session for the rest of the script.

# Function for Validating Email
# Credit https://stackoverflow.com/users/615422/vertigoray

function ValidateEmail {
   # Names the function
   param(
      [Parameter(Mandatory = $true)] # Sets the prompt to require input
      [ValidateScript( { Resolve-DnsName -Name $_.Host -Type 'MX' })] # Validates the email address sytax
      [mailaddress] # Asks for the input of an email address
      $Email
   )
   Write-Output $Email # Sets the $Email variable based on the accepted and validated entry above
}

ValidateEmail # Runs the ValidateEmail function which prompts for input from the user and then continues the script if the address meets standard email syntax. If not, it asks again until it fits this syntax.


## Create the initial connection to the Compliance Center. An authentication prompt will pop up for your input.

Connect-IPPSSession -UserPrincipalName "$Email"
 
#
function SetCaseName {
   # Names the function
   param(
      [Parameter(Mandatory = $true)] # Sets the prompt to require input 
      [string]$CaseName # Sets the prompt to write out "CaseName"
   )
   Write-Output $CaseName #Writes the output to verify it was input. 
}
#

$CaseName = SetCaseName #Sets the $CaseName variable based on the accepted and validated entry above

New-ComplianceCase -Name $CaseName -CaseType Ediscovery # Calls Casename function to set the name of the case to be used and accounted for in the future


function CaseDescription {
   # Names the function
   param(
      [Parameter(Mandatory = $true)] # Sets the prompt to require input
      $CaseDescription # Sets the prompt to write out "CaseDescription"
   )
   Write-Output $CaseDescription # Sets the $CaseDescription variable based on the accepted and validated entry above
}

$CaseDescription = CaseDescription
CaseDescription # Calls CaseDescription function to set the description of the case to be used and accounted for in the future


function Subject {
   # Names the function
   param(
      [Parameter(Mandatory = $true)] # Sets the prompt to require input 
      $subject # Sets the prompt to write out "CaseName"
   )
   Write-Output $subject #Writes the output to verify it was input. 
}

function Content {
   # Names the function
   param(
      [Parameter(Mandatory = $true)] # Sets the prompt to require input 
      $content # Sets the prompt to write out "CaseName"
   )
   Write-Output $content #Writes the output to verify it was input. 
}
function Mailbox {
   # Names the function
   param(
      [Parameter(Mandatory = $true)] # Sets the prompt to require input
      [ValidateScript( { Resolve-DnsName -Name $_.Host -Type 'MX' })] # Validates the email address sytax
      [mailaddress] # Asks for the input of an email address
      $Email
   )
   Write-Output $mailbox #Writes the output to verify it was input. 
}
#Create a Case Name for the Compliance search function. 
function SetComplianceCaseName {
   # Names the function
   param(
      [Parameter(Mandatory = $true)] # Sets the prompt to require input 
      [string]$ComplianceCaseName # Sets the prompt to write out "CaseName"
   )
   Write-Output $ComplianceCaseName #Writes the output to verify it was input. 
}
#

$Case = SetCase #Sets the $CaseName variable based on the accepted and validated entry above


## Create your search variables
$subject = "Subject: '$(subject)'" ## COMMENT OUT if not needed. Used if searching for the subject line.
$content = "'$(content)'" ## COMMENT OUT if not needed. Used to search for keywords in the email content.
$mailbox = "$(mailbox)" ## Set to 'all' or a specific user's mailbox.
$query = "$subject" + ' AND ' + "$content"
Write-Host $query


New-ComplianceSearch -Name $ComplianceCaseName -ExchangeLocation all -ContentMatchQuery $query -Confirm -Description $description

Start-ComplianceSearch -Identity $name -Confirm | Format-List

Get-ComplianceSearch -Case $ComplianceCaseName -Identity $name -ResultSize unlimited | Format-List

New-ComplianceSearchAction -SearchName "$name" -Purge -PurgeType SoftDelete -Confirm
