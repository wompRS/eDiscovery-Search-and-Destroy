## This script will authenticate with the Exchange Compliance Center, start a search based on your specifications, and then delete the emails associated.
## Take extra care when running this script. 
## Reed Sutherland
## August 4, 2021


## Set your username here. The username must have permissions for eDiscovery in the Compliance Center

# Obsolete $Email = Write-Host "Please enter your Email Address" && Read-Host  ## The username syntax is your complex email address. 

# Function for Validating Email
# Credit https://stackoverflow.com/users/615422/vertigoray

function ValidateEmail {
   param(
       [Parameter(Mandatory = $true)]
       [ValidateScript({ Resolve-DnsName -Name $_.Host -Type 'MX' })]
       [mailaddress]
       $Email
   )
   Write-Output $From
}

ValidateEmail



## Create the initial connection to the Compliance Center. An authentication prompt will pop up for your input.
## The most common error in this step is WIN-RM being disabled.

Connect-IPPSSession -UserPrincipalName $username
 
<#
New-ComplianceCase
   -Name == The name of the Compliance Case.
   -CaseType == eDiscovery, AdvancedEdiscovery.
   -Confirm] == Creates a confirmation prompt for continuing.
   -Description == Description of the Compliance Search.
   -ExternalId  == Case number for litigation, etc. Optional.
#>

## Set the Compliance Search Case variables

$case = '<ENTER NAME HERE>' ## This case name must be unique and will be used to create the compliance search.
$description = '<ENTER DESCRIPTION>' ## This is the description of the case & search for purposes of accounting. 

## Create the new Compliance Case using the variables set. You will be prompted to continue.

New-ComplianceCase -Name $case -CaseType Ediscovery -Confirm


<# Set the search requirements for the email as required. For more info see: https://docs.microsoft.com/en-us/powershell/module/exchange/new-compliancesearch?view=exchange-ps

New-ComplianceSearch
    -Name == Unique name of the Compliance Search.
    -AllowNotFoundExchangeLocationsEnabled == :$true or :$false -- This enables the inclusion of irregular mailboxes.
    -Case == The case name set in $case.
    -Confirm == Prompts to continue.
    -ContentMatchQuery == Sets the content of the email you are searching for.
    -Description == Sets the description of the search.
    -ExchangeLocation  == Sets the mailbox to include. 
    -Force == Suppress warning and confirmation messages.
    -HoldNames == Searches the locations that have been set to hold.
    -IncludeUserAppContent == :$true or :$false -- Sets the search to look in locations where a user does not have 0365 accounts in your organization
    -PublicFolderLocation == If you want Public Folders searched, select "ALL"
#>

## Create your search variables
$subject = "Subject:'<ENTER SUBJECT>'" ## COMMENT OUT if not needed. Used if searching for the subject line.
$content = "<ENTER KEYWORD STRIG>" ## COMMENT OUT if not needed. Used to search for keywords in the email content.
$mailbox = "<ENTER MAILBOX>" ## Set to 'all' or a specific user's mailbox.
$query = "$subject + ' AND ' + $content"
Write-Host $query

try {
Get-Variable subject -Scope Global -ErrorAction 'Stop'
} catch [System.Management.Automation.ItemNotFoundException]{
   Write-Warning $_;
   New-Variable -Name 
   }


New-ComplianceSearch -Name $case -ExchangeLocation all -ContentMatchQuery $query -Confirm -Description $description

<# 
    After creating the search, it is in the "not started" state. You must enter it into the started state with
    the "Start-ComplianceSearch" command. 

    Start-ComplianceSearch
     -Identity == The $name variable
     -Confirm == Created a prompt to continue
     -Force == Ignores prompts
     -RetryOnError == Retries if an error occurs. Not recommended. 

#>

Start-ComplianceSearch -Identity $name -Confirm | Format-List

<#
Get-ComplianceSearch
   [[-Identity] <ComplianceSearchIdParameter>]
   [-Case <String>]
   [-DomainController <Fqdn>]
   [-ResultSize <Unlimited>]
   [<CommonParameters>]
#>

Get-ComplianceSearch -Case $case -Identity $name -ResultSize unlimited | Format-List


<#

New-ComplianceSearchAction
   [-SearchName] <String[]>
   [-Export]
   [-ActionName <String>]
   [-ArchiveFormat <ComplianceExportArchiveFormat>]
   [-Confirm]
   [-FileTypeExclusionsForUnindexedItems <String[]>]
   [-EnableDedupe <Boolean>]
   [-ExchangeArchiveFormat <ComplianceExportArchiveFormat>]
   [-Force]
   [-Format <ComplianceDataTransferFormat>]
   [-IncludeCredential]
   [-IncludeSharePointDocumentVersions <Boolean>]
   [-JobOptions <Int32>]
   [-NotifyEmail <String>]
   [-NotifyEmailCC <String>]
   [-ReferenceActionName <String>]
   [-Region <String>]
   [-Report]
   [-RetentionReport]
   [-RetryOnError]
   [-Scenario <ComplianceSearchActionScenario>]
   [-Scope <ComplianceExportScope>]
   [-SearchNames <String[]>]
   [-SharePointArchiveFormat <ComplianceExportArchiveFormat>]
   [-ShareRootPath <String>]
   [-Version <String>]
   [-WhatIf]
   [<CommonParameters>]
#>


New-ComplianceSearchAction -SearchName "$name" -Purge -PurgeType SoftDelete -Confirm