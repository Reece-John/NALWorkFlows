

# GetM365SubscriptionSKUs.ps1

<# Header Information **********************************************************
Name: GetM365TenantSubscriptionSKUs.ps1
Created By: Mike John
Created Date: 03/28/2021
Summary:
    Logs and returns list od Tenant's Subscription SKUs
Update History *****************************************************************
Updated By: Mike John
Updated Date: 03/28/2021
    Reason Updated: original version
#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=2 )][string]$tenantDomain
     ,[Parameter(Mandatory=$True,Position=3 )][string]$alChapterSchemaFilePathAndName
     ,[Parameter(Mandatory=$True,Position=4 )][string]$tenantDefaultsPageName
     ,[Parameter(Mandatory=$True,Position=5 )][int]$tenantDefaultsStartRow
     ,[Parameter(Mandatory=$True,Position=6 )][string]$roleBasedEmailsPageName
     ,[Parameter(Mandatory=$True,Position=7 )][int]$roleBasedEmailsStartRow
     ,[Parameter(Mandatory=$True,Position=8 )][string]$alChapterUsersPageName
     ,[Parameter(Mandatory=$True,Position=9)][int]$alChapterUsersStartRow
     ,[Parameter(Mandatory=$True,Position=10)][string]$masterLogFilePathAndName
     ,[Parameter(Mandatory=$True,Position=11)][bool]$justTestingOnly
    )
begin {}
process {

}
