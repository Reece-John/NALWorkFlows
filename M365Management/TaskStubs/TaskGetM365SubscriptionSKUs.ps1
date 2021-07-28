<# Header Information **********************************************************
Name: GetM365SubscriptionSKUs.ps1
Created By: Mike John
Created Date: 03/27/2021
Summary:
    Task stub to run ManageM365TneantUsers.ps1
Update History *****************************************************************
Updated By: Mike John
Updated Date: 03/27/2021
    Reason Updated: original version
#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][ValidateSet("NAL")][string]$tenantAbbreviation
)
begin {}
process {

    
    # get administrator credentials
    [System.Management.Automation.PSCredential]$myCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation;

    # Assign values to variables

    .\M365SharePoint\DownLoadSharePointFile.ps1 tenantCredentials -myCredentials $myCredentials`
                                                             -SharePointSiteURL $SiteURL `
                                                             -SharePointFileRelativeURL $FileRelativeURL `
                                                             -LocalFileDownloadPath $DownloadPath `
                                                             -FileName $schemaFileName
                                                             -masterLogFilePathAndName $masterLogFilePathAndName;



}