<# Header Information **********************************************************
Name: TaskCreateTenantRoleBasedEmails.ps1
Created By: Mike John
Created Date: 02/20/2021
Summary:
    Task stub to run ManageM365TneantUsers.ps1
Update History *****************************************************************
Updated By: Mike John
Updated Date: 02/20/2021
    Reason Updated: original version
#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][ValidateSet("NAL")][string]$tenantAbbreviation
)
begin {}
process {
    [bool]$downLoadALChapterSchemaFile = $true;
    [bool]$justTestingonly = $false;

    # get parameters stored in the environment
    [string]$scope = "User";
    [string]$envDomainName = $tenantAbbreviation + "DomainName";
    [string]$tenantDomain = [Environment]::GetEnvironmentVariable($envDomainName, $scope);


    # create the log file name
    $dateRightNow = Get-Date;
    [string]$masterLogFilePathAndName = 'c:\logs\Create_' + $tenantAbbreviation + '_RoleBasedEmails_' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

    # get administrator credentials
    [System.Management.Automation.PSCredential]$myCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation;

    # Assign values to variables

    [string]$startUpLocationName = $tenantAbbreviation + "PSStartup";
    [string]$dataFileBaseLoc = [Environment]::GetEnvironmentVariable($startUpLocationName,"User");
    [string]$downloadPath = $dataFileBaseLoc + "\ExcelDataFiles"

    # SharePoint File copy variables
    $SiteURL = "https://" + $tenantDomain + ".sharepoint.com/sites/TechnologyTeam"
    $FileRelativeURL = "/sites/TechnologyTeam/Shared Documents/M365Management/M365DataFiles/NALSchema.xlsx"
    $DownloadPath = $dataFileBaseLoc + "\ExcelDataFiles"

    # Working Copy
    if($downLoadALChapterSchemaFile)
    {


        [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
        ,[Parameter(Mandatory=$True,Position=1)][string]$SharePointSiteURL
        ,[Parameter(Mandatory=$True,Position=2)][string]$SharePointFileRelativeURL
        ,[Parameter(Mandatory=$True,Position=3)][string]$LocalFileDownloadPath
        ,[Parameter(Mandatory=$True,Position=4)][string]$FileName
        ,[Parameter(Mandatory=$True,Position=5)][string]$masterLogFilePathAndName
   

    .\M365SharePoint\DownLoadSharePointFile.ps1 tenantCredentials -myCredentials `
                                                             -SharePointSiteURL $SiteURL `
                                                             -SharePointFileRelativeURL $FileRelativeURL `
                                                             -LocalFileDownloadPath $DownloadPath `
                                                             -FileName $schemaFileName
                                                             -masterLogFilePathAndName $masterLogFilePathAndName;
    }
    # Assign values to variables
    [string]$NALSchemaFilePathAndName = $dataFileBaseLoc + "\ExcelDataFiles\ALChapterSchema.xlsx";
    [string]$tenantDefaultsPageName = "M365TenantDefaults";
    [int]$tenantDefaultsStartRow = 1;
    [string]$rolebaseEmailsPageName = "RoleBasedEmailAddresses";
    [int]$rolebaseEmailsStartRow = 1;

    #set this to true to just write out what would be changed

    .\M365Management\CreateTenantRoleBasedEmails.ps1 -tenantCredentials $myCredentials `
                                                     -tenantAbbreviation $tenantAbbreviation `
                                                     -tenantDomain $tenantDomain `
                                                     -alChapterSchemaFilePathAndName $NALSchemaFilePathAndName `
                                                     -tenantDefaultsPageName $tenantDefaultsPageName `
                                                     -tenantDefaultsStartRow $tenantDefaultsStartRow `
                                                     -roleBasedEmailsPageName $rolebaseEmailsPageName `
                                                     -roleBasedEmailsStartRow $rolebaseEmailsStartRow `
                                                     -masterLogFilePathAndName $masterLogFilePathAndName `
                                                     -justTestingOnly $justTestingonly;

}