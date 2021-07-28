#
# TestALChapterSchemaUpload.ps1
#
Clear-Host;

# Set relative Location
#$startLoc = [Environment]::GetEnvironmentVariable("PSStartup","Machine");
$startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
Set-Location $startLoc;

$tenantAbbreviation= "NAL";

[PSCustomObject]$tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation
[string]$downloadPath = $tenantObj.PSStartUpDir + "\ExcelDataFiles"

# create the log file name
$dateRightNow = Get-Date;
[string]$myMasterLogFilePathAndName = 'c:\logs\TestALChapterSchemaUpload_' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

#Get Credentials to connect
$myCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

# Configuration Variables
[string]$upLoadSchemaFile = "ALChapterSchemaTest.xlsx"
[string]$upLoadSchemaSiteURL = "https://" + $tenantObj.DomainName + ".sharepoint.com/sites/TechnologyTeam";
[string]$upLoadschemaFileRelativeURL = "/Shared Documents/M365Management/M365DataFiles";
[string]$upLoadSourceFilePath = $downloadPath + "\" + $upLoadSchemaFile;

# Debug write statements
if($true)
{
    #region Debug write statements
    Write-Output($myMasterLogFilePathAndName);
    Write-Output($upLoadSchemaSiteURL);
    Write-Output($upLoadschemaFileRelativeURL);
    Write-Output($upLoadSourceFilePath);
    #endregion Debug write statements
}

if($true)
{
    .\M365SharePoint\UpLoadSharePointFile.ps1 -tenantCredentials $myCredentials `
                                              -SharePointSiteURL $upLoadSchemaSiteURL `
                                              -SharePointFileRelativeURL $upLoadschemaFileRelativeURL `
                                              -UploadPathAndFileName $upLoadSourceFilePath `
                                              -masterLogFilePathAndName $myMasterLogFilePathAndName;
}

Write-Host("Wrap Up TestALChapterSchemaUpload.ps1");
