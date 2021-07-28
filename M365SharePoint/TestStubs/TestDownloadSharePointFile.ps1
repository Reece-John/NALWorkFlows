#
# TestDownloadSharePointFile.ps1
#

$startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
cd $startLoc;

#Configuration Variables
$DownloadFileName = "ALChapterSchema.xlsx";
$SiteURL = "https://ALGeorgetownArea.sharepoint.com/sites/TechnologyTeam";
#$FileRelativeURL = "/sites/TechnologyTeam/Shared Documents/M365Management/M365DataFiles/" + $DownloadFileName;
$FileRelativeURL = "/sites/TechnologyTeam/Shared Documents/M365Management/M365DataFiles/" + $DownloadFileName;
$DownloadPath ="C:\PSScripts\ALGA\ExcelDataFiles";

# create the log file name
$dateRightNow = Get-Date;
[string]$masterLogFilePathAndName = 'c:\logs\TestDownLoadSharePointFile_' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

#Get Credentials to connect
$myCredentials = .\Common\ReturnCredentials.ps1

.\M365SharePoint\DownLoadSharePointFile.ps1 -tenantCredentials $myCredentials `
                                            -SharePointSiteURL $SiteURL `
                                            -SharePointFileRelativeURL $FileRelativeURL `
                                            -LocalFileDownloadPath $DownloadPath `
                                            -FileName $DownloadFileName `
                                            -masterLogFilePathAndName $masterLogFilePathAndName;

Write-Output("Wrap up TestDownloadSharePointFile.ps1");