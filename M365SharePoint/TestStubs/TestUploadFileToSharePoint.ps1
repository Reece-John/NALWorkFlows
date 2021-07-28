#
# TestUploadFileToSharePoint.ps1
#
Clear-Host;
$startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
Set-Location $startLoc;

# create the log file name
$dateRightNow = Get-Date;
[string]$myMasterLogFilePathAndName = 'c:\logs\TestUpLoadSharePointFile_' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

#Get Credentials to connect
$myCredentials = .\Common\ReturnCredentials.ps1;

# Configuration Variables
$mySiteURL = "https://algeorgetownarea.sharepoint.com/sites/NALChapterTeam";
$myDestinationPath = "/Shared Documents/General/ChapterRoster";
$mySourceFilePath ="C:\PSScripts\NAL\ExcelDataFiles\Chapter Roster.xlsx";

# Debug write statements
if($true)
{
    #region Debug write statements
    Write-Output(Get-Date);
    Write-Output($mySourceFilePath);
    Write-Output($mySiteURL);
    Write-Output($myDestinationPath);
    Write-Output($myMasterLogFilePathAndName);
    #endregion Debug write statements
}

.\M365SharePoint\UpLoadSharePointFile.ps1 -tenantCredentials $myCredentials `
                                          -SharePointSiteURL $mySiteURL `
                                          -SharePointFileRelativeURL $myDestinationPath `
                                          -UploadPathAndFileName $mySourceFilePath `
                                          -masterLogFilePathAndName $myMasterLogFilePathAndName;

Write-Host("Wrap Up TestUploadFileToSharePoint.ps1");
