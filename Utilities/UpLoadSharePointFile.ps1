#
# UpLoadSharePointFile.ps1
#
<# Header Information **********************************************************
Name: UpLoadSharePointFile.ps1
Created By: Mike John
Created Date: 09/20/2020
Summary:
    Uploads a local machine file to  SharePoint
Update History *****************************************************************
Updated By: Mike John
UpdatedDate: 09/20/2020
    Reason Updated: original version
#>
[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
     ,[Parameter(Mandatory=$True,Position=1)][string]$SharePointSiteURL
     ,[Parameter(Mandatory=$True,Position=2)][string]$SharePointFileRelativeURL
     ,[Parameter(Mandatory=$True,Position=3)][string]$UploadPathAndFileName
     ,[Parameter(Mandatory=$True,Position=4)][string]$masterLogFilePathAndName
)
begin {}
process {
#region function definitions
    function LogToMasterLogFile($logFile, $message)
    {
        $myTheDate = Get-Date;
        $fDate = $myTheDate.ToString("yyyyMMddTHHmmss");
        $lineToWrite = "[" + $fDate + "]" + $message;
        Add-Content $logFile $lineToWrite;
    }
#endregion

    [string]$logMessage = "Starting UpLoadSharePointFIle: " + $FileName;
    LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    [bool]$successfulUpload = $true;
    Try
    {
        #Connect to PNP On-line
        Connect-PnPOnline -Url $SharePointSiteURL -Credentials $tenantCredentials

        #Upload File to SharePoint On-line 
        $logMessage = "Uploading file:" + $UploadPathAndFileName;
        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $logMessage = "Uploading to: " + $SharePointSiteURL + " - " + $SharePointFileRelativeURL;
        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        Add-PnPFile -Path $UploadPathAndFileName -Folder $SharePointFileRelativeURL  | out-null
        Disconnect-PnPOnline;
    }
    catch 
    {
        write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
        $logMessage = "Error: $($_.Exception.Message)";
        LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
        $successfulUpload = $false;
    }
    if($successfulUpload)
    {
        $logMessage = "Finished UpLoadSharePointFIle: " + $FileName;
    }
    else
    {
        $logMessage = "Failed to UpLoadSharePointFIle: " + $FileName;
    }
    LogToMasterLogFile -logFile $masterLogFilePathAndName -message $logMessage;
}
