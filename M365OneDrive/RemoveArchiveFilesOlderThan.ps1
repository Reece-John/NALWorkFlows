#
# RemoveArchiveFilesOlderThan.ps1
#
<#
# Called from ArchiveProcessMaster.ps1 script
# installed in "$PSStartup\Admin\# RemoveFileOlderThan.ps1"
# Deletes files older than the $olderThanDays parameter (Recursive)
#                
#           Author: Mike John
#     Date Created: 1/6/2018
#
# Last date edited: 11/21/2020
#   Last edited By: Mike John
# Last Edit Reason: Updated log file messages
#
# Last date edited: 1/6/2018
#   Last edited By: Mike John
# Last Edit Reason: Original
#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][System.Management.Automation.PSCredential]$tenantCredentials
      ,[Parameter(Mandatory=$True,Position=1)][string]$archiveSiteURL
      ,[Parameter(Mandatory=$True,Position=2)][string]$archiveRelativeURL
      ,[Parameter(Mandatory=$True,Position=3)][int]$daysOlderThan
      ,[Parameter(Mandatory=$True,Position=4)][string[]]$extensionArray
      ,[Parameter(Mandatory=$True,Position=5)][string]$masterLogFilePathAndName
)
begin {}
process {

    # starts here

    [string]$logMessage = "Starting RemoveArchiveFilesOlderThan.ps1";
    .\LogManagement\WriteToLogFile.ps1 -logFile $masterLogFilePathAndName -message $logMessage;

    #region Set Variables
    [string]$filterString = (Get-Date).AddDays((-1 * $daysOlderThan));
    #endregion

    $logMessage = "Deleting files older than: " + $filterString;
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;

    Try
    {
        #Connect to PNP Online
        Connect-PnPOnline -Url $archiveSiteURL -Credentials $tenantCredentials

        #Get All Files from the document library - In batches of 500
        $ListItems = Get-PnPListItem -List $archiveRelativeURL -PageSize 500 | Where-Object {$_["Created"] -lt $filterString}

        ForEach($listItem in $ListItems)
        {
            $fileType =  $listItem.FieldValues['File_x0020_Type'];
            foreach($fileExtension in $extensionArray)
            {
                if($fileType -eq $fileExtension)
                {
                    $fileDirRef =  $listItem.FieldValues['FileDirRef'];
                    if($fileDirRef -like "*Chapter Hub Exports")
                    {
                        $fileURL = $listItem.FieldValues['FileRef'];
                        $createdDate = $listItem.FieldValues['Created'];

                        Write-Host($fileURL);
                        Write-Host($fileType);
                        $logMessage = "Deleting file: " + $fileURL + " - CreationTime: " + $createdDate.ToString(" MM/dd/yyyy HH:mm:ss");
                        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
                        #Remove-PnPFile -ServerRelativeUrl $fileURL -Force;
                    }
                }
            }
        }
        Disconnect-PnPOnline;
    }
    catch 
    {
        $logMessage = "Error: $($_.Exception.Message)";
        .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
    }

    $logMessage = "Finishing RemoveArchiveFilesOlderThan.ps1";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
}
end{}
