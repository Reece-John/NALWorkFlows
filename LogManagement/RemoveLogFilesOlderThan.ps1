#
# RemoveAndLogFilesOlderThan.ps1
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
      [Parameter(Mandatory=$True,Position=0)][string]$directoryPath
     ,[Parameter(Mandatory=$True,Position=1)][int]$olderThanDays
     ,[Parameter(Mandatory=$True,Position=2)][string]$masterLogFilePathAndName
)
begin {}
process {
    [string]$logMessage = "Starting RemoveLogFilesOlderThan.ps1";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;

    $limit = (Get-Date).AddDays((-1 * $olderThanDays));
    $path = $directoryPath;

    # Log information on files older than the $limit.
    $files = Get-ChildItem -Path $path -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.LastWriteTime -lt $limit }
    foreach($fil in $files)
    {
        $logMessage = "Deleting file: " + $files.Name + " - CreationTime: " + $_.CreationTime.ToString("yyyyMMddTHHmmss") + " - LastWriteTime: " + $_.LastWriteTime.ToString("yyyyMMddTHHmmss");
        .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
    }
    # Delete files older than the $limit.
    Get-ChildItem -Path $path -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.LastWriteTime -lt $limit } | Remove-Item -Force
    $logMessage = "Finished RemoveLogFilesOlderThan.ps1";
    .\LogManagement\WriteToLogFile -logFile $masterLogFilePathAndName -message $logMessage;
}
end{}
