<# Header Information **********************************************************
Name: WriteToLogFile.ps1
Created By: Mike John
Created Date: 03/24/2021
Summary:
    Write a record to a logfile with a time stamp
Update History *****************************************************************
Updated By: Mike John
Updated Date: 03/24/2021
    Reason Updated: original version
#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)]$logFile
     ,[Parameter(Mandatory=$True,Position=1)]$message
    )
begin {}
process {
        $myTheDate = Get-Date;
        $fDate = $myTheDate.ToString("yyyyMMddTHHmmss");
        $lineToWrite = "[" + $fDate + "]" + $message;
        Add-Content $logFile $lineToWrite;
}
