#
# TaskRenameAndMoveTenantCHUserExportCsvFile.ps1
#
<# Header Information **********************************************************
Name: TaskRenameAndMoveChapterHubCSVFile.ps1
Created By: Mike John
Created Date: 11/14/2020
Summary:
    Task stub to run RenameAndMoveCHUserExportCsvFile.ps1
Update History *****************************************************************
Updated By: Mike John
Updated Date: 07/06/2021
    Reason Updated: Added $justTesting parameter"
Updated By: Mike John
Updated Date: 12/27/2020
    Reason Updated: rename script; Rename script being called
Updated By: Mike John
Updated Date: 11/14/2020
    Reason Updated: original version
#>
[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][ValidateSet("ALGA","ALSA","ALLV")][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=1)][bool]$justTesting
      )
begin {}
process
{

    $tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;

    #Get Credentials to connect
    [system.Management.Automation.PSCredential]$myCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

    [string]$tenantDomainName = $tenantObj.DomainName;

    # create the log file name
    $dateRightNow = Get-Date;
    [string]$myMasterLogFilePathAndName = 'c:\logs\' + $tenantAbbreviation + '\RenameAndMoveTenantChapterHubCSVFile' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

    $logMessage = "Starting RenameAndMoveTenantChapterHubCSVFile.ps1";
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;

    [string]$returnFileName = "File Not Found";

    $returnFileName = .\Utilities\RenameAndMoveTenantCHUserExportCsvFile.ps1 -tenantCredentials $myCredentials `
                                                                             -tenantAbbreviation $tenantAbbreviation `
                                                                             -tenantDomain $tenantDomainName `
                                                                             -masterLogFilePathAndName $myMasterLogFilePathAndName;
    #
    # There can be 3 return conditions: FileName, Error, and FileNotFound
    #
    if($returnFileName -eq "File Not Found" -or $returnFileName -eq "Error")
    {
        if($returnFileName -eq "Error")
        {
            $logMessage = "Sending Email to CHWorkFlow monitors. returnFileName: " + $returnFileName;
            .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
            # email to mass mailing coordinator that PDF file is available and it first of month
            [string]$myEmailSender           = "ITSupport@" + $tenantDomainName + ".org";
            [string[]]$myEmailRecipientArray = @($myEmailSender, "MJohn@algeorgetownarea.org")
            [string]$myEmailSubject          = "File Rename Issue - $(Get-Date -Format g)";
            [string]$myEmailBody             = "ReturnFileName: " + $returnFileName + ".  Please check logs.";
            .\EMailer\SendAnEmail.ps1 -tenantCredentials $myCredentials `
                                      -emailSender $myEmailSender `
                                      -emailRecipientArray $myEmailRecipientArray `
                                      -emailSubject $myEmailSubject `
                                      -emailBody $myEmailBody `
                                      -masterLogFilePathAndName $myMasterLogFilePathAndName;
            # write failure to status
        }
        else
        {
           if ($tenantAbbreviation -eq "ALGA") { #ALGA expects to get it every night
               $logMessage = "Sending Email to ALGA CHWorkFlow monitors. returnFileName: " + $returnFileName;
               .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
               if(!$justTesting)
               {
                    # email to mass mailing coordinator that PDF file is available and it first of month
                    [string]$myEmailSender = "ITSupport@" + $tenantDomainName + ".org";
                    [string[]]$myEmailRecipientArray = @("ITSupport@algeorgetownarea.org")
                    [string]$myEmailSubject = "File Rename Issue - $(Get-Date -Format g)";
                    [string]$myEmailBody = "ReturnFileName: " + $returnFileName + ".  Please check logs.";
                    .\EMailer\SendAnEmail.ps1 -tenantCredentials $myCredentials `
                                              -emailSender $myEmailSender `
                                              -emailRecipientArray $myEmailRecipientArray `
                                              -emailSubject $myEmailSubject `
                                              -emailBody $myEmailBody `
                                              -masterLogFilePathAndName $myMasterLogFilePathAndName;
                }
            }
        }
    }
    else
    {
        # write to status
    }
    $logMessage = "Finished TaskRenameAndMoveChapterHubCSVFile.ps1";
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
}
