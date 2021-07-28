#CHWorkFlowMasterTask.ps1
<#
Task to run all steps (Tasks) in CHWorkFlow (Chapter Hub WorkFlow)
#>

[cmdletbinding()]
Param(
      [Parameter(Mandatory=$True,Position=0)][ValidateSet("NAL")][string]$tenantAbbreviation
     ,[Parameter(Mandatory=$True,Position=1)][bool]$justTesting
)
begin {}
process
{
    # create the log file name
    $dateRightNow = Get-Date;
    [string]$myMasterLogFilePathAndName = "c:\logs\"+ $tenantAbbreviation +"\CHWorkFlowMasterTask_" + $tenantAbbreviation + "_" + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';
    $logMessage = "Starting CHWorkFlowMasterTask for " + $tenantAbbreviation + " Testing: " + $justTesting.ToString();
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;

    .\Utilities\TaskStubs\TaskRenameAndMoveTenantCHUserExportCsvFile.ps1 -tenantAbbreviation $tenantAbbreviation -justTesting $justTesting;

    .\ExcelManagement\TaskStubs\TaskUpdateTenantALChapterSchemaFrom_ChapterHub.ps1 -tenantAbbreviation $tenantAbbreviation -justTesting $justTesting;

    .\ExcelRpts\TaskStubs\TaskTenantDailyRosterCreate.ps1 -tenantAbbreviation $tenantAbbreviation -justTesting $justTesting;

    <#
    .\M365Management\TaskStubs\TaskManageM365TenantUsers.ps1 -tenantAbbreviation $tenantAbbreviation -justTesting $justTesting;

    .\ExcelManagement\TaskStubs\TaskSchemaUsersMissingInCH.ps1 -tenantAbbreviation $tenantAbbreviation -justTesting $justTesting;

    .\M365Management\TaskStubs\TaskTenantUsersMissingInCH.ps1 -tenantAbbreviation $tenantAbbreviation -justTesting $justTesting;
    #>
    $logMessage = "Finished CHWorkFlowMasterTask for " + $tenantAbbreviation;
    .\LogManagement\WriteToLogFile -logFile $myMasterLogFilePathAndName -message $logMessage;
}
