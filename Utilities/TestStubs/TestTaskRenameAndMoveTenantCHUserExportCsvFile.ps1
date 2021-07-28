#
# TestTaskRenameAndMoveTenantCHUserExportCsvFile.ps1
#
# Set relative Location based off of production Machine or development machine
# if it is "Production" the location is set in the task scheduling definition

[string]$tenantAbbreviation = "ALGA";

$machineUsage = [Environment]::GetEnvironmentVariable("MachineUsage","Machine");
if($machineUsage -ne "Production")
{
    $startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
    Set-Location $startLoc;
}

.\Utilities\TaskStubs\TaskRenameAndMoveTenantCHUserExportCsvFile.ps1 $tenantAbbreviation;

