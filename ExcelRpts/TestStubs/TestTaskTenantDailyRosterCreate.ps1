#
# TestTaskTenantDailyRosterCreate.ps1
#
# Set relative Location based off of production Machine or development machine
# if it is "Production" the location is set in the task scheduling definition
Clear-Host;
$machineUsage = [Environment]::GetEnvironmentVariable("MachineUsage","Machine");
if($machineUsage -ne "Production")
{
    $startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
    Set-Location $startLoc;
}

[string]$tenantAbbreviation = "NAL";

.\ExcelRpts\TaskStubs\TaskTenantDailyRosterCreate.ps1 $tenantAbbreviation;


