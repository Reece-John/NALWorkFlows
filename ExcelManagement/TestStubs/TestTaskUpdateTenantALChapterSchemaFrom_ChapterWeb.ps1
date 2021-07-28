#
# TestTaskUpdateTenantALChapterSchemaFrom_ChapterWeb.ps1
#

# Tenant testing
[string]$tenantAbbreviation = "ALLV";


# Set relative Location based off of production Machine or development machine
# if it is "Production" the location is set in the task scheduling definition
# To set Environment Variable: [Environment]::SetEnvironmentVariable("MachineUsage",<valueToSet>,"Machine");
[string]$machineUsage = [Environment]::GetEnvironmentVariable("MachineUsage","Machine");
if($null -eq $machineUsage -or $machineUsage -ne "Production")
{
    [string]$DevStartUpEnvName = $tenantAbbreviation + "DevStartup";
    $startUpObj = Get-ChildItem Env:$DevStartupEnvName;
    [string]$startLoc = $startUpObj.Value;
    Set-Location $startLoc;
}


.\ExcelManagement\TaskStubs\TaskUpdateTenantALChapterSchema.ps1 $tenantAbbreviation;

