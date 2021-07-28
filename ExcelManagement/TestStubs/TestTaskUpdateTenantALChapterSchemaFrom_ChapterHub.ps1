#
# TestTaskUpdateTenantALChapterSchema.ps1
#

# Tenant testing
Clear-Host;

[string]$tenantAbbreviation = "NAL";


# Set relative Location based off of production Machine or development machine
# if it is "Production" the location is set in the task scheduling definition
# To set Environment Variable: [Environment]::SetEnvironmentVariable("MachineUsage",<valueToSet>,"Machine");
$machineUsage = [Environment]::GetEnvironmentVariable("MachineUsage","Machine");
if($null -eq $machineUsage -or $machineUsage -ne "Production")
{
    $startLoc = [Environment]::GetEnvironmentVariable("NALDevStartup","User");
    Set-Location $startLoc;
}


.\ExcelManagement\TaskStubs\TaskUpdateTenantALChapterSchemaFrom_ChapterHub.ps1 $tenantAbbreviation;

