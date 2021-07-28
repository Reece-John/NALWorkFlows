#
# TestTaskManageM365TenantUsers.ps1
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

.\M365Management\TaskStubs\TaskManageM365TenantUsers.ps1 $tenantAbbreviation;

