#
# TestTaskCreateTenantRoleBasedEmails.ps1
#
# Default to the correct startup directory
$machineUsage = [Environment]::GetEnvironmentVariable("MachineUsage","Machine");
if($machineUsage -ne "Production")
{
    $startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
    cd $startLoc;
}

Disconnect-ExchangeOnline -ErrorAction SilentlyContinue;

[string]$tenantAbbreviation = "NAL";

.\M365Management\TaskStubs\TaskCreateTenantRoleBasedEmails.ps1 $tenantAbbreviation;


