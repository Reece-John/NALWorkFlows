# TestGetM365SubscriptionSKUs.ps1
#
# Default to the correct startup directory
$machineUsage = [Environment]::GetEnvironmentVariable("MachineUsage","Machine");
if($machineUsage -ne "Production")
{
    $startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
    cd $startLoc;
}

Disconnect-ExchangeOnline -ErrorAction SilentlyContinue;

[string]$tenantAbbreviation = "ALSA";

.\M365Management\TaskStubs\TaskGetM365SubscriptionSKUs.ps1 $tenantAbbreviation;



