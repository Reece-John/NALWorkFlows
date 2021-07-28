#TestCHWorkFlowTask.ps1

Clear-Host;
$machineUsage = [Environment]::GetEnvironmentVariable("MachineUsage","Machine");
if($machineUsage -ne "Production")
{
    $startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
    Set-Location $startLoc;
}

# create chWorkFlowMasterConfig array to hold the different chapter Workflow configurations
$chWorkFlowMasterConfigObjs = @();
#NAL
$chWorkFlowMasterConfigObj = [PSCustomObject]@{
    tenantAbbreviation = 'NAL'
    testRunOnly        = $true
}
$chWorkFlowMasterConfigObjs += $chWorkFlowMasterConfigObj

$chWorkFlowMasterConfigObjs += $chWorkFlowMasterConfigObj

foreach($obj in $chWorkFlowMasterConfigObjs)
{
    Write-Host("Calling CHWorkFlowMasterTask.ps1 for " + $obj.tenantAbbreviation);
    
    if($obj.tenantAbbreviation -eq "NAL")
    {
        .\WorkFlowTask\CHWorkFlowMasterTask.ps1 -tenantAbbreviation $obj.tenantAbbreviation -justTesting $obj.testRunOnly;
    }
}
