# TestConnect-MsolService
Clear-Host;

Import-Module MSOnline -UseWindowsPowerShell

[string]$tenantAbbreviation = "NAL";

# get administrator credentials
[System.Management.Automation.PSCredential]$tenantCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation;


    #connect to SharePoint On-line
    Connect-MsolService -Credential $tenantCredentials;

    $tenantSubscriptions = Get-MsolAccountSku;

    foreach($tlObj in $tenantSubscriptions)
    {
        $remainingUnits = $tlObj.ActiveUnits - $tlObj.ConsumedUnits;
        $logMessage = "{0,-42} {1,11} {2,12} {3,13} {4,6}" -f $tlObj.AccountSkuId, $tlObj.ActiveUnits, $tlObj.WarningUnits, $tlObj.ConsumedUnits, $remainingUnits;
        Write-Host($logMessage);
    }
