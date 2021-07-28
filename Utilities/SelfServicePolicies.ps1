#SelfServicePolicies.ps1

<#
***************************************************************************************************************
Warning this can only be run in the ISE and not in VS Code.
***************************************************************************************************************

I put this in VS code just to get it checked into GitHub's version control

Link that I got most of this information on creating this script
https://www.enowsoftware.com/solutions-engine/blocking-self-service-purchases
#>

CLear-Host;
Connect-MSCommerce

#Get list of MSCommerceProductPolicies
Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase;
<#

ProductName                   ProductId    PolicyId                 PolicyValue
-----------                   ---------    --------                 -----------
Power Automate per user       CFQ7TTC0KP0N AllowSelfServicePurchase Disabled   
Power Apps per user           CFQ7TTC0KP0P AllowSelfServicePurchase Disabled   
Power Automate RPA            CFQ7TTC0KXG6 AllowSelfServicePurchase Enabled    
Power BI Premium (standalone) CFQ7TTC0KXG7 AllowSelfServicePurchase Enabled    
Visio Plan 2                  CFQ7TTC0KXN8 AllowSelfServicePurchase Enabled    
Visio Plan 1                  CFQ7TTC0KXN9 AllowSelfServicePurchase Enabled    
Project Plan 3                CFQ7TTC0KXNC AllowSelfServicePurchase Enabled    
Project Plan 1                CFQ7TTC0KXND AllowSelfServicePurchase Enabled    
Power BI Pro                  CFQ7TTC0L3PB AllowSelfServicePurchase Enabled  
#>

<#
[PSCustomObject]$policyObjs = new PSCustomObject;
}

[PSCustomObject]$policyObj = [PSCustomObject][ordered]@{

    ProductName       = "xxx"
    ProductId         = "XXX"
    PolicyId          = "AllowSelfServicePurchase"
    EnablePolicyValue = $false
}    
$policyObjs += $policyObj

#>

<#

Fillin the rest or make a loop




# Power Automate per user
[boolean]$EnablePowerAutomatePerUser = $False;
Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0KP0N -Enabled $EnablePowerAutomatePerUser;

# PowerApps per user
[boolean]$EnablePowerAppsPerUser = $False;
Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0KP0N -Enabled $EnablePowerAppsPerUser;


#>
# Allow Power BI Pro
#Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0L3PB -Enabled $True


#Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0KP0N -Enabled $True
#Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0KP0P -Enabled $True
#Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0KXG6 -Enabled $True
#Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0KXG7 -Enabled $True
#Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0KXN9 -Enabled $True
#Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0KXNC -Enabled $True
#Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0KXND -Enabled $True
#Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId CFQ7TTC0L3PB -Enabled $True
