#
# ValidateDeveloperSetup.ps1
#
#
#           Author: Mike John
#     Date Created: 7/28/2020
#
# Last date edited: 7/28/2020
#   Last edited By: Mike John
# Last Edit Reason: Original
#

function Validate_a_Directory([string]$dirName)
{
    [bool]$directoryNameExists = Test-Path $dirName -ErrorAction SilentlyContinue
    if(!$directoryNameExists)
    {
        Write-Host("Directory: " + $dirName + " not found.")
    }
}

function Validate_EnvironmentVariable([string]$varName, [string]$varValue, [string]$scope)
{
    $tmpVarValue =[Environment]::GetEnvironmentVariable($varName, $scope);
    if($tmpVarValue -eq $null)
    {
        Write-Host("Environment Variable: " + $varName + ":" + $scope + " not set.")
    }
    else
    {
        if($tmpVarValue -ne $varValue)
        {
            Write-Host("Environment Variable: " + $varName + ":" + $scope + " not set to " + $varValue + ".")
        }
    }
}

# Starts Here
$chapterSiteName = "ALGA";
$DomainName = "ALGeorgetownArea";
$DomainExtension = ".org";
$uName = $null;
$userDevStartUpDirectory = $null;
$powerShellProfileDirectory = $null;

$userDevStartUpDirectory = "C:\Users\mjohn_r7awhu6\Source\Repos";
[bool]$destPathExists = Test-Path $userDevStartUpDirectory -ErrorAction SilentlyContinue;
if($destPathExists)
{
    # This must be the first part of your chapter email address
    # It will be part of the email that is used to get global permissions
    #  to login with 
    $uName = "MJohn";
    $powerShellProfileDirectory = "C:\Users\mjohn_r7awhu6\Documents\WindowsPowerShell\";
}
else
{
    $userDevStartUpDirectory = "C:\Users\sjs22\Source\Repos";
    $destPathExists = Test-Path $userDevStartUpDirectory -ErrorAction SilentlyContinue;
    if($destPathExists)
    {
        # This must be the first part of your chapter email address
        # It will be part of the email that is used to get global permissions
        #  to login with 
        $uName = "SSwain";
        $powerShellProfileDirectory = "C:\Users\sjs22\Documents\WindowsPowerShell\";
    }
    else
    {
        $userDevStartUpDirectory = "C:\Users\lprui\Source\Repos\";
        $destPathExists = Test-Path $userDevStartUpDirectory -ErrorAction SilentlyContinue;
        if($destPathExists)
        {
            # This must be the first part of your chapter email address
            # It will be part of the email that is used to get global permissions
            #  to login with 
            $uName = "LPruitt";
            $powerShellProfileDirectory = "C:\Users\lprui\Documents\WindowsPowerShell\";
        }
        else
        {
            $userDevStartUpDirectory = $null;
        }
    }
}
if($uName -eq $null -or $userDevStartUpDirectory -eq $null)
{
    Write-Host("Could not find the users Source\Repos directory")
    exit;
}

$tenantUser = $uName;

# create the POSH logging directory
Validate_a_Directory "c:\Logs";

[string]$baseScriptsDir = "c:\PSScripts\$chapterSiteName\";

# create the PSScripts directory structure
Validate_a_Directory $baseScriptsDir;
Validate_a_Directory ($baseScriptsDir + "aa_PhantomDir\");
Validate_a_Directory ($baseScriptsDir + "Admin\");
Validate_a_Directory ($baseScriptsDir + "Common\");
Validate_a_Directory ($baseScriptsDir + "Deploy\");
Validate_a_Directory ($baseScriptsDir + "EMailer\");
Validate_a_Directory ($baseScriptsDir + "ExcelDataFiles\");
Validate_a_Directory ($baseScriptsDir + "ExcelManagement\");
Validate_a_Directory ($baseScriptsDir + "ExcelRpts\");
Validate_a_Directory ($baseScriptsDir + "LogManagement\");
Validate_a_Directory ($baseScriptsDir + "M365Exchange\");
Validate_a_Directory ($baseScriptsDir + "M365Management\");
Validate_a_Directory ($baseScriptsDir + "M365SharePoint\");
Validate_a_Directory ($baseScriptsDir + "M365Teams\");
Validate_a_Directory ($baseScriptsDir + "Reports\");
Validate_a_Directory ($baseScriptsDir + "Utility\");

# create environment variables at machine level
Validate_EnvironmentVariable "PSStartUp" "C:\PSScripts\$chapterSiteName" "Machine";
# create environment variables at user level
Validate_EnvironmentVariable "PSStartUp" "C:\PSScripts\$chapterSiteName" "User";
Validate_EnvironmentVariable "CStorage" "C:\PSScripts\$chapterSiteName\Common" "User";
Validate_EnvironmentVariable "SiteName" "$chapterSiteName" "User";
Validate_EnvironmentVariable "TenantUser" $tenantUser "User";
Validate_EnvironmentVariable "DevStartup" ($userDevStartUpDirectory + "\$chapterSiteName\$chapterSiteName") "User";
Validate_EnvironmentVariable "DomainName" $DomainName "User";
Validate_EnvironmentVariable "DomainExtension" $DomainExtension "User";

$x = 1;


# now get Office 365 PowerShell modules if you do not already have them
if (-not(Get-Module -ListAvailable -Name PowershellGet))
{
  Write-Host "PowershellGet Module Not Loaded.";
}

if (-not(Get-Module -ListAvailable -Name MicrosoftTeams))
{
  Write-Host "MicrosoftTeams Module Not Loaded.";
}

if (-not(Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell"))
{
  Write-Host "Microsoft.Online.SharePoint.PowerShell Module Not Loaded.";
}

if (-not(Get-Module -ListAvailable -Name "MSOnline"))
{
  Write-Host "MSOnline Module Not Loaded.";
}

if (-not(Get-Module -ListAvailable -Name AzureAD))
{
  Write-Host "AzureAD Module Not Loaded.";
}

if (-not(Get-Module -ListAvailable -Name ExchangeOnlineManagement))
{
  Write-Host "ExchangeOnlineManagement Module Not Loaded.";
}

Write-Host("At end of ValidateDeveloperSetup.ps1");