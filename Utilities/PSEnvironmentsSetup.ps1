<#
# PSEnvironmentsSetup.ps1
#
<# Header
#*********************************
#           Author: Mike John
#     Date Created: 01/11/2021
#*********************************
# Last date edited: 03/17/2021
#   Last edited By: Mike John
# Last Edit Reason: Moved to VS Code from VS
#*********************************
# Last date edited: 01/11/2021
#   Last edited By: Mike John
# Last Edit Reason: Original
#*********************************
#>

#region Functions
function Make_a_Directory([string]$dirName)
{
    [bool]$directoryNameExists = Test-Path $dirName -ErrorAction SilentlyContinue
    if(!$directoryNameExists)
    {
        New-Item $dirName -type directory -force
    }
}
#endregion

#starts here
clear-Host;

[string]$tenantAbbreviation = "NAL";

[string]$baseScriptsDir = "c:\PSScripts\$tenantAbbreviation\";

# create the POSH logging directory
Make_a_Directory "c:\Logs";

# create the PSScripts directory structure
Make_a_Directory $baseScriptsDir;
Make_a_Directory ($baseScriptsDir + "aa_PhantomDir\");
Make_a_Directory ($baseScriptsDir + "Admin\");
Make_a_Directory ($baseScriptsDir + "Common\");
Make_a_Directory ($baseScriptsDir + "Deploy\");
Make_a_Directory ($baseScriptsDir + "EMailer\");
Make_a_Directory ($baseScriptsDir + "ExcelDataFiles\");
Make_a_Directory ($baseScriptsDir + "ExcelManagement\");
Make_a_Directory ($baseScriptsDir + "ExcelRpts\");
Make_a_Directory ($baseScriptsDir + "LogManagement\");
Make_a_Directory ($baseScriptsDir + "M365Exchange\");
Make_a_Directory ($baseScriptsDir + "M365Management\");
Make_a_Directory ($baseScriptsDir + "M365SharePoint\");
Make_a_Directory ($baseScriptsDir + "M365Teams\");
Make_a_Directory ($baseScriptsDir + "Reports\");
Make_a_Directory ($baseScriptsDir + "Tasks\");
Make_a_Directory ($baseScriptsDir + "Utilities\");

# now get Office 365 PowerShell modules if you do not already have them
if (Get-Module -ListAvailable -Name PowershellGet)
{
  Write-Host "PowershellGet Module already exists"
}
else
{
    Write-Host("Installing PowershellGet Module");
    Install-Module PowershellGet -Force
}

if (Get-Module -ListAvailable -Name MicrosoftTeams)
{
  Write-Host "MicrosoftTeams Module already exists"
}
else
{
    Write-Host("Installing MicrosoftTeams Module");
    Install-Module MicrosoftTeams -Force
}

if (Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell")
{
  Write-Host "Microsoft.Online.SharePoint.PowerShell Module already exists"
}
else
{
    Write-Host("Installing Microsoft.Online.SharePoint.PowerShell Module");
    Install-Module Microsoft.Online.SharePoint.PowerShell -Force
}

if (Get-Module -ListAvailable -Name "MSOnline")
{
  Write-Host "MSOnline Module already exists"
}
else
{
    Write-Host("Installing MSOnline Module");
    Install-Module MSOnline -Force
}

if (Get-Module -ListAvailable -Name AzureAD)
{
  Write-Host "AzureAD Module already exists"
}
else
{
    Write-Host("Installing AzureAD Module");
    Install-Module AzureAD -Force
}

if (Get-Module -ListAvailable -Name ExchangeOnlineManagement)
{
  Write-Host "ExchangeOnlineManagement Module already exists"
}
else
{
    Write-Host("Installing ExchangeOnlineManagement Module");
    Install-Module ExchangeOnlineManagement -Force
}

if (Get-Module -ListAvailable -Name PnP.PowerShell)
{
  Write-Host "PnP.PowerShell Module already exists"
}
else
{
    Write-Host("Installing PnP.PowerShell Module");
    Install-Module PnP.PowerShell -Force
}

if (Get-Module -ListAvailable -Name ImportExcel)
{
  Write-Host "ImportExcel Module already exists"
}
else
{
    Write-Host("Installing ImportExcel Module");
    Install-Module ImportExcel -Force
}
