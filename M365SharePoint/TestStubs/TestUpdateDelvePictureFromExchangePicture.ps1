#
# TestUpdateDelvePictureFromExchangePicture.ps1
#

<#

The following is not true: *****************************
This script avoids data-drift between the user's photo in Exchange and 
 the user's photo in Delve

#>


$startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
cd $startLoc;
$cred = .\Common\ReturnCredentials.ps1;


$userEmailAddress = "mjohn@algeorgetownarea.org";
$userPicturePathAndFileName = "C:\1\Tomek1.jpg";

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.ofxxxfice365.com/powershell-liveid/?proxyMethod=RPS -Credential $myCredentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
$userPicture = ([Byte[]] $(Get-Content -Path $userPicturePathAndFileName -Encoding Byte -ReadCount 0))
#set-userphoto -identity $userEmailAddress -picturedata $userPicture
Remove-PSSession $Session
