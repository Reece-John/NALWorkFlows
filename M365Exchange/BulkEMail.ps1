#
# BulkEMail.ps1
#
# runs from the "$PSScripts\EMailer" directory
#                
#           Author: Mike John
#     Date Created: 11/11/2020
#
# Date edited: 11/11/2020
#   Edited By: Mike John
# Edit Reason: Original
#

<#
https://stackoverflow.com/questions/12460950/how-to-pass-credentials-to-the-send-mailmessage-command-for-sending-emails

https://gallery.technet.microsoft.com/scriptcenter/Send-HTML-Email-Powershell-6653235c

#$EmailTo = "myself@gmail.com"
$EmailFrom = "me@mydomain.com"
$Subject = "Test" 
$Body = "Test Body" 
$SMTPServer = "smtp.gmail.com" 
$filenameAndPath = "C:\CDF.pdf"
$SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$Body)
$attachment = New-Object System.Net.Mail.Attachment($filenameAndPath)
$SMTPMessage.Attachments.Add($attachment)
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587) 
$SMTPClient.EnableSsl = $true 
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential("username", "password"); 
$SMTPClient.Send($SMTPMessage)
#>


Function Send-EMail {
    Param ([Parameter(Mandatory=$true,  Position=0)][String]$EmailTo,
           [Parameter(Mandatory=$true,  Position=1)][String]$Subject,
           [Parameter(Mandatory=$true,  Position=2)][String]$Body,
           [Parameter(Mandatory=$true,  Position=3)][String]$EmailFrom="noreply@algeorgetownarea.org",  #This gives a default value to the $EmailFrom command
           [Parameter(mandatory=$true,  Position=4)][String]$Password,
           [Parameter(mandatory=$false, Position=5)][String]$attachment
    )

        $SMTPServer = "smtp.algeorgetownarea.org" 
        $SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$Body)
        if ($attachment -ne $null) {
            $SMTPattachment = New-Object System.Net.Mail.Attachment($attachment)
            $SMTPMessage.Attachments.Add($SMTPattachment)
        }
        $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587) 
        $SMTPClient.EnableSsl = $true 
        $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($EmailFrom.Split("@")[0], $Password); 
        $SMTPClient.Send($SMTPMessage)
        Remove-Variable -Name SMTPClient
        Remove-Variable -Name Password

} #End Function Send-EMail



function IsRunningAsAService($ServiceName)
{
    $arrService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($arrService -ne $null)
    {
        if ($arrService.Status -eq "Running")
        {
            return 1;
        }
    }
    return 0;
 }


#script starts here
$ServiceName = "BulkMailer";

$isNowAService = IsRunningAsAService $ServiceName
if(!$isNowAService)
{
    cls;
}

$startLoc = [Environment]::GetEnvironmentVariable("PSStartup","Machine");
cd $startLoc

############################################################################### 
 
###########Define Variables######## 
 
$fromaddress = "donotreply@ALAustin.net" 
$toaddress = "bj@quadpoint.com"
$bccaddress = "mjohn@quadpoint.com"
$CCaddress = "suevasser@gmail.com"
$Subject = "Test of bulk mailer"
$body = get-content .\content.htm
$attachment = "C:\sendemail\test.txt"
$smtpserver = "smtp.labtest.com"
 
#################################### 
 
$message = new-object System.Net.Mail.MailMessage 
$message.From = $fromaddress 
$message.To.Add($toaddress) 
$message.CC.Add($CCaddress) 
$message.Bcc.Add($bccaddress) 
$message.IsBodyHtml = $True 
$message.Subject = $Subject 
$attach = new-object Net.Mail.Attachment($attachment) 
$message.Attachments.Add($attach) 
$message.body = $body 
$smtp = new-object Net.Mail.SmtpClient($smtpserver) 
$smtp.Send($message) 