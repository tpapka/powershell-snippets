
''--- POWERSHELL - CLEAR CLIPBOARD
function Clear-Clipboard {
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Clipboard]::Clear()
}
Set-Alias -Name cc -Value Clear-Clipboard



''--- POWERSHELL - UNZIP FILE
$myZipFile = "C:\Mark\PowerShell\PSZipTest.zip"
$targetFolder = "C:\Mark\PowerShell\"

Add-Type -assembly "system.io.compression.filesystem"
[io.compression.zipfile]::ExtractToDirectory($myZipFile, $targetFolder) 



''--- POWERSHELL - LIST LOCAL USERS
Get-WmiObject -Class Win32_UserAccount
#Get-WmiObject -Class Win32_UserAccount | where {$_.Name -eq "XYZone"}



''--- POWERSHELL - SEND AN EMAIL
Send-MailMessage -to "john.doe@mail.com" -Subject "Hello" -from "john.doe@mail.com" -smtpserver "autodiscover.mail.com"



''--- POWERSHELL - SET TLS12
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12



''--- POWERSHELL - GET THE LATEST EMAILS FROM INBOX
Function Get-OutlookInBox 
{ 
 Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
 $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]  
 $outlook = new-object -comobject outlook.application 
 $namespace = $outlook.GetNameSpace("MAPI") 
 $folder = $namespace.getDefaultFolder($olFolders::olFolderInBox) 
 $folder.items |  
 Select-Object -Property Subject, ReceivedTime, Importance, SenderName 
}

Get-OutlookInBox



''--- POWERSHELL - OPENING NEW EXCEL BOOK
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.Workbooks.Add()
