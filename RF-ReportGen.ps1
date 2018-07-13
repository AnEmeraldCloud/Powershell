#------------------------------------------------------------------------------------------
#Purge any previous report files before starting the new iteration.
Remove-Item *.xls
#7/11-Need to revise this as its a plain sterilization of the directory. Perhaps by date.
#7/12-Setup date purging
$PurgeDelay = (Get-Date).AddDays(-7).Date
$Removal = Get-ChildItem "C:\Datarepo" -Recurse -Include "*.xls" | Where-Object {$_.CreationTime -lt $PurgeDelay}

$PurgeDelay = (Get-Date).AddDays(-7)
$Removal = Get-ChildItem "C:\Datarepo" -Recurse -Include "*.xls"
Where-Object {$_.CreationTime -lt $PurgeDelay}

$PurgeDelay = (Get-Date).AddDays(-1)
Get-ChildItem -Path "C:\Datarepo" -Name Automation -Recurse -Force | Where-Object { !$_.PSIContainer -and $_.CreationTime -lt $Purgedelay } | Remove-Item -Force -Whatif

$Removal = Get-ChildItem "C:\Datarepo" -Recurse -Include "*.xls" |
Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-1)}

# Working but need to refine
#$PurgeDelay = (Get-Date).AddDays(-7)
#$Removal = Get-ChildItem "C:\Datarepo" -Recurse -Include "*.xls"
# Where-Object {$_.LastWriteTime -lt $PurgeDelay}

#PS C:\DataRepo> $PurgeDelay = (Get-Date).AddDays(-1)
#Get-ChildItem "C:\Datarepo" -Recurse -Include "*.xls" | Where {$_.CreationTime -gt $PurgeDelay -and $_.Name -Match "Automation"}  | Remove-Item -Whatif

#------------------------------------------------------------------------------------------
#Setting variables for the EFR Command Line
$DateVar = Get-Date -UFormat %m/%d/%Y
$StartTime = "00:00:00 AM"
$EndTime = "11:59:59 PM"
$StartVar = $DateVar.tostring() + " " + $Starttime.tostring()
$EndVar = $Datevar.tostring() + " " + $EndTime.tostring()
#Generate new report from RightFax Server
& "C:\Program Files (x86)\RightFax\Adminutils\EnterpriseFaxReporter1.exe" -reportName "C:\Program Files (x86)\RightFax\AdminUtils\Reports\Automation_Outbound_Failed_User.rpt" -sqlServer "Quartz\SQLEXPRESS" -sqlDatabase "RightFax2" -sqlNTAuth "True" -dateStart $StartVar -dateEnd $Datevar + " " + $EndTime -paramSearchUser "HIMTEST" -outputPath " C:\DataRepo" -outputType "XLSR" -log "Verbose"
#6/22-Waits for report to finish within 30 sec before moving forward.
#Wait-Process -Name EnterpriseFaxReporter -Timeout 30
#7/12-Changed this to the Start-Sleep command instead of the Wait-Process due to issues in the process spawning. 
Start-Sleep -s 20
#------------------------------------------------------------------------------------------
#Sending Report via email
$From = "Sender@Quartz"
$To = "Tester@Quartz"
$Attachment = get-childitem -Name -Filter *.xls
Send-MailMessage -From $From -To $To -Subject "Daily RightFax Report for failed faxes." -Body "Attached is the RightFax report in XLS" -Attachments $Attachment -dno onFailure -SmtpServer Quartz
#------------------------------------------------------------------------------------------