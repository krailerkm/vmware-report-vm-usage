###############################################
#    Script 1.0a REPORT VSPP by Nostmax       #
###############################################
# Release note                                #
#                                             #
###############################################
#Script Report VM In ResourcePool and sent-email by krailerk_man
#========================= LoadPowerCLI =================================
#Add-PSSnapin VMware.VimAutomation.Core;
. "C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1";
#========================= BEGINPowerCLI =================================
$report = "";
$rootp = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent;
#Part of file export csv.
$vc1csv = "$rootp\vc1.csv";
#Remove file exportcsv.
if(Test-Path $vc1csv){
    try{
        Remove-Item $vc1csv -Force;
        $report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
        $report += " Remove Temp File :$vc1csv`r`n";
    }
    Catch{
        $report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
        $report += " Can not Remove Temp File`r`n";
    }
}
else{
    $report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
    $report += " Not have File :$vc1csv`r`n";
}
#Part of file export csv2.
$vc2csv= "$rootp\vc2.csv";
if(Test-Path $titancsv){
    try{
        Remove-Item $vc2csv-Force;
        $report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
        $report += " Remove Temp File :$titancsv`r`n";
    }
    Catch{
        $report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
        $report += " Can not Remove Temp File`r`n";
    }
}
else{
    $report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
    $report += " Not have File :$titancsv`r`n";
}
#======================== VC 1 Connect =================================
#Connect Server vCenter SKY CLOUD Cluster.
try{
    Connect-VIServer vc1.vsphere.local -user script@vsphere.local -password password;
    $report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
    $report += " Connect to VIServer vc1.vsphere.local`r`n";
}
Catch{
    $report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
    $report += " Con not Connect to VIServer vc1.vsphere.local`r`n";
}
#Part of file export csv.
#$vc1csv = "$rootp\vc1.csv";
#Export file csv.
$vc1csvf = Get-ResourcePool Infrastructure |get-vm | Select-Object Name,NumCPU,MemoryGB,UsedSpaceGB,ProvisionedSpaceGB,PowerState;
$vc1csvf | Export-Csv  $vc1csv -NoTypeInformation;
Disconnect-VIServer -Server vc1.vsphere.local -Force -Confirm:$false;
#======================== VC 2 Connect ===============================
#Connect Server vCenter TITAN Cluster.
try{
    Connect-VIServer vc2.vsphere.local -user script@vsphere.local -password password;
    $report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
    $report += " Connect to VIServer vc2.vsphere.local`r`n";
}
Catch{
    $report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
    $report += " Con not Connect to VIServer vc2.vsphere.local`r`n";
}
#Part of file export csv2.
#$vc2csv= "$rootp\vc2.csv";
#Export file csv2.
$vc2csvf = Get-ResourcePool Infrastructure |get-vm | Select-Object Name,NumCPU,MemoryGB,UsedSpaceGB,ProvisionedSpaceGB,PowerState;
$vc2csvf | Export-Csv  $vc2csv-NoTypeInformation;
Disconnect-VIServer -Server vc2.vsphere.local -Force -Confirm:$false;
#=========================== Send Email ==================================
##Sent e-mail.
##SMTP Server for relay e-mail.
$smtpServer = "smtprelay.meelab.th.com";
##E-mail detail.
$att = new-object Net.Mail.Attachment($vc1csv);
$att2 = new-object Net.Mail.Attachment($titancsv);
$msg = new-object Net.Mail.MailMessage;
$smtp = new-object Net.Mail.SmtpClient($smtpServer);
$msg.From = "admin@meelab.th.com";
$msg.To.Add("admin@meelab.th.com");
$msg.cc.Add("support@meelab.th.com");
$msg.cc.Add("system@meelab.th.com");
##Get mouth to day.
$date = Get-Date -UFormat "DAY:%d MONTH:%m YEAR:%Y";
$time = Get-Date;
$msg.Subject = "[TEST] :: REPORT INFRASTRUCTURE USE RESOURCE VSPP IN $date";
$msg.Body = "Dear ladies and gentleman, `r`n`r`nThis e-mail report infrastructure use resource vspp in $time from cloud infrastructure vc1 & vc2 as attachment file. By script build 1 `r`n`r`nBest Regards and Thank You.";
$msg.Attachments.Add($att);
$msg.Attachments.Add($att2);
$smtp.Send($msg);
$att.Dispose();
$att2.Dispose();
$report += Get-Date -Format "dd.MM.yyyy HH:mm:ss";
$report += " Send email Ready config`r`n";
#Log for active event
$report | out-file "$rootp\log_report_vspp.txt";
