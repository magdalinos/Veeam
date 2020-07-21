If (!(Get-PSSnapin -Name VeeamPSSnapIn -ErrorAction SilentlyContinue)) {
	If (!(Add-PSSnapin -PassThru VeeamPSSnapIn)) {
		Write-Error "Unable to load Veeam snapin" -ForegroundColor Red
		Exit
	}
}

$Result=@()

For ($i=0; $i -lt 7; $i++) {
    $StartDate = [DateTime]::Today.AddDays(-$i-1)
    $EndDate = [DateTime]::Today.AddDays(-$i)
	$Result += Get-VBRBackupSession | where {(($_.JobType -eq "Backup") -and ($_.CreationTime -ge $StartDate) -and ($_.CreationTime -le $EndDate))} | Sort JobName, CreationTime| Get-VBRTaskSession | Sort Name | Select-Object -Property @{N="Name";E={$_.Name}},@{N="Status";E={$_.Status}},@{N="JobName";E={$_.JobName}},@{N="Date";E={$StartDate}},@{N="Day";E={$i}}
}

$VMs = $Result | Sort Name | Group Name
$DateToday = [DateTime]::Today.ToString("ddd d MMM yyy")

$HTML="<!DOCTYPE html>
<html>
<head>
<title>Veeam Backup Report - $DateToday</title>
<style>
body {font-family: Tahoma; background-color:#ffffff;}
table {font-family: Tahoma;width: 80%;font-size: 12px;border-collapse:collapse;}
th {border: 1px solid #a7a9ac;border-bottom: none;}
td {border: 1px solid #a7a9ac;padding: 2px 3px 2px 3px;}
</style>
</head>
<body>
<h1><b><center>Veeam Backup Report - $DateToday</center></b></h1>
       <table><tbody>
       <tr>
                   <td style=`"text-align: center; color: #ffffff;`" bgcolor=`"green`">VM Name</td>
                   "

for ($i=-7;$i -le -1; $i++) {
    $HTML += "<td style=`"text-align: center;color: #ffffff;`" bgcolor=`"green`">"+[DateTime]::Today.AddDays($i).ToString("ddd d MMM yyy")+"</td>"
}
    $HTML += "</tr>"
    $EmailHTML += "</tr>"
foreach ($VM in $VMs) {
    $HTML += "<tr>
                   <td style=`"text-align: center;color: #ffffff;`" bgcolor=`"green`">"+$VM.Name.ToString()+"</td>
                   "
    $VMCell = $VM.Group|Sort Day|Group Day
    for ($i=0;$i -lt 7; $i++) {
        $bgcolor="orange"
        if ($VMCell.Name.Contains($i.ToString())) {
                $index = [array]::indexof($VMCell.Name,$i.ToString())
                if ($VMCell[$index].Group.Status.count -eq 1) {
                    $Statuses = @($VMCell[$index].Group.Status)
                } else {
                    $Statuses = $VMCell[$index].Group.Status
                }
                if ($Statuses.Contains("Success")) {
                   $OutStatus="Success"
                } elseif  ($Statuses.Contains("Warning")) {
                   $OutStatus="Warning"
                } else {
                   $OutStatus="Failed"
                } 
                Switch ($OutStatus)
                {
                 "Success"{$bgcolor="lightgreen"}
                 "Warning"{$bgcolor="yellow"}
                 "Failed"{$bgcolor="red"}
                }
                $HTML += "<td style=`"text-align: center;`" bgcolor=`""+$bgcolor+"`"><div class=`"tooltip`">"+$OutStatus+"</td>"
        } else {
               $HTML += "<td style=`"text-align: center;`" bgcolor=`"lightgrey`"></td>"
        }
    }
    $HTML+="</tr>"
}
$HTML+="</tbody></table></body></html>"

$emailTo = "receiver@mail.address.gr"       # Email address that you wish to send the report to
$smtpServerref = "10.10.10.10"               # SMTP server IP address used for emailing
$emailFrom = "sender@yourmail.gr"              # Email address that the report should come from
$Outputlocation = "C:\tmp\log.html"         # this should be the directory and file name to save a copy of the report
$HTML | Out-File $Outputlocation
#Email the report
$smtpServer = $smtpServerref
[string[]]$to = $emailTo
$from = $emailFrom
$subject = "Veeam Backup Report - ["+[DateTime]::Today.AddDays(-7).ToString("ddd d MMM yyy")+" - "+[DateTime]::Today.AddDays(-1).ToString("ddd d MMM yyy")+"]"
$body = ""
$body += $HTML

Send-MailMessage -SmtpServer $smtpServer -To $to -From $from -Subject $subject -Body $body -BodyAsHtml