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
	$Result.Count
}

$VMs = $Result | Sort Name | Group Name

$HTML="<!DOCTYPE html>
<html>
<head>
<style>
#backups {
  font-family: `"Trebuchet MS`", Arial, Helvetica, sans-serif;
  border-collapse: collapse;
  width: 100%;
}

#backups td, #backups th {
  border: 1px solid #ddd;
  padding: 8px;
}

#backups tr:nth-child(even){background-color: #f2f2f2;}

#backups tr:hover {background-color: #ddd;}

#backups th {
  padding-top: 12px;
  padding-bottom: 12px;
  text-align: left;
  background-color: #4CAF50;
  color: white;
}

/* Tooltip container */
.tooltip {
  position: relative;
  display: inline-block;
  border-bottom: 1px dotted black; /* If you want dots under the hoverable text */
}

/* Tooltip text */
.tooltip .tooltiptext {
  visibility: hidden;
  background-color: #555;
  color: #fff;
  text-align: center;
  padding: 5px 0;
  border-radius: 6px;

  /* Position the tooltip text */
  position: absolute;
  z-index: 1;
  bottom: 125%;
  left: 50%;
  margin-left: -60px;

  /* Fade in tooltip */
  opacity: 0;
  transition: opacity 0.3s;
}

/* Tooltip arrow */
.tooltip .tooltiptext::after {
  content: "";
  position: absolute;
  top: 100%;
  left: 50%;
  margin-left: -5px;
  border-width: 5px;
  border-style: solid;
  border-color: #555 transparent transparent transparent;
}

/* Show the tooltip text when you mouse over the tooltip container */
.tooltip:hover .tooltiptext {
  visibility: visible;
  opacity: 1;
}
</style>
</head>
<body>
<h1><b><center>Veeam Backup Report</center></b></h1>
       <table id=`"backups`" border=`"1`"><tbody>
       <tr>
                   <td style=`"text-align: center;`">VM Name</td>
                   "
$EmailHTML=$HTML
for ($i=-7;$i -le -1; $i++) {
    $HTML += "<td style=`"text-align: center;`">"+[DateTime]::Today.AddDays($i).ToString("ddd d MMM yyy")+"</td>"
    $EmailHTML += "<td style=`"text-align: center;`" bgcolor=`"green`">"+[DateTime]::Today.AddDays($i).ToString("ddd d MMM yyy")+"</td>"
}
    $HTML += "</tr>"
    $EmailHTML += "</tr>"
foreach ($VM in $VMs) {
    $HTML += "<tr>
                   <td style=`"text-align: center;`">"+$VM.Name.ToString()+"</td>
                   "
    $EmailHTML += "<tr>
                   <td style=`"text-align: center;`" bgcolor=`"green`">"+$VM.Name.ToString()+"</td>
                   "
    $VMCell = $VM.Group|Sort Day|Group Day
    for ($i=0;$i -lt 7; $i++) {
        $bgcolor="orange"
        if ($VMCell.Name.Contains($i.ToString())) {
                $index = [array]::indexof($VMCell.Name,$i.ToString())
                if ($VMCell[$index].count -eq 1) {
                    Switch ($VMCell[$index].Group.Status.ToString())
                    {
                     "Success"{$bgcolor="lightgreen"}
                     "Warning"{$bgcolor="yellow"}
                     "Failed"{$bgcolor="red"}
                    }
                    $HTML += "<td style=`"text-align: center;`" bgcolor=`""+$bgcolor+"`"><div class=`"tooltip`">"+$VMCell[$index].Group.Status.ToString()+"<span class=`"tooltiptext`">"+$VMCell[$index].Group.JobName+"</span></div></td>"
                    $EmailHTML += "<td style=`"text-align: center;`" bgcolor=`""+$bgcolor+"`">"+$VMCell[$index].Group.Status.ToString()+"</td>"
                    } elseif ($VMCell[$index].count -gt 1) {
                        $HTML += "<td style=`"text-align: center;`"> <table style=`"width: 100%;`"><tbody>"
                        $EmailHTML += "<td style=`"text-align: center;`"> <table style=`"width: 100%;`"><tbody>"
                        foreach ($Backup in $VMCell[$index].Group) {
                            Switch ($Backup.Status)
                            {
                             "Success"{$bgcolor="lightgreen"}
                             "Warning"{$bgcolor="yellow"}
                             "Failed"{$bgcolor="red"}
                            }
                            $HTML += "<tr><td style=`"text-align: center;`" bgcolor=`""+$bgcolor+"`"><div class=`"tooltip`">"+$Backup.Status+"<span class=`"tooltiptext`">"+$Backup.JobName+"</span></div></td></tr>"
                            $EmailHTML += "<tr><td style=`"text-align: center;`" bgcolor=`""+$bgcolor+"`">"+$Backup.Status+"</td></tr>"
                        }
                        $HTML += "</tbody></table></td>"
                        $EmailHTML += "</tbody></table></td>"
                    }
        
        } else {
                 $HTML += "<td style=`"text-align: center;`"></td>"
                 $EmailHTML += "<td style=`"text-align: center;`"></td>"
        }
    }
    $HTML+="</tr>"
    $EmailHTML+="</tr>"
}
$HTML+="</tbody></table></body></html>"
$EmailHTML+="</tbody></table></body></html>"

$emailTo = "receiver@mail.address.com"               # Email address that you wish to send the report to
$smtpServerref = "10.10.10.10"              # SMTP server IP address used for emailing
$emailFrom = "sender@mail.address.com"              # Email address that the report should come from
$Outputlocation = "C:\tmp\log.html"         # this should be the directory and file name to save a copy of the report
$HTML | Out-File $Outputlocation
$DateToday = [DateTime]::Today.ToString("ddd d MMM yyy")
#Email the report
$smtpServer = $smtpServerref
[string[]]$to = $emailTo
$from = $emailFrom
$subject = "Veeam Backup Report - $DateToday"
$body = ""
$body += $EmailHTML

Send-MailMessage -SmtpServer $smtpServer -To $to -From $from -Subject $subject -Body $body -BodyAsHtml