<# 
Name:Emailreceivedstats.ps1
Version: 1
Download Link: https://github.com/Tunsworthy/
Author: Tom Unsworth (mail@tomunsworth.net)
Comment: 
  I was asked if i could show all the emails recevied by a mailbox in a week
Requriments:
    Powershell
    Exchange Read only admin access (could be just message tracking)

Setup in Task scheduleer 

-file "<Location>Emailreceivedstats.ps1" -exchange <Server> -Monitoredemail <emailaddresstolook at> -pastdays 7 -topshown 10 -SMTPServer <Server> -SMTPFrom NoReply@<Domain> -SMTPTo <Person> -internal(:$false)

Note: the exclution file is required at the moment if you don't have this the script won't work. (ill fix this in the next version)

Features:
-Excludes monitoring emails by looking at a csv
-Shows how many emails someone sent to the mailbox
-Emails results to admin
-Can be set only to look at internal emails 

#>


<#-------------------------------
Paramaters
--------------------------------#>

Param(
  [Parameter(Mandatory=$false,Position=1,helpmessage="EmailAddress")]
  [string]$Monitoredemail,
  [Parameter(Mandatory=$false,Position=1,helpmessage="Days to monitor")]
  [string]$pastdays,
  [Parameter(Mandatory=$false,Position=1,helpmessage="Users to Show(Top X)")]
  [string]$topshown,
  [Parameter(Mandatory=$true,Position=3,helpmessage="SMTP Server to use")]
  [string]$SMTPServer,
  [Parameter(Mandatory=$true,Position=4,helpmessage="Send From Email Address")]
  [string]$SMTPFrom,
  [Parameter(Mandatory=$true,Position=5,helpmessage="Send To Email Address(s)")]
  [string]$SMTPTo,
  [Parameter(Mandatory=$true,Position=5,helpmessage="Exchange Server")]
  [string]$Exchange,
  [Parameter(Mandatory=$false,Position=5,helpmessage="Only look at internal addresses")]
  [switch]$internal 
)

<#-------------------------------
Any arrays or custom settings here please
--------------------------------#>
#Exclution CSV Format 'EmailAddress'
$Exclutions = Import-Csv "Exclutions.csv"
$script:report_file = "report.html"
$email_log_file = "emaillog.csv"
$folder = get-location
<#-------------------------------
Function Rerpot Header
-------------------------------#>
Function Report-Header {

Remove-Item $report_file -ErrorAction SilentlyContinue

$Report_infomration = "
    <style>
    head { background-color:#FFFFFF;
           font-family:arial;
           font-size:12pt; }
    body { background-color:#FFFFFF;
           font-family:arial;
           font-size:12pt; }
    table{border-collapse: collapse;}
    TH {border-width: 1px;padding: 1px;border-style: solid;border-color:#4BACC6; background-color: #4BACC6; color: #FFFFFF;} 
    td {border-width: 1px;padding: 1px;}
    tr {border-width: 1px;padding: 1px;border-style: solid;border-color: black;}
    tr:nth-child(odd) { background-color:#FFFFFF; } 
    tr:nth-child(even) { background-color:#DAEEF3; } 
    </style>
    <body>
    <p><h2>Number of Emails to $($monitoredemail)</h2>
    <h4> In the last $($pastdays) Days</h4>
    </p>
    <p>This Report is Generated from $($env:COMPUTERNAME) using account $($env:USERDOMAIN)\$($env:USERNAME)  </br>
    Script Location: $($folder)\emailreceivedstats.ps1 </br>

    </body> 
    "

$Report = $Report_infomration
$Report | Out-File -Append $report_file

}


<#-------------------------------
Function Report Body
-------------------------------#>
Function Report-Body {
#Show Totals
$report_infomration ="<body><h2>Total Number of Emails Received</h2>"
$Report = $Report_infomration
$Report | Out-File -Append $report_file

$Report = $totalemail
$Report | Out-File -Append $report_file

#Show Top 10
$report_infomration ="<body><h2>Top $($topshown) Senders</h2></body>"
$Report = $Report_infomration
$Report | Out-File -Append $report_file 

$Report = $TopSenders | select count,name | sort count -Descending | Select-Object -First $($topshown) | ConvertTo-Html
$Report | Out-File -Append $report_file

}


<#-------------------------------
Function Email Report
-------------------------------#>
Function Report_Email {
     $date = get-date -UFormat "%Y-%m-%d"
     $body = Get-Content $script:report_file
     $messageSubject = "$pastdays day received report for $monitoredemail - $date"
     Send-MailMessage -From $smtpfrom -To $smtpto -Subject $messageSubject -BodyAsHtml ($body | Out-String) -Attachments $script:report_file -dno onSuccess, onFailure -SmtpServer $smtpServer
}

<#-------------------------------
Function to connect to Exchange
-------------------------------#>
Function Exchange-Connect {
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchange/PowerShell/
Import-PSSession $Session -AllowClobber

 }
<#-------------------------------
Function to Disconnect from Exchange
-------------------------------#>
Function Exchange-Disconnect {
Get-PSSession | where{$_.ConfigurationName -like "Microsoft.Exchange"} | Remove-PSSession
 }

 <#------------------------------
In this scetion Put the code you want to excute
-------------------------------#>

Function Process {

#clear out the old file
Remove-Item $email_log_file

#get the required dates
$today = get-date
$startdate = (get-date (get-date).AddDays(-$($pastdays)))

#get everything!
$allemails = Get-MessageTrackingLog -Recipients "$($monitoredemail)" -Start "$($startdate)" -End $today -EventId RECEIVE 
$allemails | Export-Csv $email_log_file -NoTypeInformation

$cleaned = $allemails | where {($Exclutions).EmailAddress -notcontains $_.sender}
<#-----------------------------

Now that we have all the emails we will do our filtering first if we only want to look at internal emails we will filter out both the exclution file and only show emails from the Accepted Domain list
Otherwise we will just filter out the exclution list
------------------------------#>
if($internal) {
    foreach ($domain in (Get-AcceptedDomain).DomainName){ 
    $dcleaned = $cleaned | where {$_.sender -like "*$domain"}
    $fcleaned += $dcleaned
   } 

    $script:totalemail = ($fcleaned).count
    #work out who is the biggest sender
    $script:TopSenders = ($fcleaned).sender | Group-Object 
}

#
else{
    #do clean up
    $script:totalemail = ($cleaned).count
    #work out who is the biggest sender
    $script:TopSenders = ($cleaned).sender | Group-Object 
}
}

<#-----------------------------
End of your section
------------------------------#>


Report-Header
#Connect to Exchange
Exchange-Connect
#Code to Process
Process
#build the report
Report-Body
Report_Email
#close Exchnge session
Exchange-Disconnect