####Sends monthly WSUS reports to the respective teams ####
###WSUS reports generated must use the exact names in the containing folder or it will not find the attachement####
Cd .\reports
$Server = "mail.server.com"

#Send to Group Owner - Sends email report to groupDL@domain.com
$bodyADA = Get-Content "Group.html" -raw
send-mailmessage -to groupDL@domain.com -from noc@reedsmith.com -CC noc@reedsmith.com -subject "WSUS Report for Group" -body $bodyADA -BodyAsHtml -Attachments "Group.html" -smtpServer $Server