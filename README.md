# WSUS-Reporter
Getgroups.ps1
This script gets all the WSUS Groups and creates groups.txt. Shouldn't need to be run again unless the groups change. You will have to clean up the text file to only have the groups you want to create a report. Creates the reports in a subfolder \reports

WSUS_report.ps1
Master report that does a report on each group in Group.txt for pending windows updates. Converts the report to a CSV file. This script also launches HTML_Convert and Email scripts. I have these two seperate so I can change them without modifing the master script.

HTML Convert.ps1
Converts the CSV files to XML format then "needed column" to colors and then converts to HTML

emailreports.ps1 (the server must have the ablity to email from)
Emails each report in \reports to each respective team.

Shout out to Boe Prox at the link below! I used some of his scripting as a resource.
http://learn-powershell.net/2012/01/16/a-week-of-hey-scripting-guy-with-wsus-and-powershell/
