<#
.SYNOPSIS
    Generate WSUS Report of Needed Patches
.DESCRIPTION
    This scripts runs a report on needed patch count for each group in the group.txt file, then it cleans up the CSV, converts it to an HTML and launches another script to email the files to each group
#>

#Deletes previous reports
Remove-Item ".\reports\*"

#Loads module
Import-Module UpdateServices
#Make connection to WSUS Service
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer()

#Sets Computer Scope to include downstream machines
$computerscope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope 
$computerscope.IncludeDownstreamComputerTargets = "True"
$updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$updatescope.ApprovedStates = "LatestRevisionApproved"
$updatescope.UpdateApprovalActions = "Install"

#Loop for each group in text file
ForEach ($Targetgroup in Get-Content "groups.txt")
{
#Create Forward and Reverse lookup
$groups = @{}
$wsus.GetComputerTargetGroups() | ForEach {$groups[$_.Name]=$_.id;$groups[$_.ID]=$_.name}

$pcgroup = @($wsus.GetComputerTargets($computerscope) | Where {
$_.ComputerTargetGroupIds -eq $groups[$TargetGroup]
}) | Select -expand Id
$wsus.GetSummariesPerComputerTarget($updatescope,$computerscope) | Where {
$pcgroup -Contains $_.ComputerTargetID
} | ForEach {
$ComputerTarget = ($wsus.GetComputerTarget([guid]$_.ComputerTargetId))
New-Object PSObject -Property @{
ComputerTarget = $ComputerTarget.FullDomainName
NeededCount = ($_.DownloadedCount + $_.NotInstalledCount)
InstalledCount = $_.InstalledCount
FailedCount = $_.FailedCount
RebootNeeded = $_.InstalledPendingRebootCount
} | Select-Object ComputerTarget,NeededCount,InstalledCount,FailedCount,RebootNeeded| Sort NeededCount| Where-Object NeededCount -gt 0 | Export-Csv -Path ".\reports\$Targetgroup.csv" -NoTypeInformation -Append
}
}

#Renames reports to old files
Cd .\reports\
Get-ChildItem -Filter "*.csv" | Rename-Item -NewName {$_.name -replace '.csv','.csv.old' }

#Sorts reports files then exports to new csv 
$files = Get-ChildItem "" -Filter *.csv.old
ForEach ($file in $files) {Get-Content $file | ConvertFrom-Csv | Sort-Object {[int]$_.NeededCount} -Descending | Export-Csv -Path "$file.csv" -NoTypeInformation}

#Removes the duplicate csv in file names
dir *.csv.old.csv | rename-item -newname { $_.name -replace ".csv.old", "" }

#Deletes old files
Remove-Item *.old

Cd..

#Execute html conversion script
Invoke-Expression .\htmlconvert.ps1

Start-Sleep -s 10

#Execute email sub-script#
Invoke-Expression .\emailreports.ps1
