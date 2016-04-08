## Retreives WSUS Groups and dumps into a text file ##

[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") 
 
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer() 
foreach( $group in $wsus.GetComputerTargetGroups() ) 
{ 
    $group.Name | out-file groups.txt -Append
}