Import-Module ActiveDirectory
Write-Host “”
Write-Host “* This script will dump out users in named groups, all ”
Write-host “ groups or a range of groups. You will be guided ”
Write-host “ through the process ”
Write-Host “ ”
Write-Host “ All output will be saved to C:SupportScriptOutput *”
Write-Host “”
write-host
$strFileName = $(
$selection = read-host ‘Enter the name of file to save results to. Include an extention. (Default = Groupmembership.csv)’
if ($selection) {$selection} else {‘GroupMembership.csv’}
)
$strFileName = “C:SupportScriptOutput” + $strFilename
If (Test-Path $strFileName){
Remove-Item $strFileName
}
Write-Host
Write-Host ‘Enter name of group you would like to export’
Write-Host ‘The script will look for matching groups’
Write-Host
Write-Host ‘Entering the first part of the group name will return all matching groups’
Write-Host ‘For example, Entering “LG-APP-” without the quotation marks will return all application groups’
Write-Host
Write-Host ‘Pressing return will list membership of ALL groups’
Write-host
Write-Host ‘***** WARNING *****’
Write-Host
Write-Host ‘Exporting all group memberships will take some time as it will’
Write-Host ‘include all built in groups and distribution lists – use with caution’
Write-Host
$strGroupNames = $(
$selection = Read-Host ‘Enter name of group you would like to export (no value will return all groups)’
if ($selection) {$selection + ‘’} else {‘’}
)
Write-Host
Write-Host ‘Exporting teams with names like ‘ $strGroupNames ‘ to ‘ $strFilename
$data= ‘Group,UserName,UserID’
Write-Output $data | out-file -FilePath $strFileName -Append
$groups = Get-ADGroup -filter {name -like $strGroupNames}
foreach ($Group in $Groups)
{
$usernames = Get-ADGroupMember $($group.name)
foreach ($user in $usernames)
{
$data = $Null
$data = $data + $group.name + “,” + $user.name + “,” + $user.saAMAccountName
Write-Output $data | out-file -FilePath $strFileName -Append
}
