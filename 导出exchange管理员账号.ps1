$server_list = @("BJMAIL1","BJMAIL2","BJMAIL3","BJMAIL4","BJMAIL5","BJMAIL6","BJMAIL7","BJMAIL8","BJMAIL9","BJMAIL10","BJMAIL11","BJMAIL12")
# 定义文件导出目录
$FilePath = "D:\test"
Get-ADGroupMember -Identity "Organization Management" | Select-Object name,samaccountname | Export-Csv -Encoding utf8 "$FilePath\OrganizationManagement.csv" -NoTypeInformation
Get-ADGroupMember -Identity "Exchange Trusted Subsystem" | Select-Object name,samaccountname | Export-Csv -Encoding utf8 "$FilePath\ExchangeTrustedSubsystem.csv" -NoTypeInformation
Get-ADGroupMember -Identity "Domain Admins" | Select-Object name,samaccountname | Export-Csv -Encoding utf8 "$FilePath\DomainAdmins.csv" -NoTypeInformation
function get-user {
param ($computer)
$groups = Get-WmiObject -Class Win32_GroupUser -ComputerName $computer
$admins = $groups | where-object {$_.groupcomponent -like '*"Administrators"'} 
#获取groupcomponent属性与Administrators字符相关的行数
$admins | ForEach-Object{
$_.partcomponent -match ".+Domain\=(.+)\,Name\=(.+)$" > $nul
#这里使用正则表达式，获取并拿出匹配【.Domain="test",Name="Administrators"】的字符
$_.PSComputerName + "\"+$matches[1].trim('"') + "\" + $matches[2].trim('"')
#去掉分号，并把导出的字符以自己想要呈现的格式组合
}
}
foreach ( $hostname in $server_list){
if (!(test-connection $hostname -Count 1 -Quiet)){
#将无法ping通的服务器记录在一个txt文件中
Write-Output $hostname | Out-File "$FilePath\administrator-error.txt" -Append}
else{
#导出Administrators组内的成员
Get-User -computer $hostname | Out-File "$FilePath\Administrator-users.txt" -Append}
}
