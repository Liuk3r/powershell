If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
Start-Process powershell -Verb runAs -ArgumentList $arguments
Break
}
$RegLocal = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall","HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall","HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
$Softwares = $RegLocal | Get-ChildItem | Get-ItemProperty | Select-Object publisher,displayname,InstallLocation | Where-Object {($_.publisher -like "*IDM Computer Solutions*") -or ($_.displayname -like "*UltraEdit*")}
if($Softwares){
    $comname=(Get-WMIObject Win32_ComputerSystem).name
    $LogonInfo = (quser.exe) -replace '\s{2,}',',' | ConvertFrom-Csv
    if ($LogonInfo.USERNAME){
        $user = $LogonInfo.USERNAME.Substring(1,$LogonInfo.USERNAME.Length - 1)
    }
    else {
        $user = $LogonInfo.用户名.Substring(1,$LogonInfo.用户名.Length - 1)
    }
    #$sessionID = $LogonInfo.ID
    $fulluser = "sohu-inc\$user"
    
    $SoftwareLists = @()
    foreach($software in $Softwares){
        $List = New-Object -TypeName psobject
        Add-Member -InputObject $List NoteProperty "软件提供商" $Software.publisher
        Add-Member -InputObject $List NoteProperty "软件名" $Software.displayname
        Add-Member -InputObject $List NoteProperty "安装路径" $Software.InstallLocation
        Add-Member -InputObject $List NoteProperty "计算机名" $comname
        Add-Member -InputObject $List NoteProperty "最近登录用户" $fulluser
        $SoftwareLists = $SoftwareLists + $List
    }
    
    $SoftwareLists | Select-Object * | Export-Csv \\10.2.166.163\Temp\UltraEdit\$comname.csv -Encoding utf8 -NTI 
}
else {
    break
}
