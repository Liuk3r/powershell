$csvPath = "\\10.2.166.163\Temp\UltraEdit"
$csvList = Get-ChildItem -Path $csvPath
$AppList = @()
foreach ($csv in $csvList) {
  $tmp = Import-Csv $csv.VersionInfo.FileName -Encoding UTF8
  $AppList = $AppList + $tmp
}
$AppList | Select-Object * | Export-Csv C:\Users\kallengao\Desktop\AppList.csv -Encoding UTF8 -NoTypeInformation
