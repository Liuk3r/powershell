function Export-MailboxDatabaseInfo {
    param($Path='.\')
    $MailboxDatabaseInfos = @()
    $MailboxInfos = @()
    $DBs = Get-MailboxDatabase
    foreach ($DB in $DBs) {
        $MBCount = (Get-Mailbox -Database $DB | Measure-Object).Count
        $DBInfo = Get-MailboxDatabase $DB -status | Select-Object Name, ServerName, ActivationPreference, DatabaseSize, AvailableNewMailboxSpace
        
        $MailboxDatabaseInfo = New-Object PSObject
        Add-Member -InputObject $MailboxDatabaseInfo NoteProperty "DBName" $DBInfo.Name
        Add-Member -InputObject $MailboxDatabaseInfo NoteProperty "MBCount" $MBCount
        Add-Member -InputObject $MailboxDatabaseInfo NoteProperty "OnServer" $DBInfo.ServerName
        # Add-Member -InputObject $MailboxDatabaseInfo NoteProperty "主被动副本" $DBInfo.ActivationPreference
        Add-Member -InputObject $MailboxDatabaseInfo NoteProperty "DBSize" $DBInfo.DatabaseSize
        Add-Member -InputObject $MailboxDatabaseInfo NoteProperty "WhiteSize" $DBInfo.AvailableNewMailboxSpace
        if ($MBCount -eq 0) {
            Add-Member -InputObject $MailboxDatabaseInfo NoteProperty "AV.MBSize" 'N/A'
        }
        else {
            Add-Member -InputObject $MailboxDatabaseInfo NoteProperty "AV.MBSize" ($DBInfo.DatabaseSize / $MBCount)
        }        
        
        $MailboxDatabaseInfos += $MailboxDatabaseInfo
        $MailboxInfo = Get-Mailbox -Database $DB | Get-MailboxStatistics | Select-Object displayname,database,itemcount,TotalItemSize
        $MailboxInfos += $mailboxInfo
    }
    
    Import-Module ImportExcel -ErrorAction Stop
    $MailboxDatabaseInfos | Sort-Object DBName | Select-Object * | Export-Excel $Path\MailboxDatabaseInfo.xlsx -WorksheetName DBinfo -AutoSize -FreezeTopRow
    $MailboxInfos | Select-Object * | Export-Excel $Path\MailboxDatabaseInfo.xlsx -WorksheetName MBinfo -AutoSize -Append -FreezeTopRow
}
if (Get-PSSnapin "Microsoft.Exchange.Management.PowerShell.SnapIn") {
    Export-MailboxDatabaseInfo -Path C:\Users\kallengao\Desktop
    Remove-PSSnapin "Microsoft.Exchange.Management.PowerShell.SnapIn"
}
else {
    Add-PSSnapin "Microsoft.Exchange.Management.PowerShell.SnapIn"
    Export-MailboxDatabaseInfo -Path C:\Users\kallengao\Desktop
    Remove-PSSnapin "Microsoft.Exchange.Management.PowerShell.SnapIn"
}
