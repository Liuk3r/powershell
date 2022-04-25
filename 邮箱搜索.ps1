$result = $list | Search-Mailbox -SearchQuery {From:"user1@test.com" AND attachment:"readme.zip" } -DeleteContent -Force

$sum = $result | Measure-Object -Property resultitemscount -Sum

$sum

Get-Mailbox -ResultSize unlimited | Search-Mailbox -SearchQuery {Sent:"2020/06/01..2020/06/03" AND attachment:"readme.zip" } -DeleteContent -Force
