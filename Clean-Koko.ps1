
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

$upn = "koko.breadmore@quantinuum.com"

# 1) Force off Focused and legacy Clutter at the mailbox
Set-FocusedInbox -Identity $upn -FocusedInboxOn:$false
Set-Clutter       -Identity $upn -Enable:$false

# 2) Disable any enabled rule that files to a folder containing 'Promotions'
Get-InboxRule -Mailbox $upn |
  Where-Object { $_.Enabled -eq $true -and $_.MoveToFolder -like "*Promotions*" } |
  ForEach-Object { Disable-InboxRule -Mailbox $upn -Identity $_.Identity -Confirm:$false }

# 3) Remove any rule (enabled or disabled) that references 'Promotions' to prevent a future re‑enable
Get-InboxRule -Mailbox $upn |
  Where-Object { $_.MoveToFolder -like "*Promotions*" } |
  ForEach-Object { Remove-InboxRule -Mailbox $upn -Identity $_.Identity -Confirm:$false }

# 4) Show the final state for the record
Get-InboxRule -Mailbox $upn |
  Select-Object Name, Enabled, Priority, MoveToFolder |
  Format-Table -Auto
``

