# Connect to Exchange Online
Import-Module ExchangeOnlineManagement

Connect-ExchangeOnline

# Confirm connection
try {
    Get-EXOMailbox -ResultSize 1 | Out-Null
    Write-Host "Connection to Exchange Online is working."
}
catch {
    Write-Host "Connection test failed. You may not be connected."
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false