Write-Host "Resetting Windows Hello for Business (NGC folder cleanup)..." -ForegroundColor Cyan
 
$ngcPath = "C:\Windows\ServiceProfiles\LocalService\AppData\Local\Microsoft\Ngc"
 
if (Test-Path $ngcPath) {
    Write-Host "Taking ownership of NGC folder..." -ForegroundColor Yellow
    takeown /f $ngcPath /r /d y | Out-Null
 
    Write-Host "Granting Administrators full control..." -ForegroundColor Yellow
    icacls $ngcPath /grant administrators:F /t | Out-Null
 
    Write-Host "Deleting NGC folder..." -ForegroundColor Yellow
    Remove-Item $ngcPath -Recurse -Force
 
    Write-Host "NGC folder removed successfully." -ForegroundColor Green
} else {
    Write-Host "NGC folder does not exist. No cleanup required." -ForegroundColor Green
}
 
Write-Host "`nReboot is required to complete cleanup." -ForegroundColor Magenta