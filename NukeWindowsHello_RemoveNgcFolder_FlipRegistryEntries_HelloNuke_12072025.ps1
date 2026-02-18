<# 
.SYNOPSIS
    Nukes Windows Hello (Ngc folder) and sets registry policies to prevent
    Windows Hello for Business prompts / provisioning in the future.

.NOTES
    Run as Administrator.
#>

# region Pre-flight checks
Write-Host "=== Windows Hello Nuke ++ starting ==="

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

# region 1: Nuke Ngc container (existing 'Hello nuke' behavior)

$NgcPath = Join-Path $env:windir "ServiceProfiles\LocalService\AppData\Local\Microsoft\Ngc"

if (Test-Path $NgcPath) {
    Write-Host "Taking ownership of Ngc folder: $NgcPath"

    # Take ownership and grant Administrators full control
    & takeown.exe /f "$NgcPath" /r /d y | Out-Null
    & icacls.exe "$NgcPath" /grant "Administrators:F" /t /c | Out-Null

    Write-Host "Deleting Ngc folder..."
    try {
        Remove-Item -Path "$NgcPath" -Recurse -Force -ErrorAction Stop
        Write-Host "Ngc folder deleted successfully."
    }
    catch {
        Write-Warning "Failed to delete Ngc folder: $($_.Exception.Message)"
    }
}
else {
    Write-Host "Ngc folder not found (already gone): $NgcPath"
}

# OPTIONAL: If you want to also disable the Windows Hello services, 
# uncomment this block. Usually not required if we are just trying to
# stop provisioning prompts

$helloServices = @(
    "NgcSvc",      # Microsoft Passport
    "NgcCtnrSvc"   # Microsoft Passport Container
)

foreach ($svc in $helloServices) {
    try {
        if (Get-Service -Name $svc -ErrorAction SilentlyContinue) {
            Write-Host "Stopping and disabling service: $svc"
            Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
            Set-Service -Name $svc -StartupType Disabled
        }
    }
    catch {
        Write-Warning "Failed to adjust service $svc: $($_.Exception.Message)"
    }
}


# region 2: Registry – Hard-disable Windows Hello for Business & prompts

Write-Host "Configuring registry to disable Windows Hello for Business and post-logon provisioning..."

# 2.1 Disable Windows Hello for Business (policy equivalent to 'Use Windows Hello for Business = Disabled')
$pfwKey = "HKLM:\SOFTWARE\Policies\Microsoft\PassportForWork"
New-Item -Path $pfwKey -Force | Out-Null

# Enabled = 0  → WHfB disabled
New-ItemProperty -Path $pfwKey -Name "Enabled" -Value 0 -PropertyType DWord -Force | Out-Null

# DisablePostLogonProvisioning = 1 → stops the "set up a PIN / Windows Hello" nag after logon
New-ItemProperty -Path $pfwKey -Name "DisablePostLogonProvisioning" -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host "  - PassportForWork: Enabled=0, DisablePostLogonProvisioning=1 set."