# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Device.ReadWrite.All", "Directory.AccessAsUser.All"

# Define the target device name
$deviceName = "USDTH64JV94"

# Retrieve the device by display name
$device = Get-MgDevice -Filter "displayName eq '$deviceName'" | Select-Object -First 1

# Check if device was found
if ($device) {
    # Cast DeviceId to string explicitly
    $deviceId = [string]$device.Id

    # Prepare the update payload
    $updateBody = @{
        extensionAttributes = @{
            extensionAttribute15 = "ssg"
        }
    }

    # Attempt to update the device
    try {
        Update-MgDevice -DeviceId $deviceId -BodyParameter $updateBody
        Write-Host "✅ extensionAttribute15 updated successfully for '$deviceName'"
    } catch {
        Write-Host "❌ Failed to update device '$deviceName': $($_.Exception.Message)"
    }
} else {
    Write-Host "⚠️ Device '$deviceName' not found."
}