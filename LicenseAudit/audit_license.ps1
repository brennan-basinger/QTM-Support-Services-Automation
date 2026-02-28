# License Audit

$E3SkuId = "05e9a617-0261-4cee-bb44-138d3ef5d965"                 # E3
$DefenderBundleSkuId = "3dd6cf57-d688-4eed-ba52-9e40b5468c3e"      # Defender

$allUsers = Get-MgUser -All -Property "id,displayName,mail,userPrincipalName,assignedLicenses"

# Users with E3 but WITHOUT the Defender bundle
$E3_No_Defender = $allUsers | Where-Object {
    ($_.assignedLicenses.skuId -contains [guid]$E3SkuId) -and
    -not ($_.assignedLicenses.skuId -contains [guid]$DefenderBundleSkuId)
}

# Users with the Defender bundle but WITHOUT E3
$Defender_No_E3 = $allUsers | Where-Object {
    ($_.assignedLicenses.skuId -contains [guid]$DefenderBundleSkuId) -and
    -not ($_.assignedLicenses.skuId -contains [guid]$E3SkuId)
}

# Only select UserPrincipalName before exporting
$E3_No_Defender |
    Select-Object UserPrincipalName |
    Export-Csv -Path ".\E3_without_DefenderP2.csv" -NoTypeInformation

$Defender_No_E3 |
    Select-Object UserPrincipalName |
    Export-Csv -Path ".\DefenderP2_without_E3.csv" -NoTypeInformation