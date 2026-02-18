<# ====================================================================================
Invoke-UserOffboarding.ps1
------------------------------------------------------------------------------------
Offboarding runbook automation for Microsoft 365/Entra ID using:
  • Exchange Online PowerShell (EXO)
  • Microsoft Graph PowerShell (Mg)

SAFE BY DEFAULT — the script runs in Preview unless you pass -Apply
It captures BEFORE / AFTER snapshots and writes paste-ready ServiceNow work notes

USAGE (Preview)
  .\Invoke-UserOffboarding.ps1 -UserUpn "first.last@quantinuum.com" -TicketNumber "INC00001234"

USAGE (Apply)
  .\Invoke-UserOffboarding.ps1 -UserUpn "first.last@quantinuum.com" -TicketNumber "INC00001234" -Apply

NOTES
  • Distribution lists (DGs) & mail-enabled security groups are removed via EXO.
  • Microsoft 365/Security groups are removed via Graph.
  • Dynamic groups are detected and **never** changed (we only list them).
  • Mailbox is converted to Shared by default and stamped with a future expiry 
    marker in CustomAttribute15 (e.g., "Expires: 2026-04-21 (180d)").
  • AD/on-prem steps are optional and skipped unless you request them AND the AD 
    module is available.
==================================================================================== #>

[CmdletBinding()]
param(
  # Core
  [Parameter(Mandatory=$true)][string]$UserUpn,
  [Parameter(Mandatory=$true)][string]$TicketNumber,

  # Mailbox handling
  [switch]$ConvertMailboxToShared = $true,
  [int]$SharedMailboxExpiryDays = 180,

  # Supervisor / manager options
  [string]$SupervisorUpn,
  [string]$BackupOwnerUpn,
  [switch]$GrantSupervisorFullAccess,
  [switch]$GrantSupervisorSendAs,

  # Group & license cleanup
  [switch]$RemoveFromDistributionLists = $true,
  [switch]$RemoveFromGroups          = $true,
  [switch]$RemoveMailboxDelegations  = $true,
  [switch]$RemoveLicenses = $true,
  [switch]$DisableEntraSignIn = $true,

  # Active Directory (on-prem) — optional. If not available, we skip.
  [switch]$DisableAD,
  [switch]$UpdateAdDescription,
  [string]$DisabledOuDn,  # e.g. "OU=Disabled,OU=Corp,DC=contoso,DC=com"

  # Execution control
  [switch]$Apply,                          # do changes only when present
  [string]$TenantHint = '94c4857e-1130-4ab8-8eac-069b40c9db20',  # tenant id or verified domain
  [switch]$UseElevatedGraphScopes,         # adds Directory.ReadWrite.All
  [string]$OutputFolder = (Join-Path $env:USERPROFILE ("Desktop\Offboarding-" + (Get-Date -Format "yyyyMMdd-HHmmss")))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------- helpers ----------
function Say([string]$msg){ Write-Host $msg }
function Step([string]$msg){ Write-Host ("== {0}" -f $msg) -ForegroundColor Cyan }
function Act ([string]$msg){ Write-Host $msg -ForegroundColor Yellow }
function Skip([string]$msg){ Write-Warning $msg }
function Did ([string]$msg){ Write-Host $msg -ForegroundColor Green }

# ---------- collection/CSV safety helpers ----------
function As-Array {
  param($x)
  if ($null -eq $x) { return @() }
  if ($x -is [System.Array]) { return $x }
  return @($x)
}

function Get-ItemCount {
  param([AllowNull()] $InputObject)
  try {
    if ($null -eq $InputObject) { return 0 }
    $m = ($InputObject | Measure-Object)
    if ($m -and ($m.PSObject.Properties.Match('Count').Count -gt 0)) { return [int]$m.Count }
    if ($InputObject -is [System.Array]) { return [int]$InputObject.Length }
    if ($InputObject -is [System.Collections.ICollection]) { return [int]$InputObject.Count }
    return 1
  } catch { return 0 }
}

function CountOf { param([AllowNull()]$x) return (Get-ItemCount -InputObject $x) }

function Has-Prop {
  param($obj,[string]$Name)
  try { return ($null -ne $obj -and $obj.PSObject.Properties.Match($Name).Count -gt 0) } catch { return $false }
}

function Write-CsvSafe {
  param(
    [AllowNull()] $Data,
    [Parameter(Mandatory)][string]$Path,
    $Headers,
    [switch]$Append
  )
  $count = CountOf $Data
  if ($count -gt 0) {
    if ($Append -and (Test-Path $Path)) {
      $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Append
    } else {
      $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    }
  } else {
    if ($Append -and (Test-Path $Path)) { return }
    $hCount = CountOf $Headers
    if ($hCount -gt 0) {
      ($Headers -join ',') | Out-File -FilePath $Path -Encoding utf8
    } else {
      '' | Out-File -FilePath $Path -Encoding utf8
    }
  }
}

$Preview = -not $Apply

# Create output folder & transcript
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$TranscriptPath = Join-Path $OutputFolder ("Transcript-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".txt")
Start-Transcript -Path $TranscriptPath -Append | Out-Null

# ---------- environment checks ----------
function Ensure-ModuleLoaded {
  param([string]$Name,[Version]$MinVersion)

  # Prefer what's already available to avoid "in use" update errors
  $have = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
  if ($have) {
    try { Import-Module $have -ErrorAction Stop; return } catch { }
  }

  # If not available (or import failed), try to install; if install fails, last-resort import without version gate
  try {
    Act ("Installing module {0} (min {1})" -f $Name, $MinVersion)
    try { Install-Module $Name -Force -Scope AllUsers -MinimumVersion $MinVersion -ErrorAction Stop }
    catch { Install-Module $Name -Force -Scope CurrentUser -MinimumVersion $MinVersion -ErrorAction Stop }
    Import-Module $Name -ErrorAction Stop
  } catch {
    Skip ("Could not install {0}. Using any locally available version if present. Error: {1}" -f $Name, $_)
    try { Import-Module $Name -ErrorAction Stop } catch { throw ("Module {0} not available and install failed." -f $Name) }
  }
}

# Pass -Organization only when the hint looks like a domain
function Ensure-EXO {
  Ensure-ModuleLoaded -Name ExchangeOnlineManagement -MinVersion ([Version]'3.3.0')

  if (-not (Get-ConnectionInformation)) {
    Act "Connecting to Exchange Online..."
    $isDomain = ($TenantHint -and ($TenantHint -match '^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'))


    # 1) Try normal WAM sign-in
    try {
      if ($isDomain) { Connect-ExchangeOnline -ShowBanner:$false -Organization $TenantHint -ErrorAction Stop | Out-Null }
      else           { Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null }
      return
    } catch { Skip ("WAM sign-in failed, retrying with Device Code: {0}" -f $_) }

<#

    # 2) Fallback to Device Code (bypasses WAM entirely)
    try {
      if ($isDomain) { Connect-ExchangeOnline -UserPrincipalName $UserUpn -Organization $TenantHint -Device -ShowBanner:$false -ErrorAction Stop | Out-Null }
      else           { Connect-ExchangeOnline -UserPrincipalName $UserUpn -Device -ShowBanner:$false -ErrorAction Stop | Out-Null }
      return
    } catch { Skip ("Device Code also failed, trying DisableWAM if supported: {0}" -f $_) }

#>
    # 3) Try DisableWAM (EXO 3.7.2+)
    try {
      if ($isDomain) { Connect-ExchangeOnline -UserPrincipalName $UserUpn -Organization $TenantHint -DisableWAM -ShowBanner:$false -ErrorAction Stop | Out-Null }
      else           { Connect-ExchangeOnline -UserPrincipalName $UserUpn -DisableWAM -ShowBanner:$false -ErrorAction Stop | Out-Null }
      return
    } catch {
      throw ("Failed to connect to Exchange Online after all fallbacks: {0}" -f $_)
    }



  }
}

function Ensure-Graph {
  param([string[]]$Scopes)
  Ensure-ModuleLoaded -Name Microsoft.Graph -MinVersion ([Version]'2.16.0')
  try {
    $ctx = (Get-MgContext) 2>$null
    $need = $true
    if ($ctx) {
      $haveScopes = @($ctx.Scopes)
      $missing = @($Scopes | Where-Object { $_ -notin $haveScopes })
      if ($missing.Count -eq 0) { $need = $false }
    }
    if ($need) {
      Act ("Connecting to Microsoft Graph with scopes: {0}" -f ($Scopes -join ', '))
      if ($TenantHint) {
        Connect-MgGraph -Scopes $Scopes -TenantId $TenantHint -NoWelcome | Out-Null
      } else {
        Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null
      }
    }
  } catch {
    throw ("Failed to connect to Microsoft Graph: {0}" -f $_)
  }
}

function Ensure-ADLocal {
  $ad = Get-Module -ListAvailable -Name ActiveDirectory | Select-Object -First 1
  if ($ad) { Import-Module ActiveDirectory -ErrorAction Stop; return $true }
  return $false
}

# ---------- utility ----------
function Resolve-GraphUser {
  param([Parameter(Mandatory=$true)][string]$Identity)
  try {
    $u = Get-MgUser -UserId $Identity -Property "id,userPrincipalName,displayName,mail" -ErrorAction Stop
  } catch {
    $u = Get-MgUser -Filter ("userPrincipalName eq '{0}'" -f $Identity) -Property "id,userPrincipalName,displayName,mail" -ErrorAction SilentlyContinue
    if (-not $u) { $u = Get-MgUser -Filter ("mail eq '{0}'" -f $Identity) -Property "id,userPrincipalName,displayName,mail" -ErrorAction SilentlyContinue }
  }
  if (-not $u) { throw ("Cannot resolve user '{0}' in Graph." -f $Identity) }
  return $u
}

# ---------- connect services ----------
$graphScopes = @('User.ReadWrite.All','Group.ReadWrite.All','GroupMember.ReadWrite.All','Directory.Read.All','AuditLog.Read.All')
if ($UseElevatedGraphScopes) { $graphScopes += 'Directory.ReadWrite.All' }

try { Ensure-Graph -Scopes $graphScopes } catch {
  Skip ("Failed to connect to Microsoft Graph: {0}" -f $_)
  Stop-Transcript | Out-Null
  Write-Host "Cannot proceed without Graph connectivity. Exiting." -ForegroundColor Red
  return
}
try { Ensure-EXO } catch {
  Skip ("Failed to connect to Exchange Online: {0}" -f $_)
  try { Disconnect-MgGraph | Out-Null } catch {}
  Stop-Transcript | Out-Null
  Write-Host "Cannot proceed without Exchange Online. Exiting." -ForegroundColor Red
  return
}

# ---------- locate the user ----------
Step ("Locating user '{0}'" -f $UserUpn)
try { $User = Resolve-GraphUser -Identity $UserUpn } catch {
  Skip ("Failed to locate user '{0}' in Graph: {1}" -f $UserUpn, $_)
  Stop-Transcript | Out-Null
  Write-Host "Cannot proceed without a valid target user. Exiting." -ForegroundColor Red
  return
}

# ---------- snapshot helpers ----------
function Snapshot-GraphGroups {
  param([string]$UserId)
  $out = @()

  # Try typed endpoint first; if unavailable, fall back to generic + verify
  $groups = $null
  try {
    $groups = Get-MgUserMemberOfAsGroup -UserId $UserId -All -ErrorAction Stop
  } catch {
    try {
      $all = Get-MgUserMemberOf -UserId $UserId -All -ErrorAction Stop
      foreach ($o in $all) {
        try {
          $gg = Get-MgGroup -GroupId $o.Id -Property "id,displayName" -ErrorAction SilentlyContinue
          if ($gg) { $groups += $gg }
        } catch { }
      }
    } catch {
      Skip ("Failed to retrieve Graph group memberships for user {0}: {1}" -f $UserId, $_)
      return @()
    }
  }

  foreach ($g in $groups) {
    try {
      $gg = Get-MgGroup -GroupId $g.Id -Property "id,displayName,groupTypes,securityEnabled,mail,mailEnabled,membershipRule,membershipRuleProcessingState" -ErrorAction SilentlyContinue
      if ($gg) {
        $isDynamic = -not [string]::IsNullOrEmpty($gg.membershipRule) -or ($gg.groupTypes -contains 'DynamicMembership')
        $isUnified = ($gg.groupTypes -contains 'Unified')
        $out += [pscustomobject]@{
          GroupId     = $gg.Id
          DisplayName = $gg.DisplayName
          Mail        = $gg.Mail
          MailEnabled = [bool]$gg.MailEnabled
          IsSecurity  = [bool]$gg.SecurityEnabled
          IsUnified   = $isUnified
          IsDynamic   = $isDynamic
        }
      }
    } catch { Skip ("Failed to read group {0}: {1}" -f $g.Id, $_) }
  }
  if ((CountOf $out) -gt 0) { return $out | Sort-Object DisplayName } else { return @() }
}

function Snapshot-GraphOwnedGroups {
  param([string]$UserId)
  $out = @()
  $ownedGroups = $null
  try {
    $ownedGroups = Get-MgUserOwnedObject -UserId $UserId -All -ErrorAction Stop
  } catch {
    Skip ("Failed to retrieve Graph owned objects for user {0}: {1}" -f $UserId, $_)
    return @()
  }
  foreach ($o in $ownedGroups) {
    try {
      $gg = Get-MgGroup -GroupId $o.Id -Property "id,displayName,groupTypes" -ErrorAction SilentlyContinue
      if ($gg) {
        $owners = Get-MgGroupOwner -GroupId $gg.Id -All -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }
        $out += [pscustomobject]@{
          GroupId     = $gg.Id
          DisplayName = $gg.DisplayName
          OwnersCount = (CountOf $owners)
          IsUnified   = ($gg.GroupTypes -contains 'Unified')
        }
      }
    } catch { }
  }
  return $out
}

function Snapshot-EXO-DLs {
  param([string]$UserSmtp)
  $dlMatches = @()
  try { $dls = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop } catch {
    Skip ("Unable to enumerate distribution groups in EXO: {0}" -f $_)
    return @()
  }

  foreach ($dl in $dls) {
    $isDynamic = $dl.RecipientTypeDetails -eq 'DynamicDistributionGroup'
    try {
      if (-not $isDynamic) {
        $members = Get-DistributionGroupMember -Identity $dl.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue
        if ($members | Where-Object { $_.PrimarySmtpAddress -ieq $UserSmtp }) {
          $dlMatches += [pscustomobject]@{
            DisplayName = $dl.DisplayName
            PrimarySmtp = $dl.PrimarySmtpAddress
            IsDynamic   = $false
          }
        }
      } else {
        $dlMatches += [pscustomobject]@{
          DisplayName = $dl.DisplayName
          PrimarySmtp = $dl.PrimarySmtpAddress
          IsDynamic   = $true
        }
      }
    } catch { Skip ("DL scan failed for {0}: {1}" -f $dl.DisplayName, $_) }
  }
  if ((CountOf $dlMatches) -gt 0) { return $dlMatches | Sort-Object DisplayName } else { return @() }
}

function Snapshot-EXO-Delegations {
  param([string]$UserSmtp)
  $out = @()
  try { $mbx = Get-Mailbox -Identity $UserSmtp -ErrorAction Stop } catch { return @() }

  try {
    $fa = Get-MailboxPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
          Where-Object { -not $_.IsInherited -and $_.User -notmatch 'NT AUTHORITY\\SELF' -and $_.AccessRights -contains 'FullAccess' }
    foreach ($p in $fa) { $out += [pscustomobject]@{ Mailbox=$mbx.PrimarySmtpAddress; Right='FullAccess'; Trustee=$p.User } }
  } catch {}

  try {
    $sa = Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
          Where-Object { -not $_.IsInherited -and $_.Trustee -ne 'NT AUTHORITY\SELF' -and $_.AccessRights -contains 'SendAs' }
    foreach ($p in $sa) { $out += [pscustomobject]@{ Mailbox=$mbx.PrimarySmtpAddress; Right='SendAs'; Trustee=$p.Trustee } }
  } catch {}

  try {
    $sob = (Get-Mailbox -Identity $mbx.Identity -Property GrantSendOnBehalfTo -ErrorAction SilentlyContinue).GrantSendOnBehalfTo
    foreach ($t in $sob) { $out += [pscustomobject]@{ Mailbox=$mbx.PrimarySmtpAddress; Right='SendOnBehalf'; Trustee=$t.PrimarySmtpAddress } }
  } catch {}

  if ((CountOf $out) -gt 0) { return $out | Sort-Object Mailbox, Right, Trustee } else { return @() }
}

function Snapshot-Licenses {
  param([string]$UserId)
  $lic = $null
  try {
    $lic = Get-MgUserLicenseDetail -UserId $UserId -All -ErrorAction Stop
  } catch {
    if ($_.Exception.Message -match 'No such host') {
      Skip "Graph endpoint unreachable for license detail. Retrying connection..."
      try { Disconnect-MgGraph -Confirm:$false | Out-Null } catch {}
      try {
        Ensure-Graph -Scopes $script:graphScopes
        $lic = Get-MgUserLicenseDetail -UserId $UserId -All -ErrorAction Stop
        Did "Reconnected to Graph and retrieved license details"
      } catch {
        Skip ("License detail fetch failed after reconnect: {0}" -f $_)
      }
    } else {
      Skip ("License snapshot failed: {0}" -f $_)
    }
  }
  if (-not $lic) { return @() }
  $rows = @()
  foreach ($l in $lic) {
    $rows += [pscustomobject]@{
      SkuId   = $l.SkuId
      SkuPart = $l.SkuPartNumber
      Svc     = ($l.ServicePlans | Where-Object { $_.ProvisioningStatus -eq 'Success' } | Select-Object -ExpandProperty ServicePlanName) -join ';'
    }
  }
  return As-Array $rows
}

function Snapshot-ADGroups {
  param([string]$SamAccountName)
  $out = @()
  try {
    $adUser = Get-ADUser -Identity $SamAccountName -Properties MemberOf,Description,Enabled,DistinguishedName -ErrorAction Stop
    $out += [pscustomobject]@{
      AD_Enabled     = $adUser.Enabled
      AD_Description = $adUser.Description
    }
    foreach ($dn in $adUser.MemberOf) {
      try {
        $g = Get-ADGroup -Identity $dn -ErrorAction SilentlyContinue
        if ($g) { $out += [pscustomobject]@{ GroupName=$g.Name; DistinguishedName=$g.DistinguishedName } }
      } catch { }
    }
  } catch {
    Skip "AD snapshot skipped (module not available or user not found)."
  }
  return $out
}

# ---------- BEFORE snapshot ----------
Step "Snapshot BEFORE"
$Before = [ordered]@{
  Identity = @{ DisplayName=$User.DisplayName; UPN=$User.UserPrincipalName; Id=$User.Id; Mail=$User.Mail }
  EXO      = [ordered]@{}
  Graph    = [ordered]@{}
  Licenses = @()
  AD       = @()
}

# Check mailbox
try {
  $mbx = Get-Mailbox -Identity $UserUpn -ErrorAction Stop
  $Before.EXO.Mailbox = @{
    PrimarySmtp          = $mbx.PrimarySmtpAddress
    RecipientTypeDetails = $mbx.RecipientTypeDetails
    CustomAttribute15    = $mbx.CustomAttribute15
  }
} catch { $mbx = $null; $Before.EXO.Mailbox = $null }

# Gather memberships and delegations
$Before.EXO.DLs         = Snapshot-EXO-DLs         -UserSmtp $UserUpn
$Before.EXO.Delegations = Snapshot-EXO-Delegations -UserSmtp $UserUpn
$Before.Graph.Groups    = Snapshot-GraphGroups     -UserId   $User.Id
$Before.Graph.Owns      = Snapshot-GraphOwnedGroups -UserId  $User.Id
$Before.Licenses        = Snapshot-Licenses        -UserId   $User.Id

# Normalize to arrays for safe export
$Before.EXO.DLs         = As-Array $Before.EXO.DLs
$Before.EXO.Delegations = As-Array $Before.EXO.Delegations
$Before.Graph.Groups    = As-Array $Before.Graph.Groups
$Before.Licenses        = As-Array $Before.Licenses

# Optional AD snapshot
$HaveAD = $false
if ($DisableAD -or $UpdateAdDescription -or $DisabledOuDn) {
  $HaveAD = Ensure-ADLocal
  if ($HaveAD) {
    $sam = ($User.UserPrincipalName -split '@')[0]
    $Before.AD = Snapshot-ADGroups -SamAccountName $sam
  } else {
    Skip "AD module not available locally — AD steps will be skipped."
  }
}

# Write BEFORE snapshots
Write-CsvSafe -Data $Before.EXO.DLs         -Path (Join-Path $OutputFolder 'Before-EXO-DLs.csv')         -Headers @('DisplayName','PrimarySmtp','IsDynamic')
Write-CsvSafe -Data $Before.EXO.Delegations -Path (Join-Path $OutputFolder 'Before-EXO-Delegations.csv') -Headers @('Mailbox','Right','Trustee')
Write-CsvSafe -Data $Before.Graph.Groups    -Path (Join-Path $OutputFolder 'Before-Graph-Groups.csv')    -Headers @('GroupId','DisplayName','Mail','MailEnabled','IsSecurity','IsUnified','IsDynamic')
Write-CsvSafe -Data $Before.Licenses        -Path (Join-Path $OutputFolder 'Before-Licenses.csv')        -Headers @('SkuId','SkuPart','Svc')
if ($Before.AD) { $Before.AD | Export-Csv -Path (Join-Path $OutputFolder 'Before-AD.csv') -NoTypeInformation -Encoding UTF8 }

# ---------- build plan ----------
$Plan = New-Object System.Collections.Generic.List[object]
function Add-Plan { param([string]$Area,[string]$Action) $Plan.Add([pscustomobject]@{ Area=$Area; Action=$Action }) }

$willConvert = $ConvertMailboxToShared -and $mbx -and $mbx.RecipientTypeDetails -notlike '*SharedMailbox*'
if ($willConvert) { Add-Plan 'Mailbox'  ("Convert mailbox to Shared and stamp expiry +{0} days in CustomAttribute15" -f $SharedMailboxExpiryDays) }
if ($SupervisorUpn) {
  if ($GrantSupervisorFullAccess) { Add-Plan 'Mailbox' ("Grant FullAccess to {0}" -f $SupervisorUpn) }
  if ($GrantSupervisorSendAs)     { Add-Plan 'Mailbox' ("Grant SendAs to {0}" -f $SupervisorUpn)     }
}

$staticDLs        = @($Before.EXO.DLs | Where-Object { $_ -and (Has-Prop $_ 'IsDynamic') -and -not $_.IsDynamic })
$staticDLCount    = CountOf $staticDLs
if ($RemoveFromDistributionLists -and $staticDLCount -gt 0) {
  Add-Plan 'EXO/DLs' ("Remove user from {0} static distribution / mail-enabled security groups" -f $staticDLCount)
}

$graphStatic      = @($Before.Graph.Groups | Where-Object { $_ -and (Has-Prop $_ 'IsDynamic') -and -not $_.IsDynamic })
$graphStaticCount = CountOf $graphStatic
if ($RemoveFromGroups -and $graphStaticCount -gt 0) {
  Add-Plan 'Graph/Groups' ("Remove user from {0} static M365/Security groups" -f $graphStaticCount)
}

$delegCount = CountOf $Before.EXO.Delegations
if ($RemoveMailboxDelegations -and $delegCount -gt 0) {
  Add-Plan 'Mailbox' ("Remove {0} mailbox delegation entries" -f $delegCount)
}

if ((CountOf $Before.Licenses) -gt 0 -and $RemoveLicenses) { Add-Plan 'Licensing' "Remove all assigned licenses" }

if ($DisableEntraSignIn) { Add-Plan 'Entra' "Block sign-in and revoke refresh tokens" }

if ($HaveAD -and ($DisableAD -or $UpdateAdDescription -or $DisabledOuDn)) {
  Add-Plan 'AD' ("Apply AD actions (Disable={0}; UpdateDesc={1}; MoveToOU='{2}')" -f $DisableAD, $UpdateAdDescription, $DisabledOuDn)
}

# Write plan
$planPath = Join-Path $OutputFolder 'Plan-WhatWeWillDo.md'
"Offboarding plan for $($User.DisplayName) <$($User.UserPrincipalName)>" | Out-File $planPath -Encoding utf8
"Ticket: $TicketNumber"            | Out-File $planPath -Append -Encoding utf8
"Preview mode: $Preview"           | Out-File $planPath -Append -Encoding utf8
""                                 | Out-File $planPath -Append -Encoding utf8
foreach ($p in $Plan) { "- [$($p.Area)] $($p.Action)" | Out-File $planPath -Append -Encoding utf8 }

# ---------- APPLY ----------
if ($Preview) {
  Step "Preview mode — no changes will be made. Use -Apply to execute changes."
} else {
  Step "Applying changes"

  # Mailbox conversion & expiry stamping
  if ($willConvert) {
    try {
      Act "Converting mailbox to Shared..."
      Set-Mailbox -Identity $UserUpn -Type Shared -ErrorAction Stop
      Did "Converted to Shared"
    } catch { Skip ("Mailbox conversion failed: {0}" -f $_) }
  }

  # Stamp expiry in CustomAttribute15
  if ($mbx) {
    try {
      $expiry = (Get-Date).AddDays([int]$SharedMailboxExpiryDays)
      $marker = ("Expires: {0:yyyy-MM-dd} ({1}d)" -f $expiry, $SharedMailboxExpiryDays)
      Act ("Stamping CustomAttribute15 with '{0}'" -f $marker)
      Set-Mailbox -Identity $UserUpn -CustomAttribute15 $marker -ErrorAction Stop
      Did "Stamped mailbox CustomAttribute15"
    } catch { Skip ("Failed to stamp CustomAttribute15: {0}" -f $_) }
  }

  # Supervisor rights
  if ($SupervisorUpn -and $mbx) {
    if ($GrantSupervisorFullAccess) {
      try {
        Act ("Grant FullAccess to {0}" -f $SupervisorUpn)
        Add-MailboxPermission -Identity $UserUpn -User $SupervisorUpn -AccessRights FullAccess -AutoMapping:$true -Confirm:$false
        Did "Granted FullAccess"
      } catch { Skip ("Grant FullAccess failed: {0}" -f $_) }
    }
    if ($GrantSupervisorSendAs) {
      try {
        Act ("Grant SendAs to {0}" -f $SupervisorUpn)
        Add-RecipientPermission -Identity $UserUpn -Trustee $SupervisorUpn -AccessRights SendAs -Confirm:$false
        Did "Granted SendAs"
      } catch { Skip ("Grant SendAs failed: {0}" -f $_) }
    }
  }

  # Remove from EXO DLs (static only)
  if ($RemoveFromDistributionLists -and $staticDLCount -gt 0) {
    foreach ($g in $staticDLs) {
      try {
        Act ("Remove from DL: {0} <{1}>" -f $g.DisplayName, $g.PrimarySmtp)
        Remove-DistributionGroupMember -Identity $g.PrimarySmtp -Member $UserUpn -BypassSecurityGroupManagerCheck -Confirm:$false
        Did ("Removed from EXO DL: {0}" -f $g.DisplayName)
      } catch { Skip ("DL removal failed for {0}: {1}" -f $g.DisplayName, $_) }
    }
  }

  # Remove mailbox delegations
  if ($RemoveMailboxDelegations -and $delegCount -gt 0) {
    foreach ($d in $Before.EXO.Delegations) {
      try {
        switch ($d.Right) {
          'FullAccess'   { Remove-MailboxPermission   -Identity $d.Mailbox -User $UserUpn -AccessRights FullAccess -Confirm:$false }
          'SendAs'       { Remove-RecipientPermission -Identity $d.Mailbox -Trustee $UserUpn -AccessRights SendAs -Confirm:$false }
          'SendOnBehalf' { Set-Mailbox -Identity $d.Mailbox -GrantSendOnBehalfTo @{ Remove = $UserUpn } }
        }
        Did ("Removed mailbox delegation: {0} on {1}" -f $d.Right, $d.Mailbox)
      } catch { Skip ("Delegation removal failed for {0} [{1}]: {2}" -f $d.Mailbox, $d.Right, $_) }
    }
  }

  # Backup owner on groups where the user is sole owner
  if ($BackupOwnerUpn) {
    try {
      $backup = Resolve-GraphUser -Identity $BackupOwnerUpn
      $soleOwner = $Before.Graph.Owns | Where-Object { $_.OwnersCount -le 1 }
      foreach ($o in $soleOwner) {
        try {
          Act ("Adding backup owner '{0}' to group: {1}" -f $backup.UserPrincipalName, $o.DisplayName)
          Add-MgGroupOwnerByRef -GroupId $o.GroupId -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($backup.Id)" } | Out-Null
          Did ("Added backup owner to {0}" -f $o.DisplayName)
        } catch { Skip ("Failed to add backup owner to {0}: {1}" -f $o.DisplayName, $_) }
      }
    } catch { Skip ("Backup owner '{0}' not resolved in Graph: {1}" -f $BackupOwnerUpn, $_) }
  }

  # Remove from static Graph groups
  if ($RemoveFromGroups -and $graphStaticCount -gt 0) {
    foreach ($g in $graphStatic) {
      try {
        Act ("Remove from Graph group: {0}" -f $g.DisplayName)
        Remove-MgGroupMemberByRef -GroupId $g.GroupId -DirectoryObjectId $User.Id -Confirm:$false
        Did ("Removed from group: {0}" -f $g.DisplayName)
      } catch { Skip ("Graph group removal failed for {0}: {1}" -f $g.DisplayName, $_) }
    }
  }

  # Licenses
  if ($RemoveLicenses -and (CountOf $Before.Licenses) -gt 0) {
    try {
      $toRemove = @($Before.Licenses | Select-Object -ExpandProperty SkuId)
      Act ("Removing licenses: {0}" -f (@($Before.Licenses | ForEach-Object {$_.SkuPart}) -join ', '))
      Set-MgUserLicense -UserId $User.Id -RemoveLicenses $toRemove -AddLicenses @() | Out-Null
      Did "Removed all licenses"
    } catch { Skip ("License removal failed: {0}" -f $_) }
  }

  # Entra sign-in
  if ($DisableEntraSignIn) {
    try {
      Act "Blocking sign-in (accountEnabled=false) and revoking sessions"
      Update-MgUser -UserId $User.Id -AccountEnabled:$false | Out-Null
      Revoke-MgUserSignInSession -UserId $User.Id | Out-Null
      Did "Blocked sign-in & revoked sessions"
    } catch { Skip ("Failed to block sign-in: {0}" -f $_) }
  }

  # AD (optional - skipped by default)
  if ($HaveAD -and ($DisableAD -or $UpdateAdDescription -or $DisabledOuDn)) {
    try {
      $sam = ($User.UserPrincipalName -split '@')[0]
      $adUser = Get-ADUser -Identity $sam -Properties Enabled,Description,DistinguishedName -ErrorAction Stop
      if ($DisableAD -and $adUser.Enabled) {
        Act "Disabling AD account"
        Disable-ADAccount -Identity $adUser.SamAccountName
        Did "AD account disabled"
      }
      if ($UpdateAdDescription) {
        $desc = ("Offboarded {0}; Ticket {1}" -f (Get-Date -Format 'yyyy-MM-dd'), $TicketNumber)
        Act ("Updating AD description to '{0}'" -f $desc)
        Set-ADUser -Identity $adUser.SamAccountName -Description $desc
        Did "AD description updated"
      }
      if ($DisabledOuDn) {
        Act ("Moving user to Disabled OU: {0}" -f $DisabledOuDn)
        Move-ADObject -Identity $adUser.DistinguishedName -TargetPath $DisabledOuDn
        Did ("Moved AD object to '{0}'" -f $DisabledOuDn)
      }
    } catch { Skip ("AD actions failed: {0}" -f $_) }
  }
}

# ---------- AFTER snapshot ----------
Step "Snapshot AFTER"
$After = [ordered]@{
  EXO      = [ordered]@{}
  Graph    = [ordered]@{}
  Licenses = @()
  AD       = @()
}

try {
  $mbx2 = Get-Mailbox -Identity $UserUpn -ErrorAction SilentlyContinue
  if ($mbx2) {
    $After.EXO.Mailbox = @{
      PrimarySmtp          = $mbx2.PrimarySmtpAddress
      RecipientTypeDetails = $mbx2.RecipientTypeDetails
      CustomAttribute15    = $mbx2.CustomAttribute15
    }
  }
} catch { }

$After.EXO.DLs         = Snapshot-EXO-DLs         -UserSmtp $UserUpn
$After.EXO.Delegations = Snapshot-EXO-Delegations -UserSmtp $UserUpn
$After.Graph.Groups    = Snapshot-GraphGroups     -UserId   $User.Id
$After.Licenses        = Snapshot-Licenses        -UserId   $User.Id
if ($HaveAD) {
  $sam = ($User.UserPrincipalName -split '@')[0]
  $After.AD = Snapshot-ADGroups -SamAccountName $sam
}

# Normalize to arrays for safe export
$After.EXO.DLs         = As-Array $After.EXO.DLs
$After.EXO.Delegations = As-Array $After.EXO.Delegations
$After.Graph.Groups    = As-Array $After.Graph.Groups
$After.Licenses        = As-Array $After.Licenses

# Write AFTER snapshots
Write-CsvSafe -Data $After.EXO.DLs         -Path (Join-Path $OutputFolder 'After-EXO-DLs.csv')           -Headers @('DisplayName','PrimarySmtp','IsDynamic')
Write-CsvSafe -Data $After.EXO.Delegations -Path (Join-Path $OutputFolder 'After-EXO-Delegations.csv')   -Headers @('Mailbox','Right','Trustee')
Write-CsvSafe -Data $After.Graph.Groups    -Path (Join-Path $OutputFolder 'After-Graph-Groups.csv')      -Headers @('GroupId','DisplayName','Mail','MailEnabled','IsSecurity','IsUnified','IsDynamic')
Write-CsvSafe -Data $After.Licenses        -Path (Join-Path $OutputFolder 'After-Licenses.csv')          -Headers @('SkuId','SkuPart','Svc')
if ($After.AD) { $After.AD | Export-Csv -Path (Join-Path $OutputFolder 'After-AD.csv') -NoTypeInformation -Encoding UTF8 }

# ---------- ServiceNow work notes ----------
function Summ($label,$before,$after) {
  $b = CountOf $before
  $a = CountOf $after
  return ("{0}: {1} -> {2}" -f $label, $b, $a)
}

$notesPath = Join-Path $OutputFolder ('ServiceNow-WorkNotes.txt')

# Guarded dynamic/static splits
$b_staticDL = @($Before.EXO.DLs | Where-Object { $_ -and (Has-Prop $_ 'IsDynamic') -and -not $_.IsDynamic })
$a_staticDL = @($After.EXO.DLs  | Where-Object { $_ -and (Has-Prop $_ 'IsDynamic') -and -not $_.IsDynamic })
$b_dynDL    = @($Before.EXO.DLs | Where-Object { $_ -and (Has-Prop $_ 'IsDynamic') -and $_.IsDynamic })
$a_dynDL    = @($After.EXO.DLs  | Where-Object { $_ -and (Has-Prop $_ 'IsDynamic') -and $_.IsDynamic })

$b_graph    = $Before.Graph.Groups
$a_graph    = $After.Graph.Groups
$a_graphDyn = CountOf (@($a_graph | Where-Object { $_ -and (Has-Prop $_ 'IsDynamic') -and $_.IsDynamic }))

$b_deleg    = $Before.EXO.Delegations
$a_deleg    = $After.EXO.Delegations

$b_lic      = $Before.Licenses
$a_lic      = $After.Licenses

$foundMailboxType = if ($Before.EXO.Mailbox -and $Before.EXO.Mailbox.RecipientTypeDetails) { $Before.EXO.Mailbox.RecipientTypeDetails } else { 'None' }
$expiryDatePreview = (Get-Date).AddDays([int]$SharedMailboxExpiryDays).ToString('yyyy-MM-dd')

$workNotes = @"
Offboarding — $($User.DisplayName) <$($User.UserPrincipalName)>
Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Ticket: $TicketNumber
Analyst: $env:USERNAME
Mode: $(if($Preview){"Preview (no changes)"}else{"Applied"})

Summary at a glance
- $(Summ "EXO DLs (static)" $b_staticDL $a_staticDL)  | dynamic listed: $a_graphDyn
- $(Summ "Graph groups (all)" $b_graph $a_graph)
- $(Summ "Mailbox delegations" $b_deleg $a_deleg)
- $(Summ "Assigned licenses" $b_lic $a_lic)
- $(if($HaveAD){"AD snapshot written"}else{"AD not executed"})

Mailbox
- Found mailbox: $foundMailboxType
- $( if ($willConvert) { "Converted to Shared (or already Shared). Expiry marker: $expiryDatePreview" } else { "No mailbox conversion requested" } )
$( if ($SupervisorUpn) {
    $rights = @(); if ($GrantSupervisorFullAccess) { $rights += 'FullAccess' } ; if ($GrantSupervisorSendAs) { $rights += 'SendAs' }
    "Supervisor access: " + ($rights -join ' & ') + " for $SupervisorUpn"
} )

Groups & DLs
- We do not remove dynamic membership. It is shown for visibility only.
- Removed from static EXO DLs: $(if($RemoveFromDistributionLists){"Yes (see 'After-EXO-DLs.csv')"}else{"No"})
- Removed from static Graph groups: $(if($RemoveFromGroups){"Yes (see 'After-Graph-Groups.csv')"}else{"No"})`n$(if($BackupOwnerUpn){"- Added backup owner '$BackupOwnerUpn' where user was sole owner"})

Licenses & Sign-in
- Licenses removed: $(if($RemoveLicenses){"Yes"}else{"No"})
- Entra sign-in blocked: $(if($DisableEntraSignIn){"Yes"}else{"No"})

Artifacts
- Before snapshots: $(Join-Path $OutputFolder 'Before-*')
- After snapshots:  $(Join-Path $OutputFolder 'After-*')
- Plan:             $(Join-Path $OutputFolder 'Plan-WhatWeWillDo.md')
- Transcript:       $TranscriptPath
"@

$workNotes | Out-File $notesPath -Encoding utf8

# ---------- disconnect & wrap ----------
try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}
try { Disconnect-MgGraph | Out-Null } catch {}
Stop-Transcript | Out-Null

Write-Host ("`nDone. Preview: {0}  Evidence folder: {1}" -f $Preview, $OutputFolder) -ForegroundColor Cyan
Write-Host ("ServiceNow notes file: {0}" -f $notesPath) -ForegroundColor Cyan
