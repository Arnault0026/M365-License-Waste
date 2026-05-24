<#PSScriptInfo
.VERSION 1.0.0
.GUID 7340b189-f668-45f1-9baa-791bd562af16
.AUTHOR Arnault0026
.COMPANYNAME
.COPYRIGHT (c) 2026. MIT License.
.TAGS Microsoft365 M365 EntraID AzureAD Licensing License Graph Report Cost Sysadmin
.LICENSEURI https://opensource.org/licenses/MIT
.PROJECTURI https://github.com/Arnault0026/M365-License-Waste
.RELEASENOTES
Initial release. Read-only M365/Entra ID license waste report (console + HTML).
#>

#Requires -Version 5.1

<#
.SYNOPSIS
    M365 License Waste Report — finds Microsoft 365 / Entra ID licenses you pay for
    but nobody uses (purchased units that aren't assigned).

.DESCRIPTION
    Connects to Microsoft Graph (read-only), pulls every subscribed SKU in the tenant,
    and reports per license: purchased units, assigned units, and UNUSED units, then
    estimates the monthly/yearly wasted spend. Outputs a console table plus a clean
    HTML report.

    Read-only. Makes no changes to the tenant.
    Required permission: Organization.Read.All (consented on first run).

.PARAMETER OutputPath
    Path for the HTML report. Defaults to the current folder, with today's date.

.PARAMETER SkipHtml
    Print only the console table; skip generating the HTML file.

.EXAMPLE
    .\Get-M365LicenseWaste.ps1

.EXAMPLE
    .\Get-M365LicenseWaste.ps1 -OutputPath "C:\Reports\licenses.html"

.NOTES
    Prices are estimates (public monthly list price, USD) and ARE EDITABLE in the
    $SkuPrices table below. Adjust them to your own currency / agreement.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PWD ("M365-License-Waste-{0}.html" -f (Get-Date -Format 'yyyy-MM-dd'))),
    [switch]$SkipHtml
)

# ----------------------------------------------------------------------------
# 1. Make sure the required Microsoft Graph module is installed
# ----------------------------------------------------------------------------
$RequiredModule = 'Microsoft.Graph.Identity.DirectoryManagement'
if (-not (Get-Module -ListAvailable -Name $RequiredModule)) {
    Write-Host "Installing module $RequiredModule (first run only)..." -ForegroundColor Yellow
    try {
        Install-Module $RequiredModule -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    } catch {
        Write-Error "Could not install $RequiredModule : $_"
        return
    }
}
Import-Module $RequiredModule -ErrorAction Stop

# ----------------------------------------------------------------------------
# 2. Map SKU part numbers to human-readable names (most common ones)
# ----------------------------------------------------------------------------
$SkuNames = @{
    'ENTERPRISEPACK'            = 'Office 365 E3'
    'ENTERPRISEPREMIUM'         = 'Office 365 E5'
    'STANDARDPACK'              = 'Office 365 E1'
    'SPE_E3'                    = 'Microsoft 365 E3'
    'SPE_E5'                    = 'Microsoft 365 E5'
    'SPE_F1'                    = 'Microsoft 365 F3'
    'SPB'                       = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'     = 'Microsoft 365 Business Standard'
    'O365_BUSINESS_ESSENTIALS'  = 'Microsoft 365 Business Basic'
    'O365_BUSINESS'             = 'Microsoft 365 Apps for Business'
    'OFFICESUBSCRIPTION'        = 'Microsoft 365 Apps for Enterprise'
    'EXCHANGESTANDARD'          = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'        = 'Exchange Online (Plan 2)'
    'EMS'                       = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM'               = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2'            = 'Microsoft Entra ID P2'
    'POWER_BI_PRO'              = 'Power BI Pro'
    'POWER_BI_STANDARD'         = 'Power BI (free)'
    'FLOW_FREE'                 = 'Power Automate (free)'
    'TEAMS_EXPLORATORY'         = 'Teams Exploratory'
    'MCOEV'                     = 'Microsoft Teams Phone'
    'MCOMEETADV'                = 'Microsoft 365 Audio Conferencing'
    'VISIOCLIENT'               = 'Visio Plan 2'
    'PROJECTPROFESSIONAL'       = 'Project Plan 3'
    'WIN_DEF_ATP'               = 'Defender for Endpoint P2'
    'DESKLESSPACK'              = 'Office 365 F3'
}

# ----------------------------------------------------------------------------
# 3. Estimated monthly prices (USD, public list) -> EDIT THESE
#    Key = SkuPartNumber. Adjust to your currency / contract.
# ----------------------------------------------------------------------------
$SkuPrices = @{
    'ENTERPRISEPACK'           = 23.00
    'ENTERPRISEPREMIUM'        = 38.00
    'STANDARDPACK'             = 10.00
    'SPE_E3'                   = 36.00
    'SPE_E5'                   = 57.00
    'SPE_F1'                   = 8.00
    'SPB'                      = 22.00
    'O365_BUSINESS_PREMIUM'    = 12.50
    'O365_BUSINESS_ESSENTIALS' = 6.00
    'O365_BUSINESS'            = 8.25
    'OFFICESUBSCRIPTION'       = 12.00
    'EXCHANGESTANDARD'         = 4.00
    'EXCHANGEENTERPRISE'       = 8.00
    'EMS'                      = 10.60
    'EMSPREMIUM'               = 16.40
    'AAD_PREMIUM'              = 6.00
    'AAD_PREMIUM_P2'           = 9.00
    'POWER_BI_PRO'             = 10.00
    'MCOEV'                    = 8.00
    'VISIOCLIENT'              = 15.00
    'PROJECTPROFESSIONAL'      = 30.00
}

function Get-FriendlyName([string]$PartNumber) {
    if ($SkuNames.ContainsKey($PartNumber)) { return $SkuNames[$PartNumber] }
    return $PartNumber
}

# ----------------------------------------------------------------------------
# 4. Connect to Microsoft Graph (read-only)
# ----------------------------------------------------------------------------
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes 'Organization.Read.All' -NoWelcome -ErrorAction Stop
} catch {
    Write-Error "Connection failed: $_"
    return
}

# ----------------------------------------------------------------------------
# 5. Pull and analyze the SKUs
# ----------------------------------------------------------------------------
Write-Host "Retrieving subscriptions..." -ForegroundColor Cyan
$skus = Get-MgSubscribedSku -All

$report = foreach ($sku in $skus) {
    $purchased = [int]$sku.PrepaidUnits.Enabled
    $assigned  = [int]$sku.ConsumedUnits
    $unused    = [math]::Max(0, $purchased - $assigned)

    $unitPrice = if ($SkuPrices.ContainsKey($sku.SkuPartNumber)) { $SkuPrices[$sku.SkuPartNumber] } else { $null }
    $wasteMonthly = if ($unitPrice) { [math]::Round($unused * $unitPrice, 2) } else { $null }

    [pscustomobject]@{
        License       = Get-FriendlyName $sku.SkuPartNumber
        PartNumber    = $sku.SkuPartNumber
        Purchased     = $purchased
        Assigned      = $assigned
        Unused        = $unused
        UsagePercent  = if ($purchased -gt 0) { [math]::Round(($assigned / $purchased) * 100, 0) } else { 0 }
        WasteMonthly  = $wasteMonthly
        WasteYearly   = if ($wasteMonthly) { [math]::Round($wasteMonthly * 12, 2) } else { $null }
    }
}

$report = $report | Sort-Object -Property Unused -Descending
$totalWasteYear = ($report | Where-Object { $_.WasteYearly } | Measure-Object -Property WasteYearly -Sum).Sum

# ----------------------------------------------------------------------------
# 6. Console output
# ----------------------------------------------------------------------------
Write-Host ""
$report | Format-Table License, Purchased, Assigned, Unused,
    @{N='Usage%';E={"$($_.UsagePercent)%"}},
    @{N='Waste/yr ($)';E={ if ($_.WasteYearly) { '{0:N0}' -f $_.WasteYearly } else { 'n/a' } }} -AutoSize

Write-Host ("ESTIMATED TOTAL WASTE: {0:N0} $ / year" -f $totalWasteYear) -ForegroundColor Green
Write-Host "(prices are estimates - adjust the `$SkuPrices table to your contract)" -ForegroundColor DarkGray

# ----------------------------------------------------------------------------
# 7. HTML report
# ----------------------------------------------------------------------------
if (-not $SkipHtml) {
    $rows = foreach ($r in $report) {
        $cls = if ($r.UsagePercent -lt 70) { 'warn' } elseif ($r.UsagePercent -lt 90) { 'mid' } else { 'ok' }
        $wy  = if ($r.WasteYearly) { '{0:N0} $' -f $r.WasteYearly } else { '—' }
        "<tr class='$cls'><td>$($r.License)</td><td>$($r.Purchased)</td><td>$($r.Assigned)</td><td><b>$($r.Unused)</b></td><td>$($r.UsagePercent)%</td><td>$wy</td></tr>"
    }

    $html = @"
<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'>
<title>M365 License Waste Report</title>
<style>
 body{font-family:Segoe UI,Arial,sans-serif;background:#f4f6f9;color:#1f2937;margin:0;padding:32px}
 .wrap{max-width:900px;margin:auto;background:#fff;border-radius:14px;box-shadow:0 4px 24px rgba(0,0,0,.06);padding:36px}
 h1{font-size:22px;margin:0 0 4px}.sub{color:#6b7280;font-size:13px;margin-bottom:24px}
 .hero{background:linear-gradient(135deg,#0f766e,#14b8a6);color:#fff;border-radius:12px;padding:22px 26px;margin-bottom:28px}
 .hero .big{font-size:34px;font-weight:700}.hero .lbl{opacity:.9;font-size:13px;text-transform:uppercase;letter-spacing:.5px}
 table{width:100%;border-collapse:collapse;font-size:14px}
 th{text-align:left;background:#f9fafb;padding:10px 12px;border-bottom:2px solid #e5e7eb;font-size:12px;text-transform:uppercase;color:#6b7280}
 td{padding:10px 12px;border-bottom:1px solid #f0f1f3}
 tr.warn td{background:#fef2f2}tr.mid td{background:#fffbeb}
 .foot{margin-top:22px;font-size:12px;color:#9ca3af}
</style></head><body><div class='wrap'>
 <h1>Microsoft 365 License Waste Report</h1>
 <div class='sub'>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm')</div>
 <div class='hero'><div class='lbl'>Estimated waste</div><div class='big'>$('{0:N0}' -f $totalWasteYear) $ / year</div></div>
 <table><thead><tr><th>License</th><th>Purchased</th><th>Assigned</th><th>Unused</th><th>Usage</th><th>Waste/yr</th></tr></thead>
 <tbody>$($rows -join "`n")</tbody></table>
 <div class='foot'>Red rows = under 70% usage. Prices are estimates (public list) — adjust to your contract. Data is read-only via Microsoft Graph.</div>
</div></body></html>
"@

    $html | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host ("HTML report generated: {0}" -f $OutputPath) -ForegroundColor Green
}

Disconnect-MgGraph | Out-Null
