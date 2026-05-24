# M365 License Waste Report

A small, read-only PowerShell script that finds the Microsoft 365 / Entra ID
licenses you're **paying for but not using** — and puts a dollar figure on the waste.

It connects to Microsoft Graph, looks at every subscription (SKU) in your tenant,
and for each one shows how many seats you bought, how many are actually assigned,
and how many are sitting **unused**. You get a console table and a clean HTML report.

> Read-only. It never changes anything in your tenant.

---

## Why

Seats get over-provisioned, people get offboarded, projects end — and the licenses
keep getting paid for. This gives you a one-command snapshot of where the money is
leaking, plus a tidy report you can hand to whoever signs off on renewals.

---

## Example output

```
License                          Purchased Assigned Unused Usage%  Waste/yr ($)
-------                          --------- -------- ------ ------  ------------
Microsoft 365 E5                       120      102     18  85%          12,312
Office 365 E3                           80       61     19  76%           5,244
Microsoft Entra ID P2                   50       33     17  66%           1,836
...
ESTIMATED TOTAL WASTE: 19,392 $ / year
```

The HTML report shows the same data with the total estimated annual waste front and center.


---

## Requirements

- PowerShell 5.1+ (Windows PowerShell or PowerShell 7)
- A Microsoft 365 / Entra ID tenant
- An account that can consent to or has the **`Organization.Read.All`** Graph permission
- The `Microsoft.Graph.Identity.DirectoryManagement` module (the script installs it for you on first run)

---

## Usage

```powershell
# Default: prints the table and writes an HTML report to the current folder
.\Get-M365LicenseWaste.ps1

# Choose where the HTML report goes
.\Get-M365LicenseWaste.ps1 -OutputPath "C:\Reports\licenses.html"

# Console only, no HTML file
.\Get-M365LicenseWaste.ps1 -SkipHtml
```

On the first run you'll get a Microsoft sign-in prompt to consent to the
read-only permission.

---

## Notes on pricing

The per-license prices are **estimates** based on public list prices (USD), stored
in the `$SkuPrices` table near the top of the script. Real prices depend on your
agreement, region, and currency — edit that table to match your contract and the
waste figures will follow.

Licenses without a price entry are still shown (unused count and all), they just
won't contribute to the waste total.

---

## How it works (short version)

1. Ensures the Graph module is present, then connects read-only.
2. Calls `Get-MgSubscribedSku` to list every subscription.
3. For each SKU: `Purchased = PrepaidUnits.Enabled`, `Assigned = ConsumedUnits`,
   `Unused = Purchased - Assigned`.
4. Multiplies unused seats by the (editable) unit price to estimate waste.
5. Renders a console table and an HTML report.

---

## Contributing

SKU names and prices change, and there are a lot of them. PRs welcome — especially
new entries for the `$SkuNames` / `$SkuPrices` tables, or edge cases on large/unusual
tenants. Issues and feedback equally welcome.

---

## License

MIT
