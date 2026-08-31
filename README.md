# Windows 365 Cloud PC Resize Tool

A PowerShell + Windows Forms (GUI) tool for viewing **Windows 365 Enterprise and Business**
Cloud PCs in a tenant and **resizing** (upgrading) or **downsizing** a single Cloud PC to
another available license of the **same edition**.

- Script: [`W365_ResizeUI.ps1`](./W365_ResizeUI.ps1)

---

## What it does

1. **Loads Cloud PCs** directly from the Cloud PC Graph API
   (`Get-MgBetaDeviceManagementVirtualEndpointCloudPc`) and shows every **Windows 365
   Enterprise or Business** (dedicated) Cloud PC. **Frontline/Flex (shared)** editions are
   excluded. GPU plans are excluded as resize targets.
2. Displays them in a grid with: **Cloud PC Name** (`managedDeviceName`), **User, Status,
   Edition, Current License, vCPU, RAM (GB), Disk (GB)** (plus hidden Cloud PC Id / Service
   Plan Id used internally).
3. Lets an admin **select exactly one** Cloud PC and **Resize** it:
   - Reads the tenant's **Windows 365 licenses** (`subscribedSkus`) with their **seat counts**
     (`enabled` vs. `consumed`), including modern **leveling SKUs** (e.g. `CPC_LVL_x`) whose
     config lives in a child plan, and the internal **CloudPC_Lite** add-on.
   - Presents a popup with a **single-select radio list**, restricted to the **same edition**
     as the selected Cloud PC (Enterprise→Enterprise, Business→Business).
   - A license is **selectable only if** it has seats available, is not the current license,
     is resolvable as a resize target, and passes the **downsize disk rule**.
4. Shows a **confirmation dialog** (Cloud PC name, current license, target license, and whether
   it is a **Resize (Upgrade)** or **Downsize**) before submitting.
5. Submits the resize to Graph, validates the **HTTP response status**, and shows a **Success**
   or **Failed** message (with Cloud PC name, action, target license, and a note to check the
   Intune portal). On success the Cloud PC becomes **unselectable** for the session and the
   inventory reloads.
6. Supports **filtering** (by name/user) and a **module bootstrap** prompt that installs the
   required Microsoft Graph modules if missing.

### The downsize rule
- **Downsizing** to a license whose **disk is smaller than the current disk is not allowed**
  (the option is disabled in the popup with the reason shown).
- **Upsizing** (resize up) has no disk restriction.

### Status & selectability
- Resize is only allowed when the Cloud PC status is **`provisioned`** or
  **`provisionedWithWarnings`**. On load/reload, any other status (e.g. `Resizing`, failed)
  renders the row's checkbox **unselectable**, and the resize-time guard blocks it as a second
  layer.
- Cloud PCs currently **`Resizing`** are highlighted with a **yellow row background**; the
  highlight clears automatically once the status returns to provisioned on the next reload.

---

## How it works (internals)

| Area | Detail |
|------|--------|
| Edition detection | `Get-CloudPcEdition` — returns `Enterprise`, `Business`, or `$null` (excluded). Excludes `provisioningType = shared` and Frontline/Flex; classifies by `servicePlanType`/plan name. |
| Frontline/Flex filter | `Test-IsFrontlineOrFlex` — central helper matching `frontline` or `flex` **anywhere** (any prefix/suffix), so a future "Windows 365 Flex" is always excluded. |
| License list | `Get-W365Licenses` — reads `GET /v1.0/subscribedSkus`, keeps Enterprise (`CPC_E*` / `Windows_365_Enterprise*`) and Business (`CPC_B*` / `Windows_365_Business*`); config for leveling SKUs comes from the child `CPC_E*/CPC_B*` plan. Seats = `prepaidUnits.enabled - consumedUnits`. GPU/DR/shared excluded. |
| Manual targets | `$script:manualLicenseTargets` — explicit mappings for licenses that can't be resolved from the SKU (currently **CloudPC_Lite**: service plan `6b97ad6a-…`, 2 vCPU / 4 GB / 128 GB). Excluded from the auto config map to avoid collisions. |
| Service plan map | `Get-ServicePlanMap` — `GET /beta/deviceManagement/virtualEndpoint/servicePlans`, keyed by **`<edition>\|vCPU\|RAM\|storage`**. Classifies by display name first (Frontline plans report `type=enterprise`), excludes GPU/Frontline/Flex/manual-target plans. |
| Config parsing | `Get-W365ConfigFromText` — extracts vCPU/RAM/disk from names/part numbers in several formats (`..._2_vCPU_8_GB_128_GB`, `2 vCPU/8 GB/128 GB`, `CPC_E_4C_16GB_256GB`). |
| Resize call | `POST /beta/deviceManagement/virtualEndpoint/cloudPCs/{id}/resize` with body `{ "targetServicePlanId": "<id>" }`, sent via `-OutputType HttpResponseMessage`; only a real **2xx** counts as success. |
| Upgrade vs. downsize | `Get-ResizeOperationType` compares vCPU, then RAM, then storage. |
| Same-edition rule | The popup only lists licenses whose `Edition` matches the selected Cloud PC. |
| Single selection | Enforced in the grid (`CellValueChanged` clears other checkboxes) and re-checked at click time. |
| Resilience | `Invoke-WithGraphRetry` retries throttling/5xx with exponential backoff. |
| Diagnostics | `W365_ResizeGUI.log` logs service plans, license SKUs, and each resize request/response. |

---

## Requirements

- **Windows PowerShell 5.1** or **PowerShell 7+** on Windows (uses Windows Forms).
- Microsoft Graph PowerShell modules (auto-prompted for install if missing):
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Beta.DeviceManagement`
- **Graph delegated permissions** (consented on first connect):
  - `CloudPC.Read.All`
  - `Organization.Read.All` (to read tenant license/SKU seat counts)
  - `CloudPC.ReadWrite.All` (requested only when submitting a resize)

---

## Usage

```powershell
# From the folder containing the script
.\W365_ResizeUI.ps1
```

1. Click **Load Cloud PCs** and sign in when prompted.
2. (Optional) Filter by name/user.
3. Tick **Select** on exactly one Cloud PC (only `provisioned`/`provisionedWithWarnings` rows are selectable).
4. Click **Resize Selected Cloud PC**.
5. Pick a target license (only eligible, same-edition ones are enabled) and click **Continue**.
6. Review the confirmation and click **Yes** to submit.

A log file `W365_ResizeGUI.log` is written next to the script for troubleshooting.

> The **Export CSV** button is currently hidden (`$exportButton.Visible = $false`); set it back
> to `$true` to re-enable it.

---

## Status

- ✅ Rewritten from the original stale-device tool; **syntax validates** with no parse errors.
- ✅ Intune Delete / Entra Delete actions removed.
- ✅ Load logic returns **all Enterprise and Business Cloud PCs** (no Intune stale comparison); Frontline/Flex excluded.
- ✅ Resize / Downsize workflow with license seat checks, disk downsize rule, same-edition restriction, and confirmation.
- ✅ Modern **leveling SKUs** (`CPC_LVL_x`) and the internal **CloudPC_Lite** add-on are supported (display + resize target).
- ✅ **GPU** licenses excluded as resize targets.
- ✅ Resize call verified end to end against a live tenant (status updates to **Resizing** in Intune and the UI); HTTP status is validated so failures no longer show as success.
- ✅ Status guard + row selectability (`provisioned` / `provisionedWithWarnings` only) and yellow highlight for `Resizing`.
- ✅ Success/Failure popups include Cloud PC name, action, target license, and an "check the Intune portal" note.
- ⚠️ Remaining PSScriptAnalyzer warnings are **stylistic** and inherited from the original
  (`Ensure-*` unapproved verbs, `$sender` param name).

### Notes / maintenance
- **CloudPC_Lite** and any other unresolved-by-SKU licenses are handled via
  `$script:manualLicenseTargets` at the top of the script — add entries there (SKU match,
  edition, service plan id, config) as needed.
- The **single-entity resize endpoint** (`cloudPCs/{id}/resize`) is used, matching the
  documented `cloudPC: resize` action.
- Verbose `SERVICE PLAN raw:` diagnostics remain in the log; they can be trimmed once no
  longer needed.
