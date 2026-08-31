# Cloud PC (Windows 365 Enterprise) Resize Tool
# Lists all Windows 365 Enterprise Cloud PCs from the Cloud PC Graph API and lets an
# admin resize (upgrade) or downsize a single selected Cloud PC to another available
# Windows 365 Enterprise license in the tenant.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global data storage
$script:allData = @()
$script:filteredData = @()
$script:servicePlans = @()
$script:suppressSelectSync = $false
$script:resizedCloudPcIds = @{}
$script:logPath = Join-Path -Path $PSScriptRoot -ChildPath "W365_ResizeGUI.log"
$script:forceModuleBootstrap = ($env:W365_FORCE_MODULE_BOOTSTRAP -eq "1")

# Manual/internal license targets that cannot be resolved automatically from the SKU
# (e.g. the CloudPC_Lite add-on, whose SKU has no config-bearing child plan).
# SkuMatch is tested (case-insensitive) against the SKU part number and its child plan
# names. These are excluded from the automatic config-triple map to avoid collisions.
$script:manualLicenseTargets = @(
    [pscustomobject]@{
        SkuMatch      = 'cloudpc[_ ]?add-?on|cloudpc[_ ]?lite'
        Edition       = 'Enterprise'
        ServicePlanId = '6b97ad6a-be15-4cbe-afbb-4eb74ecb0243'
        PlanName      = 'CloudPC_Lite'
        DisplayName   = 'CloudPC_Lite 2 vCPU / 4 GB / 128 GB'
        VCpu          = 2
        RamGB         = 4
        StorageGB     = 128
    }
)

function Get-NormalizedName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return ""
    }

    return $Name.Trim().ToLowerInvariant()
}

function Write-DebugLog {
    param([string]$Message)

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
    Add-Content -Path $script:logPath -Value "[$timestamp] $Message" -ErrorAction SilentlyContinue
}

# Reads a value by name from either a hashtable (Invoke-MgGraphRequest result) or a
# PSObject (Graph SDK object), trying each candidate key/property name in order.
function Get-DictValue {
    param(
        $Dict,
        [string[]]$Names
    )

    if ($null -eq $Dict) {
        return $null
    }

    if ($Dict -is [System.Collections.IDictionary]) {
        foreach ($key in $Dict.Keys) {
            foreach ($name in $Names) {
                if ([string]$key -ieq $name) {
                    return $Dict[$key]
                }
            }
        }
        return $null
    }

    foreach ($name in $Names) {
        $prop = $Dict.PSObject.Properties[$name]
        if ($prop) {
            return $prop.Value
        }
    }
    return $null
}

# Reads a Cloud PC property, checking direct properties then AdditionalProperties.
function Get-CloudPcPropertyValue {
    param(
        $CloudPc,
        [string[]]$Names
    )

    if ($null -eq $CloudPc) {
        return $null
    }

    foreach ($name in $Names) {
        $prop = $CloudPc.PSObject.Properties[$name]
        if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
            return $prop.Value
        }
    }

    if ($CloudPc.PSObject.Properties.Name -contains 'AdditionalProperties' -and $CloudPc.AdditionalProperties -is [System.Collections.IDictionary]) {
        foreach ($name in $Names) {
            if ($CloudPc.AdditionalProperties.Contains($name) -and -not [string]::IsNullOrWhiteSpace([string]$CloudPc.AdditionalProperties[$name])) {
                return $CloudPc.AdditionalProperties[$name]
            }
        }
    }

    return $null
}

# True if the text contains "Frontline" or "Flex" anywhere (any prefix/suffix),
# case-insensitive. Covers current and future naming (e.g. "Windows 365 Flex").
function Test-IsFrontlineOrFlex {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    return [bool]($Text -match '(?i)frontline|flex')
}

# Returns the Cloud PC edition eligible for resize: 'Enterprise', 'Business', or $null
# (excluded). Frontline/Flex (shared) Cloud PCs are excluded.
function Get-CloudPcEdition {
    param($CloudPc)

    $servicePlanType = [string](Get-CloudPcPropertyValue -CloudPc $CloudPc -Names @('servicePlanType', 'ServicePlanType'))
    $provisioningType = [string](Get-CloudPcPropertyValue -CloudPc $CloudPc -Names @('provisioningType', 'ProvisioningType'))
    $servicePlanName = [string](Get-CloudPcPropertyValue -CloudPc $CloudPc -Names @('servicePlanName', 'ServicePlanName'))

    # Frontline/Flex Cloud PCs are provisioned as shared.
    if ($provisioningType -match '(?i)shared') {
        return $null
    }

    if (-not [string]::IsNullOrWhiteSpace($servicePlanType)) {
        if ($servicePlanType -match '(?i)enterprise') { return 'Enterprise' }
        if ($servicePlanType -match '(?i)business') { return 'Business' }
        return $null
    }

    # Fallback when servicePlanType is not present: infer from the plan name.
    if ($servicePlanName -match '(?i)shared' -or (Test-IsFrontlineOrFlex $servicePlanName)) {
        return $null
    }
    if ($servicePlanName -match '(?i)business') {
        return 'Business'
    }
    if ($servicePlanName -match '(?i)enterprise') {
        return 'Enterprise'
    }

    return $null
}

# Extracts vCPU / RAM(GB) / Storage(GB) from a service plan display name or SKU part
# number. Handles multiple formats, e.g.:
#   "Windows_365_Enterprise_2_vCPU_8_GB_128_GB", "2 vCPU/8 GB/128 GB",
#   "Windows 365 Enterprise 4 vCPU, 16 GB, 256 GB (Preview)", "CPC_E_4C_16GB_256GB".
function Get-W365ConfigFromText {
    param([string]$Text)

    $result = [pscustomobject]@{ VCpu = $null; RamGB = $null; StorageGB = $null }

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $result
    }

    $normalized = $Text -replace '/', ' '

    # vCPU: either an explicit "<n> vCPU" token or the "<n>C" core form (e.g. CPC_E_4C_...).
    if ($normalized -match '(?i)(\d+)\s*_?\s*vcpu') {
        $result.VCpu = [int]$matches[1]
    }
    elseif ($normalized -match '(?i)(?:^|[_\s])(\d+)\s*c(?=[_\s\d])') {
        $result.VCpu = [int]$matches[1]
    }

    # RAM = first "<n> GB", Storage = last "<n> GB".
    $gbMatches = [regex]::Matches($normalized, '(?i)(\d+)\s*_?\s*gb')
    if ($gbMatches.Count -ge 2) {
        $result.RamGB = [int]$gbMatches[0].Groups[1].Value
        $result.StorageGB = [int]$gbMatches[$gbMatches.Count - 1].Groups[1].Value
    }
    elseif ($gbMatches.Count -eq 1) {
        $result.StorageGB = [int]$gbMatches[0].Groups[1].Value
    }

    return $result
}

# Classifies a target license relative to the current license.
function Get-ResizeOperationType {
    param($Current, $Target)

    if ($Target.VCpu -gt $Current.VCpu) { return 'Resize (Upgrade)' }
    if ($Target.VCpu -lt $Current.VCpu) { return 'Downsize' }
    if ($Target.RamGB -gt $Current.RamGB) { return 'Resize (Upgrade)' }
    if ($Target.RamGB -lt $Current.RamGB) { return 'Downsize' }
    if ($Target.StorageGB -gt $Current.StorageGB) { return 'Resize (Upgrade)' }
    if ($Target.StorageGB -lt $Current.StorageGB) { return 'Downsize' }
    return 'Resize'
}

function Set-UiInstallMode {
    param(
        [bool]$Enabled,
        [string]$Message = ""
    )

    if (Get-Variable -Name loadButton -Scope Script -ErrorAction SilentlyContinue) {
        $loadButton.Enabled = -not $Enabled
    }
    if (Get-Variable -Name filterButton -Scope Script -ErrorAction SilentlyContinue) {
        $filterButton.Enabled = -not $Enabled
    }
    if (Get-Variable -Name clearButton -Scope Script -ErrorAction SilentlyContinue) {
        $clearButton.Enabled = -not $Enabled
    }
    if (Get-Variable -Name exportButton -Scope Script -ErrorAction SilentlyContinue) {
        $exportButton.Enabled = -not $Enabled
    }
    if (Get-Variable -Name resizeButton -Scope Script -ErrorAction SilentlyContinue) {
        $resizeButton.Enabled = -not $Enabled
    }
    if (Get-Variable -Name searchTextBox -Scope Script -ErrorAction SilentlyContinue) {
        $searchTextBox.Enabled = -not $Enabled
    }

    if ((Get-Variable -Name statusLabel -Scope Script -ErrorAction SilentlyContinue) -and $Message) {
        $statusLabel.Text = $Message
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    }

    if (Get-Variable -Name progressBar -Scope Script -ErrorAction SilentlyContinue) {
        $progressBar.Visible = $Enabled
        $progressBar.Style = if ($Enabled) { [System.Windows.Forms.ProgressBarStyle]::Marquee } else { [System.Windows.Forms.ProgressBarStyle]::Blocks }
        if (-not $Enabled) {
            $progressBar.Value = 0
        }
    }
    if (Get-Variable -Name progressLabel -Scope Script -ErrorAction SilentlyContinue) {
        $progressLabel.Visible = $Enabled
        $progressLabel.Text = if ($Enabled) { "..." } else { "0%" }
    }

    if (Get-Variable -Name form -Scope Script -ErrorAction SilentlyContinue) {
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Ensure-RequiredGraphModules {
    $requiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Beta.DeviceManagement"
    )

    $missingModules = @($requiredModules | Where-Object {
        -not (Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue)
    })

    if ($script:forceModuleBootstrap) {
        $missingModules = @($requiredModules)
    }

    if ($missingModules.Count -eq 0) {
        return
    }

    $modulesText = ($missingModules -join ", ")
    $popupText = "Required Graph module(s): $modulesText`n`nInstall/repair now from PSGallery?"

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        $popupText,
        "Install Required Modules",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        throw "Required module(s) not installed: $modulesText"
    }

    try {
        Set-UiInstallMode -Enabled $true -Message "Installing required Graph modules..."

        # Avoid PowerShell prompts by preparing PSGallery/NuGet in current user scope.
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
        }

        try {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        }
        catch {
            # Non-fatal on locked-down environments; installation may still proceed.
            Write-DebugLog "Set-PSRepository warning: $($_.Exception.Message)"
        }

        foreach ($moduleName in $missingModules) {
            Set-UiInstallMode -Enabled $true -Message "Installing module: $moduleName ..."
            Install-Module -Name $moduleName -Repository PSGallery -Scope CurrentUser -AllowClobber -Force -Confirm:$false -ErrorAction Stop
        }

        foreach ($moduleName in $requiredModules) {
            Import-Module $moduleName -ErrorAction SilentlyContinue
        }

        Set-UiInstallMode -Enabled $false
        if (Get-Variable -Name statusLabel -Scope Script -ErrorAction SilentlyContinue) {
            $statusLabel.Text = "Required modules installed successfully."
            $statusLabel.ForeColor = [System.Drawing.Color]::Green
        }
    }
    catch {
        Set-UiInstallMode -Enabled $false
        throw "Module installation failed: $($_.Exception.Message)"
    }
}

function Ensure-GraphConnection {
    param(
        [switch]$RequireWrite
    )

    $requiredScopes = @(
        "CloudPC.Read.All",
        "Organization.Read.All"
    )

    if ($RequireWrite) {
        $requiredScopes += "CloudPC.ReadWrite.All"
    }

    Ensure-RequiredGraphModules

    $requiredCmdlets = @(
        "Connect-MgGraph",
        "Get-MgContext",
        "Get-MgBetaDeviceManagementVirtualEndpointCloudPc",
        "Invoke-MgGraphRequest"
    )

    foreach ($cmd in $requiredCmdlets) {
        if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
            throw "Required Graph cmdlet '$cmd' was not found. Install/update module 'Microsoft.Graph.Beta' and retry."
        }
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        return
    }

    $contextScopes = @($context.Scopes)
    $missingScopes = @($requiredScopes | Where-Object { $_ -notin $contextScopes })
    if ($missingScopes.Count -gt 0) {
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    }
}

function Invoke-WithGraphRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [int]$MaxAttempts = 4
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return (& $Action)
        }
        catch {
            $message = $_.Exception.Message
            $isRetryable = ($message -match '(?i)429|too many requests|throttl') -or ($message -match '(?i)\b5\d\d\b|internal server error|bad gateway|service unavailable|gateway timeout')
            if ($isRetryable -and $attempt -lt $MaxAttempts) {
                $backoffSeconds = [math]::Pow(2, $attempt - 1)
                if (Get-Variable -Name statusLabel -Scope Script -ErrorAction SilentlyContinue) {
                    $statusLabel.Text = "Throttled/server busy. Retry $attempt/$MaxAttempts in ${backoffSeconds}s..."
                    $statusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
                    $form.Refresh()
                }
                Start-Sleep -Seconds $backoffSeconds
                continue
            }
            throw
        }
    }
}

function Update-DataGrid {
    param($Data)

    if ($null -eq $Data) {
        $Data = @()
    } elseif ($Data -isnot [System.Array] -and $Data -isnot [System.Collections.IList]) {
        $Data = @($Data)
    }

    $dataGridView.SuspendLayout()
    $previousRowsMode = $dataGridView.AutoSizeRowsMode
    $dataGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
    foreach ($col in $dataGridView.Columns) {
        $col.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
    }
    $dataGridView.DataSource = $null

    try {
        if ($Data.Count -eq 0) {
            $countLabel.Text = "Records: 0"
            $script:filteredData = @()
            $resizeButton.Enabled = $false
            return
        }

        $dataTable = New-Object System.Data.DataTable
        $dataTable.BeginLoadData() | Out-Null

        $dataTable.Columns.Add("Select", [bool]) | Out-Null
        $dataTable.Columns.Add("Cloud PC Name", [string]) | Out-Null
        $dataTable.Columns.Add("User", [string]) | Out-Null
        $dataTable.Columns.Add("Status", [string]) | Out-Null
        $dataTable.Columns.Add("Edition", [string]) | Out-Null
        $dataTable.Columns.Add("Current License", [string]) | Out-Null
        $dataTable.Columns.Add("vCPU", [string]) | Out-Null
        $dataTable.Columns.Add("RAM (GB)", [string]) | Out-Null
        $dataTable.Columns.Add("Disk (GB)", [string]) | Out-Null
        $dataTable.Columns.Add("Cloud PC Id", [string]) | Out-Null
        $dataTable.Columns.Add("Service Plan Id", [string]) | Out-Null

        foreach ($item in $Data) {
            $row = $dataTable.NewRow()
            $row["Select"] = $false
            $row["Cloud PC Name"] = $item.ManagedDeviceName
            $row["User"] = $item.UserPrincipalName
            $row["Status"] = $item.Status
            $row["Edition"] = $item.Edition
            $row["Current License"] = $item.ServicePlanName
            $row["vCPU"] = if ($null -ne $item.VCpu) { [string]$item.VCpu } else { "" }
            $row["RAM (GB)"] = if ($null -ne $item.RamGB) { [string]$item.RamGB } else { "" }
            $row["Disk (GB)"] = if ($null -ne $item.StorageGB) { [string]$item.StorageGB } else { "" }
            $row["Cloud PC Id"] = $item.CloudPcId
            $row["Service Plan Id"] = $item.ServicePlanId
            $dataTable.Rows.Add($row)
        }

        $dataTable.EndLoadData() | Out-Null
        $dataGridView.DataSource = $dataTable

        if ($dataGridView.Columns.Contains("Cloud PC Id")) {
            $dataGridView.Columns["Cloud PC Id"].Visible = $false
        }
        if ($dataGridView.Columns.Contains("Service Plan Id")) {
            $dataGridView.Columns["Service Plan Id"].Visible = $false
        }

        if ($dataGridView.Columns.Contains("Select")) {
            $dataGridView.Columns["Select"].ReadOnly = $false
            $dataGridView.Columns["Select"].ThreeState = $false
            $dataGridView.Columns["Select"].DefaultCellStyle.NullValue = ""
            $dataGridView.Columns["Select"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
        }

        foreach ($col in $dataGridView.Columns) {
            if ($col.Name -ne "Select") {
                $col.ReadOnly = $true
            }
        }

        for ($i = 0; $i -lt $dataGridView.Columns.Count - 1; $i++) {
            $dataGridView.Columns[$i].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
        }
        $dataGridView.Columns[$dataGridView.Columns.Count - 1].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

        # Cloud PCs already resized this session cannot be selected again.
        foreach ($gridRow in $dataGridView.Rows) {
            $rowId = [string]$gridRow.Cells["Cloud PC Id"].Value
            $rowStatus = [string]$gridRow.Cells["Status"].Value

            # Only 'provisioned'/'provisionedWithWarnings' Cloud PCs (and not already
            # resized this session) can be selected for resize.
            $isResizable = ($rowStatus -match '(?i)^provisioned$' -or $rowStatus -match '(?i)^provisionedWithWarnings$')
            if (-not $isResizable -or ($rowId -and $script:resizedCloudPcIds.ContainsKey($rowId))) {
                $gridRow.Cells["Select"].Value = $false
                $gridRow.Cells["Select"].ReadOnly = $true
                $gridRow.Cells["Select"].Style.ForeColor = [System.Drawing.Color]::LightGray
                $gridRow.Cells["Select"].Style.SelectionForeColor = [System.Drawing.Color]::LightGray
            }

            # Highlight Cloud PCs currently being resized.
            if ($rowStatus -match '(?i)resizing') {
                $gridRow.DefaultCellStyle.BackColor = [System.Drawing.Color]::Yellow
            }
        }

        $countLabel.Text = "Records: $($Data.Count)"
        $script:filteredData = @($Data)

        $resizeButton.Enabled = $Data.Count -gt 0
    }
    finally {
        $dataGridView.AutoSizeRowsMode = $previousRowsMode
        $dataGridView.ResumeLayout()
    }
}

# Marks a Cloud PC's row as unselectable in the grid (used after a successful resize).
function Disable-CloudPcRowSelection {
    param([string]$CloudPcId)

    if ([string]::IsNullOrWhiteSpace($CloudPcId)) {
        return
    }

    foreach ($gridRow in $dataGridView.Rows) {
        if ([string]$gridRow.Cells["Cloud PC Id"].Value -eq $CloudPcId) {
            $gridRow.Cells["Select"].Value = $false
            $gridRow.Cells["Select"].ReadOnly = $true
            $gridRow.Cells["Select"].Style.ForeColor = [System.Drawing.Color]::LightGray
            $gridRow.Cells["Select"].Style.SelectionForeColor = [System.Drawing.Color]::LightGray
            break
        }
    }
}


# Returns the tenant's enterprise Cloud PC service plans (resize targets) as raw objects.
function Get-W365ServicePlans {
    $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/servicePlans"
    $plans = @()
    while ($uri) {
        $resp = Invoke-WithGraphRetry -Action { Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop }
        $value = Get-DictValue -Dict $resp -Names @('value')
        if ($value) {
            $plans += @($value)
        }
        $uri = [string](Get-DictValue -Dict $resp -Names @('@odata.nextLink'))
    }
    return $plans
}

# Builds a lookup of Enterprise and Business service plans keyed by
# "<edition>|vcpu|ram|storage" (e.g. "enterprise|2|8|128"). GPU plans are excluded.
function Get-ServicePlanMap {
    param($ServicePlans)

    $planMap = @{}
    foreach ($sp in $ServicePlans) {
        $id = [string](Get-DictValue -Dict $sp -Names @('id'))
        $name = [string](Get-DictValue -Dict $sp -Names @('displayName'))
        $type = [string](Get-DictValue -Dict $sp -Names @('type'))

        # Raw diagnostic dump of every service plan returned by the tenant.
        $rawVcpu = Get-DictValue -Dict $sp -Names @('vCpuCount')
        $rawRam = Get-DictValue -Dict $sp -Names @('ramInGB', 'ramInGb')
        $rawStorage = Get-DictValue -Dict $sp -Names @('storageInGB', 'storageInGb')
        $rawUserProfile = Get-DictValue -Dict $sp -Names @('userProfileInGB', 'userProfileInGb')
        Write-DebugLog "SERVICE PLAN raw: id=$id name='$name' type='$type' vCpu=$rawVcpu ram=$rawRam storage=$rawStorage userProfile=$rawUserProfile"

        # Exclude plans handled by a manual/internal target (e.g. CloudPC_Lite) so they
        # do not collide with a standard plan of the same vCPU/RAM/storage.
        if ($id -and ($script:manualLicenseTargets | Where-Object { $_.ServicePlanId -eq $id })) {
            Write-DebugLog "SERVICE PLAN skipped (handled as manual target): id=$id name='$name'"
            continue
        }

        # Skip GPU plans (GPU resize unsupported) and Frontline/Flex/shared plans.
        # NOTE: Frontline service plans report type='enterprise', so classify by the
        # display name (which explicitly says Enterprise/Business/Frontline) first.
        if ($name -match '(?i)gpu') {
            Write-DebugLog "SERVICE PLAN skipped (gpu): '$name'"
            continue
        }
        if ($name -match '(?i)shared' -or (Test-IsFrontlineOrFlex $name)) {
            Write-DebugLog "SERVICE PLAN skipped (frontline/flex/shared): '$name'"
            continue
        }

        $edition = $null
        if ($name -match '(?i)business') { $edition = 'business' }
        elseif ($name -match '(?i)enterprise') { $edition = 'enterprise' }
        elseif ($type -match '(?i)business') { $edition = 'business' }
        elseif ($type -match '(?i)enterprise') { $edition = 'enterprise' }
        else {
            Write-DebugLog "SERVICE PLAN skipped (edition not resolved): name='$name' type='$type'"
            continue
        }

        $vcpu = Get-DictValue -Dict $sp -Names @('vCpuCount')
        $ram = Get-DictValue -Dict $sp -Names @('ramInGB', 'ramInGb')
        $storage = Get-DictValue -Dict $sp -Names @('storageInGB', 'storageInGb')

        if (-not $vcpu -or -not $ram -or -not $storage) {
            $parsed = Get-W365ConfigFromText -Text $name
            if (-not $vcpu) { $vcpu = $parsed.VCpu }
            if (-not $ram) { $ram = $parsed.RamGB }
            if (-not $storage) { $storage = $parsed.StorageGB }
        }

        if ($vcpu -and $ram -and $storage -and $id) {
            $key = "$edition|$([int]$vcpu)|$([int]$ram)|$([int]$storage)"
            if (-not $planMap.ContainsKey($key)) {
                $planMap[$key] = [pscustomobject]@{
                    Id        = $id
                    Name      = $name
                    Edition   = $edition
                    VCpu      = [int]$vcpu
                    RamGB     = [int]$ram
                    StorageGB = [int]$storage
                }
                Write-DebugLog "SERVICE PLAN mapped: key=$key id=$id name='$name' type='$type'"
            }
        }
    }
    return $planMap
}

# Returns Windows 365 Enterprise and Business licenses (subscribed SKUs) with seat counts
# and the resize service plan id. GPU licenses are excluded (GPU resize is not supported).
function Get-W365Licenses {
    param($ServicePlans)

    $planMap = Get-ServicePlanMap -ServicePlans $ServicePlans
    Write-DebugLog "LICENSES: service plan keys = $((@($planMap.Keys)) -join ', ')"

    $resp = Invoke-WithGraphRetry -Action { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -ErrorAction Stop }
    $skus = @()
    $value = Get-DictValue -Dict $resp -Names @('value')
    if ($value) {
        $skus = @($value)
    }
    Write-DebugLog "LICENSES: subscribedSkus returned = $($skus.Count)"

    $licenses = @()
    foreach ($sku in $skus) {
        $partNumber = [string](Get-DictValue -Dict $sku -Names @('skuPartNumber'))

        # Collect this SKU's child service plan names (config lives here for leveling SKUs).
        $childPlans = @(Get-DictValue -Dict $sku -Names @('servicePlans'))
        $childNames = @()
        foreach ($cp in $childPlans) {
            $cpName = [string](Get-DictValue -Dict $cp -Names @('servicePlanName'))
            if (-not [string]::IsNullOrWhiteSpace($cpName)) {
                $childNames += $cpName
            }
        }

        # Identify the Enterprise or Business Cloud PC provisioning plan and its config.
        # The config-bearing token can be the SKU part number (e.g. CPC_E_2C_8GB_128GB,
        # CPC_B_2C_8GB_128GB or Windows_365_Enterprise_...) or a child plan
        # (e.g. CPC_LVL_3 -> CPC_E_4C_16GB_256GB).
        # Exclude Frontline/Flex/shared (CPC_S_*, Windows 365 Flex), cross-region DR
        # add-ons, GPU plans, and non-Cloud-PC plans.
        $planToken = $null
        $edition = $null
        $config = [pscustomobject]@{ VCpu = $null; RamGB = $null; StorageGB = $null }
        foreach ($candidate in (@($partNumber) + $childNames)) {
            $token = ([string]$candidate).Trim()
            if ([string]::IsNullOrWhiteSpace($token)) { continue }
            if ($token -match '(?i)crossregion|disasterrecovery') { continue }
            if ($token -match '(?i)gpu') { continue }
            # Frontline / Flex (incl. future "Windows 365 Flex") / shared plans.
            if ((Test-IsFrontlineOrFlex $token) -or ($token -match '(?i)cpc_s_')) { continue }

            $tokenEdition = $null
            if (($token -match '(?i)cpc_e') -or ($token -match '(?i)windows[_ ]?365[_ ]?enterprise')) {
                $tokenEdition = 'Enterprise'
            }
            elseif (($token -match '(?i)cpc_b') -or ($token -match '(?i)windows[_ ]?365[_ ]?business')) {
                $tokenEdition = 'Business'
            }
            if (-not $tokenEdition) { continue }

            $cfg = Get-W365ConfigFromText -Text $token
            if ($cfg.VCpu -and $cfg.RamGB -and $cfg.StorageGB) {
                $planToken = $token
                $edition = $tokenEdition
                $config = $cfg
                break
            }
        }

        if (-not $planToken) {
            # Manual/internal license targets (e.g. CloudPC_Lite add-on) whose SKU has no
            # config-bearing child plan. Match against the part number or any child name.
            $manual = $null
            foreach ($mt in $script:manualLicenseTargets) {
                $matched = ($partNumber -match "(?i)$($mt.SkuMatch)")
                if (-not $matched) {
                    foreach ($cn in $childNames) {
                        if ($cn -match "(?i)$($mt.SkuMatch)") { $matched = $true; break }
                    }
                }
                if ($matched) { $manual = $mt; break }
            }

            if ($manual) {
                $prepaid = Get-DictValue -Dict $sku -Names @('prepaidUnits')
                $enabled = [int](Get-DictValue -Dict $prepaid -Names @('enabled'))
                $consumed = [int](Get-DictValue -Dict $sku -Names @('consumedUnits'))
                $available = $enabled - $consumed
                if ($available -lt 0) { $available = 0 }

                Write-DebugLog "LICENSE SKU '$partNumber' -> manual target '$($manual.PlanName)' [$($manual.Edition)]: enabled=$enabled consumed=$consumed available=$available servicePlanId=$($manual.ServicePlanId)"

                $licenses += [pscustomobject]@{
                    SkuId           = [string](Get-DictValue -Dict $sku -Names @('skuId'))
                    SkuPartNumber   = $partNumber
                    DisplayName     = $manual.DisplayName
                    Edition         = $manual.Edition
                    ServicePlanId   = $manual.ServicePlanId
                    ServicePlanName = $manual.PlanName
                    VCpu            = [int]$manual.VCpu
                    RamGB           = [int]$manual.RamGB
                    StorageGB       = [int]$manual.StorageGB
                    Enabled         = $enabled
                    Consumed        = $consumed
                    Available       = $available
                }
                continue
            }

            Write-DebugLog "LICENSE SKU skipped (no Enterprise/Business Cloud PC plan): '$partNumber' children=[$($childNames -join '; ')]"
            continue
        }

        $prepaid = Get-DictValue -Dict $sku -Names @('prepaidUnits')
        $enabled = [int](Get-DictValue -Dict $prepaid -Names @('enabled'))
        $consumed = [int](Get-DictValue -Dict $sku -Names @('consumedUnits'))
        $available = $enabled - $consumed
        if ($available -lt 0) { $available = 0 }

        $vcpu = [int]$config.VCpu
        $ram = [int]$config.RamGB
        $storage = [int]$config.StorageGB

        $servicePlanId = $null
        $planName = $planToken
        $key = "$($edition.ToLower())|$vcpu|$ram|$storage"
        if ($planMap.ContainsKey($key)) {
            $servicePlanId = $planMap[$key].Id
            $planName = $planMap[$key].Name
        }

        $displayName = "Windows 365 $edition $vcpu vCPU / $ram GB / $storage GB"

        Write-DebugLog "LICENSE SKU '$partNumber' -> plan '$planToken' [$edition]: enabled=$enabled consumed=$consumed available=$available vCPU=$vcpu ram=$ram storage=$storage servicePlanId=$servicePlanId"

        $licenses += [pscustomobject]@{
            SkuId           = [string](Get-DictValue -Dict $sku -Names @('skuId'))
            SkuPartNumber   = $partNumber
            DisplayName     = $displayName
            Edition         = $edition
            ServicePlanId   = $servicePlanId
            ServicePlanName = $planName
            VCpu            = $vcpu
            RamGB           = $ram
            StorageGB       = $storage
            Enabled         = $enabled
            Consumed        = $consumed
            Available       = $available
        }
    }

    return @($licenses | Sort-Object Edition, VCpu, RamGB, StorageGB)
}

# Modal popup: shows available Windows 365 Enterprise licenses as single-select radio
# buttons. Returns the chosen license object, or $null if cancelled.
function Show-ResizeLicenseDialog {
    param($Current, $Licenses)

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Select target Windows 365 $($Current.Edition) license"
    $dialog.Size = New-Object System.Drawing.Size(900, 560)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.BackColor = [System.Drawing.Color]::White

    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Location = New-Object System.Drawing.Point(15, 12)
    $headerLabel.Size = New-Object System.Drawing.Size(860, 40)
    $headerLabel.Text = "Cloud PC: $($Current.DisplayName)`r`nCurrent license: $($Current.ServicePlanName)  (vCPU $($Current.VCpu) / RAM $($Current.RamGB) GB / Disk $($Current.StorageGB) GB)"
    $headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $dialog.Controls.Add($headerLabel)

    $hintLabel = New-Object System.Windows.Forms.Label
    $hintLabel.Location = New-Object System.Drawing.Point(15, 56)
    $hintLabel.Size = New-Object System.Drawing.Size(860, 20)
    $hintLabel.Text = "Downsizing to a license with a smaller disk than the current one is not allowed."
    $hintLabel.ForeColor = [System.Drawing.Color]::DimGray
    $dialog.Controls.Add($hintLabel)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point(15, 82)
    $panel.Size = New-Object System.Drawing.Size(860, 390)
    $panel.AutoScroll = $true
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $dialog.Controls.Add($panel)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(680, 482)
    $okButton.Size = New-Object System.Drawing.Size(90, 30)
    $okButton.Text = "Continue"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Enabled = $false
    $dialog.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(780, 482)
    $cancelButton.Size = New-Object System.Drawing.Size(90, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($cancelButton)

    $dialog.AcceptButton = $okButton
    $dialog.CancelButton = $cancelButton

    $y = 10
    $radioButtons = @()
    foreach ($license in $Licenses) {
        $hasPlanId = -not [string]::IsNullOrWhiteSpace([string]$license.ServicePlanId)
        $hasConfig = ($null -ne $license.VCpu) -and ($null -ne $license.RamGB) -and ($null -ne $license.StorageGB)
        $isCurrent = $false
        if ($hasPlanId -and -not [string]::IsNullOrWhiteSpace([string]$Current.ServicePlanId)) {
            # When both have a service plan id, match strictly by id so distinct plans that
            # share the same vCPU/RAM/storage (e.g. CloudPC_Lite vs a standard plan) are
            # not treated as the current plan.
            $isCurrent = ([string]$license.ServicePlanId -eq [string]$Current.ServicePlanId)
        }
        elseif ($hasConfig -and ($null -ne $Current.VCpu)) {
            $isCurrent = ($license.VCpu -eq $Current.VCpu) -and ($license.RamGB -eq $Current.RamGB) -and ($license.StorageGB -eq $Current.StorageGB)
        }

        $diskOk = $true
        if ($hasConfig -and ($null -ne $Current.StorageGB)) {
            $diskOk = ($license.StorageGB -ge $Current.StorageGB)
        }

        $enabled = $hasPlanId -and $hasConfig -and (-not $isCurrent) -and ($license.Available -gt 0) -and $diskOk

        $reason = ""
        $opTag = ""
        if ($isCurrent) {
            $reason = "  (current license)"
        }
        elseif (-not $hasPlanId) {
            $reason = "  (not available as a resize target)"
        }
        elseif ($license.Available -le 0) {
            $reason = "  (no seats available)"
        }
        elseif (-not $diskOk) {
            $reason = "  (disk smaller than current - downsize not allowed)"
        }
        elseif ($hasConfig) {
            $op = Get-ResizeOperationType -Current $Current -Target $license
            $opTag = if ($op -eq 'Downsize') { "  [Downsize]" } else { "  [Upgrade]" }
        }

        $radio = New-Object System.Windows.Forms.RadioButton
        $radio.Location = New-Object System.Drawing.Point(12, $y)
        $radio.Size = New-Object System.Drawing.Size(810, 34)
        $radio.Text = "$($license.DisplayName)   |   Seats: $($license.Available) of $($license.Enabled) available$opTag$reason"
        $radio.Enabled = $enabled
        $radio.Tag = $license
        if (-not $enabled) {
            $radio.ForeColor = [System.Drawing.Color]::Gray
        }
        $radio.Add_CheckedChanged({
            if ($this.Checked) {
                $okButton.Enabled = $true
            }
        }.GetNewClosure())
        $panel.Controls.Add($radio)
        $radioButtons += $radio
        $y += 38
    }

    if ($radioButtons.Count -eq 0) {
        $emptyLabel = New-Object System.Windows.Forms.Label
        $emptyLabel.Location = New-Object System.Drawing.Point(12, 12)
        $emptyLabel.Size = New-Object System.Drawing.Size(600, 40)
        $emptyLabel.Text = "No Windows 365 Enterprise licenses were found in the tenant."
        $panel.Controls.Add($emptyLabel)
    }

    $result = $dialog.ShowDialog()
    $selected = $null
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($radio in $radioButtons) {
            if ($radio.Checked) {
                $selected = $radio.Tag
                break
            }
        }
    }
    $dialog.Dispose()
    return $selected
}

function Invoke-ResizeSelectedCloudPc {
    if ($dataGridView.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No data loaded.",
            "Info",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    $selectedRows = @($dataGridView.Rows | Where-Object {
        $v = $_.Cells["Select"].Value
        ($null -ne $v) -and ($v -isnot [System.DBNull]) -and ([bool]$v)
    })

    if ($selectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Select one Cloud PC to resize.",
            "Info",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    if ($selectedRows.Count -gt 1) {
        [System.Windows.Forms.MessageBox]::Show(
            "Only one Cloud PC can be resized at a time. Please select a single Cloud PC.",
            "Info",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    $gridRow = $selectedRows[0]

    # Only Cloud PCs that are fully provisioned can be resized (provisionedWithWarnings is allowed).
    $currentStatus = [string]$gridRow.Cells["Status"].Value
    if ($currentStatus -notmatch '(?i)^provisioned$' -and $currentStatus -notmatch '(?i)^provisionedWithWarnings$') {
        [System.Windows.Forms.MessageBox]::Show(
            "This Cloud PC cannot be resized because its status is '$currentStatus'.`n`nOnly Cloud PCs in a 'provisioned' (or 'provisionedWithWarnings') state can be resized.",
            "Resize not allowed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    $currentVCpu = $null
    $currentRam = $null
    $currentStorage = $null
    if (-not [string]::IsNullOrWhiteSpace([string]$gridRow.Cells["vCPU"].Value)) { $currentVCpu = [int]$gridRow.Cells["vCPU"].Value }
    if (-not [string]::IsNullOrWhiteSpace([string]$gridRow.Cells["RAM (GB)"].Value)) { $currentRam = [int]$gridRow.Cells["RAM (GB)"].Value }
    if (-not [string]::IsNullOrWhiteSpace([string]$gridRow.Cells["Disk (GB)"].Value)) { $currentStorage = [int]$gridRow.Cells["Disk (GB)"].Value }

    $current = [pscustomobject]@{
        CloudPcId       = [string]$gridRow.Cells["Cloud PC Id"].Value
        DisplayName     = [string]$gridRow.Cells["Cloud PC Name"].Value
        Edition         = [string]$gridRow.Cells["Edition"].Value
        ServicePlanId   = [string]$gridRow.Cells["Service Plan Id"].Value
        ServicePlanName = [string]$gridRow.Cells["Current License"].Value
        VCpu            = $currentVCpu
        RamGB           = $currentRam
        StorageGB       = $currentStorage
    }

    try {
        $statusLabel.Text = "Loading Windows 365 licenses..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()

        Ensure-GraphConnection

        if (-not $script:servicePlans -or $script:servicePlans.Count -eq 0) {
            $script:servicePlans = @(Get-W365ServicePlans)
        }

        $allLicenses = @(Get-W365Licenses -ServicePlans $script:servicePlans)

        # Cross-edition resize is not allowed: only Enterprise->Enterprise and Business->Business.
        $licenses = @($allLicenses | Where-Object { $_.Edition -eq $current.Edition })

        $statusLabel.Text = "Ready."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue

        if ($licenses.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No Windows 365 $($current.Edition) licenses were found in the tenant.",
                "Info",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }

        $target = Show-ResizeLicenseDialog -Current $current -Licenses $licenses
        if ($null -eq $target) {
            return
        }

        $operation = Get-ResizeOperationType -Current $current -Target $target

        $confirmText = "Cloud PC: $($current.DisplayName)`n`n" +
            "Current license: $($current.ServicePlanName)`n" +
            "   vCPU $($current.VCpu) / RAM $($current.RamGB) GB / Disk $($current.StorageGB) GB`n`n" +
            "Target license: $($target.DisplayName)`n" +
            "   vCPU $($target.VCpu) / RAM $($target.RamGB) GB / Disk $($target.StorageGB) GB`n`n" +
            "Action: $operation`n`n" +
            "Do you want to proceed?"

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $confirmText,
            "Confirm $operation",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }

        if ([string]::IsNullOrWhiteSpace($target.ServicePlanId)) {
            throw "The selected target license has no resolvable service plan id and cannot be used as a resize target."
        }

        if (($target.ServicePlanName -match '(?i)shared') -or (Test-IsFrontlineOrFlex $target.ServicePlanName)) {
            throw "Resolved target service plan '$($target.ServicePlanName)' is a Frontline/Flex plan, which is not a valid resize target."
        }

        Ensure-GraphConnection -RequireWrite

        $statusLabel.Text = "Submitting $operation for $($current.DisplayName)..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()

        $resizeUri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs/$($current.CloudPcId)/resize"
        $body = @{ targetServicePlanId = $target.ServicePlanId }
        $bodyJson = $body | ConvertTo-Json -Depth 5 -Compress

        Write-DebugLog "RESIZE REQUEST: POST $resizeUri body=$bodyJson currentServicePlanId=$($current.ServicePlanId) targetServicePlanId=$($target.ServicePlanId) currentPlan='$($current.ServicePlanName)' targetPlan='$($target.ServicePlanName)' operation=$operation"

        try {
            $httpResp = Invoke-WithGraphRetry -Action {
                Invoke-MgGraphRequest -Method POST -Uri $resizeUri -Body $bodyJson -ContentType 'application/json' -OutputType HttpResponseMessage -ErrorAction Stop
            }

            $statusCode = [int]$httpResp.StatusCode
            $respContent = ""
            try { $respContent = $httpResp.Content.ReadAsStringAsync().Result } catch { }
            Write-DebugLog "RESIZE RESPONSE: status=$statusCode content=$respContent"

            if ($statusCode -lt 200 -or $statusCode -ge 300) {
                throw "Graph resize returned HTTP $statusCode. $respContent"
            }
        }
        catch {
            Write-DebugLog "RESIZE CALL FAILED for '$($current.DisplayName)' ($($current.CloudPcId)): $($_.Exception.Message)"
            $statusLabel.Text = "Resize failed for $($current.DisplayName)."
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show(
                ("Failed" + "`n`n" +
                    "Cloud PC: $($current.DisplayName)`n" +
                    "Action: $operation`n" +
                    "Target license: $($target.DisplayName)`n`n" +
                    "Error: $($_.Exception.Message)`n`n" +
                    "Please check the Intune portal (Devices > Windows 365 > All Cloud PCs) for further details."),
                "Resize failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        # Success: prevent re-selecting this Cloud PC for the rest of the session and grey its row.
        $script:resizedCloudPcIds[$current.CloudPcId] = $true
        Disable-CloudPcRowSelection -CloudPcId $current.CloudPcId

        $statusLabel.Text = "Resize submitted for $($current.DisplayName)."
        $statusLabel.ForeColor = [System.Drawing.Color]::Green

        [System.Windows.Forms.MessageBox]::Show(
            ("Success" + "`n`n" +
                "Cloud PC: $($current.DisplayName)`n" +
                "Action: $operation`n" +
                "Target license: $($target.DisplayName)`n`n" +
                "Please check the Intune portal (Devices > Windows 365 > All Cloud PCs) for further details."),
            "Resize submitted",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        # Refresh inventory to reflect the new state (resized Cloud PC stays unselectable).
        Get-CloudPcInventory
    }
    catch {
        Write-DebugLog "RESIZE WORKFLOW ERROR: $($_.Exception.Message)"
        $statusLabel.Text = "Resize failed: $($_.Exception.Message)"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show(
            "Resize failed:`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Invoke-Filters {
    if ($script:allData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please load data first.",
            "Info",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    $filtered = @($script:allData)

    if (-not [string]::IsNullOrWhiteSpace($searchTextBox.Text)) {
        $term = $searchTextBox.Text.Trim()
        $filtered = @($filtered | Where-Object {
            ($_.DisplayName -like "*$term*") -or ($_.UserPrincipalName -like "*$term*")
        })
    }

    Update-DataGrid -Data $filtered
    $statusLabel.Text = "Filter applied. Showing $($filtered.Count) of $($script:allData.Count) Cloud PCs."
    $statusLabel.ForeColor = [System.Drawing.Color]::Green
}

function Clear-Filters {
    $searchTextBox.Clear()

    if ($script:allData.Count -gt 0) {
        Update-DataGrid -Data $script:allData
        $statusLabel.Text = "Filters cleared. Showing all $($script:allData.Count) Cloud PCs."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    }
}

function Export-ToCsv {
    if ($script:filteredData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No data to export.",
            "Info",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "W365_CloudPC_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:filteredData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show(
                "Data exported successfully to:`n$($saveDialog.FileName)",
                "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Error exporting data:`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
}

function Get-CloudPcInventory {
    try {
        $progressBar.Value = 0
        $progressBar.Visible = $true
        $progressLabel.Visible = $true
        $progressLabel.Text = "0%"

        $statusLabel.Text = "Checking Graph SDK connection..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()

        Ensure-GraphConnection

        $progressBar.Value = 20
        $progressLabel.Text = "20%"
        $statusLabel.Text = "Loading Cloud PC service plans..."
        $form.Refresh()

        $script:servicePlans = @(Get-W365ServicePlans)
        $planById = @{}
        foreach ($sp in $script:servicePlans) {
            $id = [string](Get-DictValue -Dict $sp -Names @('id'))
            if ([string]::IsNullOrWhiteSpace($id)) { continue }

            $name = [string](Get-DictValue -Dict $sp -Names @('displayName'))
            $vcpu = Get-DictValue -Dict $sp -Names @('vCpuCount')
            $ram = Get-DictValue -Dict $sp -Names @('ramInGB', 'ramInGb')
            $storage = Get-DictValue -Dict $sp -Names @('storageInGB', 'storageInGb')
            if (-not $vcpu -or -not $ram -or -not $storage) {
                $parsed = Get-W365ConfigFromText -Text $name
                if (-not $vcpu) { $vcpu = $parsed.VCpu }
                if (-not $ram) { $ram = $parsed.RamGB }
                if (-not $storage) { $storage = $parsed.StorageGB }
            }

            $planById[$id] = [pscustomobject]@{
                Name      = $name
                VCpu      = if ($vcpu) { [int]$vcpu } else { $null }
                RamGB     = if ($ram) { [int]$ram } else { $null }
                StorageGB = if ($storage) { [int]$storage } else { $null }
            }
        }

        $progressBar.Value = 45
        $progressLabel.Text = "45%"
        $statusLabel.Text = "Loading Cloud PCs..."
        $form.Refresh()

        $cloudPcs = Get-MgBetaDeviceManagementVirtualEndpointCloudPc -All -ErrorAction Stop -Property @(
            "id",
            "displayName",
            "managedDeviceName",
            "userPrincipalName",
            "status",
            "servicePlanId",
            "servicePlanName",
            "servicePlanType",
            "provisioningType"
        )

        $progressBar.Value = 75
        $progressLabel.Text = "75%"
        $statusLabel.Text = "Filtering Windows 365 Enterprise/Business Cloud PCs..."
        $form.Refresh()

        $rows = @()
        foreach ($cpc in $cloudPcs) {
            $edition = Get-CloudPcEdition -CloudPc $cpc
            if (-not $edition) {
                continue
            }

            $servicePlanId = [string](Get-CloudPcPropertyValue -CloudPc $cpc -Names @('servicePlanId', 'ServicePlanId'))
            $servicePlanName = [string](Get-CloudPcPropertyValue -CloudPc $cpc -Names @('servicePlanName', 'ServicePlanName'))

            $vcpu = $null
            $ram = $null
            $storage = $null
            if ($servicePlanId -and $planById.ContainsKey($servicePlanId)) {
                $vcpu = $planById[$servicePlanId].VCpu
                $ram = $planById[$servicePlanId].RamGB
                $storage = $planById[$servicePlanId].StorageGB
                if ([string]::IsNullOrWhiteSpace($servicePlanName)) {
                    $servicePlanName = $planById[$servicePlanId].Name
                }
            }
            if (-not $vcpu -or -not $ram -or -not $storage) {
                $parsed = Get-W365ConfigFromText -Text $servicePlanName
                if (-not $vcpu) { $vcpu = $parsed.VCpu }
                if (-not $ram) { $ram = $parsed.RamGB }
                if (-not $storage) { $storage = $parsed.StorageGB }
            }

            $rows += [pscustomobject]@{
                CloudPcId         = [string](Get-CloudPcPropertyValue -CloudPc $cpc -Names @('id', 'Id'))
                DisplayName       = [string](Get-CloudPcPropertyValue -CloudPc $cpc -Names @('displayName', 'DisplayName'))
                ManagedDeviceName = [string](Get-CloudPcPropertyValue -CloudPc $cpc -Names @('managedDeviceName', 'ManagedDeviceName'))
                UserPrincipalName = [string](Get-CloudPcPropertyValue -CloudPc $cpc -Names @('userPrincipalName', 'UserPrincipalName'))
                Status            = [string](Get-CloudPcPropertyValue -CloudPc $cpc -Names @('status', 'Status'))
                Edition           = $edition
                ServicePlanId     = $servicePlanId
                ServicePlanName   = $servicePlanName
                VCpu              = $vcpu
                RamGB             = $ram
                StorageGB         = $storage
            }
        }

        $script:allData = @($rows | Sort-Object DisplayName)

        $progressBar.Value = 90
        $progressLabel.Text = "90%"
        $statusLabel.Text = "Preparing UI..."
        $form.Refresh()

        Update-DataGrid -Data $script:allData

        $progressBar.Value = 100
        $progressLabel.Text = "100%"
        $form.Refresh()

        Start-Sleep -Milliseconds 300

        $statusLabel.Text = "Loaded successfully. Windows 365 Enterprise/Business Cloud PCs: $($script:allData.Count)."
        $statusLabel.ForeColor = [System.Drawing.Color]::Green

        $exportButton.Enabled = $script:allData.Count -gt 0
        $progressBar.Visible = $false
        $progressLabel.Visible = $false
    }
    catch {
        Write-DebugLog "ERROR: $($_.Exception.Message)"
        Write-DebugLog "TYPE: $($_.Exception.GetType().FullName)"
        Write-DebugLog "STACK: $($_.ScriptStackTrace)"
        Write-DebugLog "DETAILS: $($_ | Out-String)"
        $progressBar.Visible = $false
        $progressLabel.Visible = $false
        $statusLabel.Text = "Error loading data: $($_.Exception.Message). Check log: $script:logPath"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show(
            "Error loading Cloud PCs:`n`n$($_.Exception.Message)`n`nLog: $script:logPath",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows 365 Enterprise Cloud PC Resize Tool"
$form.Size = New-Object System.Drawing.Size(1400, 800)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.MinimumSize = New-Object System.Drawing.Size(1200, 600)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 10)
$statusLabel.Size = New-Object System.Drawing.Size(700, 20)
$statusLabel.Text = "Click 'Load Cloud PCs' to query Windows 365 Enterprise Cloud PCs..."
$statusLabel.ForeColor = [System.Drawing.Color]::Blue
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($statusLabel)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(720, 10)
$progressBar.Size = New-Object System.Drawing.Size(500, 20)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Visible = $false
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($progressBar)

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(1210, 10)
$progressLabel.Size = New-Object System.Drawing.Size(60, 20)
$progressLabel.Text = "0%"
$progressLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$progressLabel.ForeColor = [System.Drawing.Color]::Blue
$progressLabel.Visible = $false
$progressLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($progressLabel)

# Buttons
$loadButton = New-Object System.Windows.Forms.Button
$loadButton.Location = New-Object System.Drawing.Point(10, 40)
$loadButton.Size = New-Object System.Drawing.Size(170, 30)
$loadButton.Text = "Load Cloud PCs"
$loadButton.BackColor = [System.Drawing.Color]::DodgerBlue
$loadButton.ForeColor = [System.Drawing.Color]::White
$loadButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($loadButton)

$filterButton = New-Object System.Windows.Forms.Button
$filterButton.Location = New-Object System.Drawing.Point(190, 40)
$filterButton.Size = New-Object System.Drawing.Size(105, 30)
$filterButton.Text = "Apply Filter"
$filterButton.BackColor = [System.Drawing.Color]::Green
$filterButton.ForeColor = [System.Drawing.Color]::White
$filterButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($filterButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Location = New-Object System.Drawing.Point(300, 40)
$clearButton.Size = New-Object System.Drawing.Size(100, 30)
$clearButton.Text = "Clear Filter"
$clearButton.BackColor = [System.Drawing.Color]::Gray
$clearButton.ForeColor = [System.Drawing.Color]::White
$clearButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($clearButton)

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(410, 40)
$exportButton.Size = New-Object System.Drawing.Size(110, 30)
$exportButton.Text = "Export CSV"
$exportButton.BackColor = [System.Drawing.Color]::Orange
$exportButton.ForeColor = [System.Drawing.Color]::White
$exportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$exportButton.Enabled = $false
$exportButton.Visible = $false
$form.Controls.Add($exportButton)

$resizeButton = New-Object System.Windows.Forms.Button
$resizeButton.Location = New-Object System.Drawing.Point(410, 40)
$resizeButton.Size = New-Object System.Drawing.Size(200, 30)
$resizeButton.Text = "Resize Selected Cloud PC"
$resizeButton.BackColor = [System.Drawing.Color]::MediumSeaGreen
$resizeButton.ForeColor = [System.Drawing.Color]::White
$resizeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$resizeButton.Enabled = $false
$form.Controls.Add($resizeButton)

# Search controls
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Location = New-Object System.Drawing.Point(10, 82)
$searchLabel.Size = New-Object System.Drawing.Size(150, 20)
$searchLabel.Text = "Name/User contains:"
$form.Controls.Add($searchLabel)

$searchTextBox = New-Object System.Windows.Forms.TextBox
$searchTextBox.Location = New-Object System.Drawing.Point(165, 80)
$searchTextBox.Size = New-Object System.Drawing.Size(520, 25)
$searchTextBox.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($searchTextBox)

# Result count
$countLabel = New-Object System.Windows.Forms.Label
$countLabel.Location = New-Object System.Drawing.Point(10, 115)
$countLabel.Size = New-Object System.Drawing.Size(1360, 20)
$countLabel.Text = "Records: 0"
$countLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($countLabel)

# Data grid
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 140)
$dataGridView.Size = New-Object System.Drawing.Size(1360, 620)
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $false
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::DodgerBlue
$dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$dataGridView.ColumnHeadersHeight = 40
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($dataGridView)

# Event handlers
$loadButton.Add_Click({ Get-CloudPcInventory })
$filterButton.Add_Click({ Invoke-Filters })
$clearButton.Add_Click({ Clear-Filters })
$exportButton.Add_Click({ Export-ToCsv })
$resizeButton.Add_Click({ Invoke-ResizeSelectedCloudPc })

$searchTextBox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Invoke-Filters
    }
})

$dataGridView.Add_CurrentCellDirtyStateChanged({
    if ($dataGridView.IsCurrentCellDirty) {
        $dataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

$dataGridView.Add_CellBeginEdit({
    param($sender, $e)
    if ($e.ColumnIndex -lt 0 -or $e.RowIndex -lt 0) { return }

    $columnName = $dataGridView.Columns[$e.ColumnIndex].Name
    if ($columnName -ne "Select") {
        $e.Cancel = $true
    }
})

# Enforce single-selection: checking one row clears any other selected row.
$dataGridView.Add_CellValueChanged({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
    if ($dataGridView.Columns[$e.ColumnIndex].Name -ne "Select") { return }
    if ($script:suppressSelectSync) { return }

    $value = $dataGridView.Rows[$e.RowIndex].Cells["Select"].Value
    if (($null -ne $value) -and ($value -isnot [System.DBNull]) -and ([bool]$value)) {
        $script:suppressSelectSync = $true
        try {
            foreach ($otherRow in $dataGridView.Rows) {
                if ($otherRow.Index -ne $e.RowIndex) {
                    $otherValue = $otherRow.Cells["Select"].Value
                    if (($null -ne $otherValue) -and ($otherValue -isnot [System.DBNull]) -and ([bool]$otherValue)) {
                        $otherRow.Cells["Select"].Value = $false
                    }
                }
            }
        }
        finally {
            $script:suppressSelectSync = $false
        }
    }
})

# Show form
[void]$form.ShowDialog()
