<#
.SYNOPSIS
    Exports all Conditional Access policies from a tenant using Microsoft Graph.

.DESCRIPTION
    - Uses delegated authentication (Connect-MgGraph) with Policy.Read.All plus additional read scopes.
    - Installs missing modules for the current user.
    - Reads /v1.0/identity/conditionalAccess/policies, including pagination.
    - Saves one JSON file per policy, AllPolicies.json, PolicyOverview.csv, and PolicyOverview.xlsx.
    - The overview includes the documented policy, condition, grant, and session-control fields from Graph.

.VERSION 1.0.0

.PARAMETER OutputFolder
    Output folder. Default: .\CA-Export_<yyyy-MM-dd>

.PARAMETER TenantId
    Optional tenant ID or domain to authenticate against a specific tenant.
#>
[CmdletBinding()]
param(
    [string]$OutputFolder = ".\CA-Export_$(Get-Date -Format 'yyyy-MM-dd')",
    [string]$TenantId
)

$ErrorActionPreference = 'Stop'
$script:NameCache = @{}
$script:Locations = @{}
$script:Roles = @{}
$ResolveNames = $true
$exitCode = 0

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

function Write-GraphError {
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [string]$Context
    )

    $message = $ErrorRecord.Exception.Message
    Write-Host "`nERROR in: $Context" -ForegroundColor Red

    if ($message -match 'Forbidden|403|Authorization_RequestDenied|Insufficient privileges') {
        Write-Host 'Access denied. Check the required Graph permissions or admin consent.' -ForegroundColor Yellow
    }
    elseif ($message -match 'AADSTS65001|consent') {
        Write-Host 'Consent is missing for the requested Microsoft Graph permissions.' -ForegroundColor Yellow
    }
    elseif ($message -match '401|Unauthorized|token') {
        Write-Host 'Authentication failed or the token expired.' -ForegroundColor Yellow
    }

    Write-Host "Details: $message" -ForegroundColor DarkGray
}

function Get-GraphPaged {
    param([Parameter(Mandatory)][string]$Uri)

    $results = New-Object System.Collections.Generic.List[object]
    $nextUri = $Uri
    while ($nextUri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
        if ($response.value) {
            foreach ($item in $response.value) { $results.Add($item) }
        }
        $nextUri = $response.'@odata.nextLink'
    }
    return $results.ToArray()
}

function Add-ParameterEntry {
    param(
        $Value,
        [AllowEmptyString()][string]$Path,
        [System.Collections.Generic.List[string]]$Entries,
        [int]$Depth = 0
    )

    if ($null -eq $Value -or $Depth -ge 12) { return }

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in $Value.Keys) {
            $childPath = if ($Path) { "$Path.$key" } else { [string]$key }
            Add-ParameterEntry -Value $Value[$key] -Path $childPath -Entries $Entries -Depth ($Depth + 1)
        }
        return
    }

    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $values = @($Value)
        $nonEmptyValues = @($values | Where-Object {
                $null -ne $_ -and ($_ -isnot [string] -or -not [string]::IsNullOrWhiteSpace($_))
            })
        if ($nonEmptyValues.Count -eq 0) { return }

        $simpleValues = @($nonEmptyValues | Where-Object {
                $_ -is [string] -or $_.GetType().IsPrimitive -or $_ -is [decimal] -or $_ -is [guid]
            })
        if ($simpleValues.Count -eq $nonEmptyValues.Count) {
            $text = $simpleValues -join ', '
            $Entries.Add("$Path=$text;")
            return
        }

        for ($i = 0; $i -lt $values.Count; $i++) {
            if ($null -ne $values[$i] -and ($values[$i] -isnot [string] -or -not [string]::IsNullOrWhiteSpace($values[$i]))) {
                Add-ParameterEntry -Value $values[$i] -Path "$Path[$i]" -Entries $Entries -Depth ($Depth + 1)
            }
        }
        return
    }

    if ($Value -is [string] -or $Value.GetType().IsPrimitive -or $Value -is [decimal] -or $Value -is [guid]) {
        $text = [string]$Value
        if (-not [string]::IsNullOrWhiteSpace($text)) { $Entries.Add("$Path=$text;") }
        return
    }

    foreach ($property in @($Value.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' })) {
        $childPath = if ($Path) { "$Path.$($property.Name)" } else { $property.Name }
        Add-ParameterEntry -Value $property.Value -Path $childPath -Entries $Entries -Depth ($Depth + 1)
    }
}

function ConvertTo-ParameterText {
    param($Value)

    if ($null -eq $Value) { return '' }
    $entries = New-Object 'System.Collections.Generic.List[string]'
    Add-ParameterEntry -Value $Value -Path '' -Entries $entries
    return ($entries -join ' ')
}

function ConvertTo-SafeFileName {
    param([string]$Name)
    foreach ($zeichen in [System.IO.Path]::GetInvalidFileNameChars()) {
        $Name = $Name.Replace([string]$zeichen, '_')
    }
    $Name = $Name.Trim()
    if ($Name.Length -gt 100) { $Name = $Name.Substring(0, 100) }
    if ([string]::IsNullOrWhiteSpace($Name)) { $Name = 'Unnamed' }
    return $Name
}

function Resolve-Name {
    param(
        [Parameter(Mandatory)][ValidateSet('User', 'Group', 'App', 'Location', 'Role')][string]$Type,
        [Parameter(Mandatory)][string]$Id
    )

    $guid = [guid]::Empty
    if (-not [guid]::TryParse($Id, [ref]$guid)) { return $Id }

    $cacheKey = "$Type|$Id"
    if ($script:NameCache.ContainsKey($cacheKey)) { return $script:NameCache[$cacheKey] }

    $name = $null
    try {
        switch ($Type) {
            'User' {
                $uri = 'https://graph.microsoft.com/v1.0/users/' + $Id + '?$select=displayName,userPrincipalName'
                $r = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop
                $name = [string]$r.userPrincipalName
            }
            'Group' {
                $uri = 'https://graph.microsoft.com/v1.0/groups/' + $Id + '?$select=displayName'
                $r = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop
                $name = $r.displayName
            }
            'App' {
                $uri = ('https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId%20eq%20''{0}''&$select=displayName' -f $Id)
                $r = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop
                if ($r.value -and $r.value.Count -gt 0) { $name = $r.value[0].displayName }
            }
            'Location' {
                if ($script:Locations.ContainsKey($Id)) { $name = $script:Locations[$Id] }
            }
            'Role' {
                if ($script:Roles.ContainsKey($Id)) { $name = $script:Roles[$Id] }
            }
        }
    }
    catch {
        $message = $_.Exception.Message
        if ($Type -in @('User', 'Group')) {
            if ($message -notmatch '(?i)\b404\b|Request_ResourceNotFound|ResourceNotFound') {
                Write-Warning "Could not resolve $Type; counting it as deleted: $message"
            }
            $name = $null
        }
        elseif ($message -match '(?i)\b404\b|Request_ResourceNotFound|ResourceNotFound') {
            $name = "$Id (not found)"
        }
        else {
            Write-Warning "Could not resolve the name for $Type with ID $Id`: $message"
            $name = "$Id (resolution error)"
        }
    }

    if ([string]::IsNullOrWhiteSpace($name) -and $Type -notin @('User', 'Group')) {
        $name = if ($Type -eq 'Role') { "$Id (unresolved)" } else { "$Id (not found)" }
    }

    $script:NameCache[$cacheKey] = $name
    return $name
}

function Convert-List {
    param(
        $Values,
        [ValidateSet('User', 'Group', 'App', 'Location', 'Role', 'Raw')][string]$Type = 'Raw'
    )

    $items = @($Values | Where-Object { $_ })
    if ($items.Count -eq 0) { return '' }

    $unresolvedCount = 0
    $output = @(
        foreach ($item in $items) {
            if ($ResolveNames -and $Type -ne 'Raw') {
                $name = Resolve-Name -Type $Type -Id ([string]$item)
                if ($Type -in @('User', 'Group') -and [string]::IsNullOrWhiteSpace($name)) {
                    $unresolvedCount++
                }
                else {
                    $name
                }
            }
            else {
                [string]$item
            }
        }
    )

    if ($unresolvedCount -gt 0) { $output += "Deleted: $unresolvedCount" }
    return ($output -join '; ')
}

# ---------------------------------------------------------------------------
# 1) Check/install modules
# ---------------------------------------------------------------------------
try {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host 'Microsoft.Graph.Authentication is missing; installing for the current user ...' -ForegroundColor Yellow
        Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}
catch {
    Write-GraphError -ErrorRecord $_ -Context 'Installing/importing Microsoft.Graph.Authentication'
    Write-Host 'Check internet/proxy access and the PowerShell Gallery.' -ForegroundColor Yellow
    exit 1
}

# ---------------------------------------------------------------------------
# 2) Sign in, retrieve, and export
# ---------------------------------------------------------------------------
try {
    $scopes = @('Policy.Read.All', 'User.Read.All', 'Group.Read.All', 'Application.Read.All', 'RoleManagement.Read.Directory')

    try {
        $connection = @{ Scopes = $scopes; NoWelcome = $true }
        if ($TenantId) { $connection['TenantId'] = $TenantId }
        Connect-MgGraph @connection
    }
    catch {
        Write-GraphError -ErrorRecord $_ -Context 'Signing in to Microsoft Graph'
        throw
    }

    $context = Get-MgContext
    Write-Host "Signed in as $($context.Account) (Tenant: $($context.TenantId))" -ForegroundColor Green

    $missingScopes = @($scopes | Where-Object { $context.Scopes -notcontains $_ })
    if ($missingScopes.Count -gt 0) {
        throw "The active Graph session is missing required permissions: $($missingScopes -join ', '). Start a new PowerShell session and grant the requested scopes."
    }

    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }
    $OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).Path

    Write-Host 'Reading Conditional Access policies ...'
    try {
        $policies = @(Get-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies')
    }
    catch {
        Write-GraphError -ErrorRecord $_ -Context 'Retrieving Conditional Access policies'
        throw
    }
    Write-Host "Found $($policies.Count) policies."

    $allPoliciesPath = Join-Path $OutputFolder 'AllPolicies.json'
    $csvPath = Join-Path $OutputFolder 'PolicyOverview.csv'
    $excelPath = Join-Path $OutputFolder 'PolicyOverview.xlsx'
    $oldFileNames = @()

    if (Test-Path -LiteralPath $allPoliciesPath) {
        try {
            $oldPolicies = @(Get-Content -LiteralPath $allPoliciesPath -Raw | ConvertFrom-Json -ErrorAction Stop)
            foreach ($oldPolicy in $oldPolicies) {
                $oldId = [string]$oldPolicy.id
                $validId = [guid]::Empty
                if ($oldPolicy.displayName -and [guid]::TryParse($oldId, [ref]$validId)) {
                    $oldFileNames += '{0}_{1}.json' -f (ConvertTo-SafeFileName -Name $oldPolicy.displayName), $oldId
                }
            }
        }
        catch {
            Write-Warning "Could not read the previous AllPolicies.json; stale files will not be removed: $($_.Exception.Message)"
        }
    }

    $currentFileNames = @{}
    if ($policies.Count -eq 0) {
        Write-Warning 'No policies found. Removing existing overview exports.'
        foreach ($exportPath in @($csvPath, $excelPath)) {
            if (Test-Path -LiteralPath $exportPath) {
                Remove-Item -LiteralPath $exportPath -ErrorAction Stop
            }
        }
    }
    else {
        try {
            $locations = Get-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations'
            foreach ($location in $locations) {
                $script:Locations[[string]$location.id] = [string]$location.displayName
            }
            Write-Host "Loaded $($locations.Count) named locations."
        }
        catch {
            Write-Warning "Could not load named locations; their GUIDs will be retained: $($_.Exception.Message)"
        }

        try {
            $roleDefinitions = @(Get-GraphPaged -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions')
            foreach ($role in $roleDefinitions) {
                $roleId = [string]$role.templateId
                $roleName = [string]$role.displayName
                if (-not [string]::IsNullOrWhiteSpace($roleId) -and -not [string]::IsNullOrWhiteSpace($roleName)) {
                    $script:Roles[$roleId] = $roleName
                }
            }
            Write-Host "Loaded $($roleDefinitions.Count) directory roles."
        }
        catch {
            Write-Warning "Could not load directory roles; role GUIDs will be marked as unresolved in the CSV: $($_.Exception.Message)"
        }

        foreach ($p in $policies) {
            $fileName = '{0}_{1}.json' -f (ConvertTo-SafeFileName -Name $p.displayName), $p.id
            $currentFileNames[$fileName] = $true
            $p | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $OutputFolder $fileName) -Encoding UTF8
        }

        $overview = foreach ($p in $policies) {
            $users = $p.conditions.users
            $applications = $p.conditions.applications
            $locations = $p.conditions.locations
            $conditions = $p.conditions
            $grant = $p.grantControls
            $session = $p.sessionControls

            [pscustomobject][ordered]@{
                displayName                                             = $p.displayName
                id                                                      = $p.id
                templateId                                              = $p.templateId
                state                                                   = $p.state
                createdDateTime                                        = $p.createdDateTime
                modifiedDateTime                                       = $p.modifiedDateTime
                'conditions.applications.includeApplications'          = Convert-List $applications.includeApplications 'App'
                'conditions.applications.excludeApplications'          = Convert-List $applications.excludeApplications 'App'
                'conditions.applications.includeUserActions'            = Convert-List $applications.includeUserActions 'Raw'
                'conditions.applications.applicationFilter'             = ConvertTo-ParameterText $applications.applicationFilter
                'conditions.authenticationFlows.transferMethods'        = $conditions.authenticationFlows.transferMethods
                'conditions.clientApplications.includeServicePrincipals' = Convert-List $conditions.clientApplications.includeServicePrincipals 'Raw'
                'conditions.clientApplications.excludeServicePrincipals' = Convert-List $conditions.clientApplications.excludeServicePrincipals 'Raw'
                'conditions.clientApplications.servicePrincipalFilter'  = ConvertTo-ParameterText $conditions.clientApplications.servicePrincipalFilter
                'conditions.clientAppTypes'                             = Convert-List $conditions.clientAppTypes 'Raw'
                'conditions.devices.deviceFilter'                       = ConvertTo-ParameterText $conditions.devices.deviceFilter
                'conditions.insiderRiskLevels'                          = Convert-List $conditions.insiderRiskLevels 'Raw'
                'conditions.locations.includeLocations'                 = Convert-List $locations.includeLocations 'Location'
                'conditions.locations.excludeLocations'                 = Convert-List $locations.excludeLocations 'Location'
                'conditions.platforms.includePlatforms'                 = Convert-List $conditions.platforms.includePlatforms 'Raw'
                'conditions.platforms.excludePlatforms'                 = Convert-List $conditions.platforms.excludePlatforms 'Raw'
                'conditions.servicePrincipalRiskLevels'                  = Convert-List $conditions.servicePrincipalRiskLevels 'Raw'
                'conditions.signInRiskLevels'                           = Convert-List $conditions.signInRiskLevels 'Raw'
                'conditions.userRiskLevels'                             = Convert-List $conditions.userRiskLevels 'Raw'
                'conditions.users.includeUsers'                         = Convert-List $users.includeUsers 'User'
                'conditions.users.excludeUsers'                         = Convert-List $users.excludeUsers 'User'
                'conditions.users.includeGroups'                        = Convert-List $users.includeGroups 'Group'
                'conditions.users.excludeGroups'                        = Convert-List $users.excludeGroups 'Group'
                'conditions.users.includeRoles'                         = Convert-List $users.includeRoles 'Role'
                'conditions.users.excludeRoles'                         = Convert-List $users.excludeRoles 'Role'
                'conditions.users.includeGuestsOrExternalUsers'         = ConvertTo-ParameterText $users.includeGuestsOrExternalUsers
                'conditions.users.excludeGuestsOrExternalUsers'         = ConvertTo-ParameterText $users.excludeGuestsOrExternalUsers
                'grantControls.builtInControls'                         = Convert-List $grant.builtInControls 'Raw'
                'grantControls.customAuthenticationFactors'             = Convert-List $grant.customAuthenticationFactors 'Raw'
                'grantControls.operator'                                = $grant.operator
                'grantControls.termsOfUse'                              = Convert-List $grant.termsOfUse 'Raw'
                'grantControls.authenticationStrength'                  = ConvertTo-ParameterText $grant.authenticationStrength
                'sessionControls.applicationEnforcedRestrictions'       = ConvertTo-ParameterText $session.applicationEnforcedRestrictions
                'sessionControls.cloudAppSecurity'                      = ConvertTo-ParameterText $session.cloudAppSecurity
                'sessionControls.disableResilienceDefaults'             = $session.disableResilienceDefaults
                'sessionControls.persistentBrowser'                     = ConvertTo-ParameterText $session.persistentBrowser
                'sessionControls.signInFrequency'                       = ConvertTo-ParameterText $session.signInFrequency
            }
        }

        $overviewRows = @($overview)
        if ($overviewRows.Count -gt 0) {
            $columns = @($overviewRows[0].PSObject.Properties.Name)
            $nonEmptyColumns = @(
                foreach ($column in $columns) {
                    $hasValue = $false
                    foreach ($row in $overviewRows) {
                        $value = $row.$column
                        if ($null -ne $value -and ($value -isnot [string] -or -not [string]::IsNullOrWhiteSpace($value))) {
                            $hasValue = $true
                            break
                        }
                    }
                    if ($hasValue) { $column }
                }
            )

            if ($nonEmptyColumns.Count -eq 0) {
                throw 'The overview has no non-empty columns; CSV and Excel export were aborted.'
            }

            $filteredOverview = @($overviewRows | Select-Object -Property $nonEmptyColumns)

            $csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
            $filteredOverview | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Delimiter ';' -Encoding $csvEncoding

            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                Write-Host 'ImportExcel is missing; installing for the current user ...' -ForegroundColor Yellow
                Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
            }
            Import-Module ImportExcel -ErrorAction Stop
            if (Test-Path -LiteralPath $excelPath) {
                Remove-Item -LiteralPath $excelPath -ErrorAction Stop
            }
            $filteredOverview | Export-Excel -Path $excelPath -WorksheetName 'Policies' -TableName 'CAPolicies' -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
        }
    }

    foreach ($fileName in ($oldFileNames | Sort-Object -Unique)) {
        if (-not $currentFileNames.ContainsKey($fileName)) {
            $oldPath = Join-Path $OutputFolder $fileName
            if (Test-Path -LiteralPath $oldPath) {
                Remove-Item -LiteralPath $oldPath -ErrorAction Stop
            }
        }
    }

    ConvertTo-Json -InputObject @($policies) -Depth 20 | Set-Content -LiteralPath $allPoliciesPath -Encoding UTF8

    Write-Host ''
    Write-Host '==================== Summary ====================' -ForegroundColor Cyan
    Write-Host "Exported policies : $($policies.Count)"
    Write-Host "Output folder     : $OutputFolder"
    if ($policies.Count -gt 0) { Write-Host "Excel overview    : $excelPath" }
    Write-Host "Names resolved    : yes"
    Write-Host '=================================================' -ForegroundColor Cyan
}
catch {
    $exitCode = 1
    Write-Host ''
    Write-Host "Export failed: $($_.Exception.Message)" -ForegroundColor Red
}

exit $exitCode
