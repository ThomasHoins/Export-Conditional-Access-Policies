# Export Conditional Access Policies

Version: 1.0.1

This script exports all Conditional Access policies from a Microsoft Entra tenant by using Microsoft Graph. It creates one JSON file per policy, a combined JSON export, a CSV overview, and an Excel workbook with a formatted table.

## Features

- Reads all Conditional Access policies with pagination
- Exports raw policy JSON for each policy and a combined AllPolicies.json
- Builds a CSV overview for policy comparison and reporting
- Creates an Excel overview with auto-sized columns and a formatted table
- Resolves GUID-based references for users, groups, apps, named locations, and directory roles
- Runs without local administrator rights by installing modules for CurrentUser

## Requirements

- PowerShell 5.1 or newer
- Microsoft Graph PowerShell module: Microsoft.Graph.Authentication
- ImportExcel module for Excel export
- A Microsoft Entra tenant with access to Conditional Access policies
- A signed-in Microsoft account with the required Graph permissions

## Required Graph permissions

The script uses delegated authentication. At minimum it requires:

- Policy.Read.All

When name resolution is enabled, it also requests:

- User.Read.All
- Group.Read.All
- Application.Read.All
- RoleManagement.Read.Directory

## Usage

### Basic export

```powershell
.\Export-CAPolicies.ps1
```

### Export to a custom folder

```powershell
.\Export-CAPolicies.ps1 -OutputFolder "C:\Temp\CA-Export"
```

### Export against a specific tenant

```powershell
.\Export-CAPolicies.ps1 -TenantId "<tenant-id-or-domain>"
```

## Output

The script creates the following files in the target folder:

- AllPolicies.json
- PolicyOverview.csv
- PolicyOverview.xlsx
- One JSON file per policy using the safe policy name and ID, for example:
  - Require MFA_12345678-90ab-cdef-1234-567890abcdef.json

## Name resolution

The script can resolve GUID references in the overview table to friendly names for:

- Users
- Groups
- Service principals / apps
- Named locations
- Directory roles

Unresolvable entries are retained as GUIDs or marked as deleted/unresolved when applicable.

## Versioning

The script includes a version value in the script metadata and in the runtime summary output.

## Version History

### 1.0.1
- Reduced script verbosity and simplified Graph error output
- Removed the explicit switch pattern and kept name resolution enabled by default
- Improved script readability while preserving functionality
- Added version metadata and project documentation

### 1.0.0
- Initial release
- Export of all Conditional Access policies
- JSON, CSV, and Excel output
- Optional GUID-to-name resolution

## License

This project is licensed under the MIT License. See the LICENSE file for details.
