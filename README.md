    Exports all Conditional Access policies from a tenant using Microsoft Graph.

    - Uses delegated authentication (Connect-MgGraph), scope: Policy.Read.All.
    - Runs without local administrator privileges (installs modules for CurrentUser).
    - Reads /v1.0/identity/conditionalAccess/policies, including pagination.
    - Writes one JSON file per policy, AllPolicies.json, PolicyOverview.csv,
      and PolicyOverview.xlsx as a formatted table.
    - The overview includes documented policy, condition, grant, and session-control
      fields returned by Graph v1.0. Nested values use Parameter=Value; format,
      and columns that are empty for every policy are removed.
    - Optional (-ResolveNames): resolves user, group, app, directory role,
      and named location GUIDs in the overview to readable names. Requires
      User.Read.All, Group.Read.All, Application.Read.All, and
      RoleManagement.Read.Directory for directory roles.
