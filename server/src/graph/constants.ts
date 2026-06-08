/**
 * Microsoft Graph delegated scopes used per setup step.
 * The OBO call requests exactly the scopes needed for each operation.
 */
export const GraphScopes = {
  user: ['User.Read'],
  directory: ['Directory.ReadWrite.All'],
  apps: [
    'Application.ReadWrite.All',
    'Directory.ReadWrite.All',
    'AppRoleAssignment.ReadWrite.All',
    'AgentIdentityBlueprint.Create',
    'AgentIdentityBlueprint.UpdateAuthProperties.All',
    'AgentIdentityBlueprint.AddRemoveCreds.All',
    'AgentIdentityBlueprintPrincipal.Create',
  ],
  entitlement: ['EntitlementManagement.ReadWrite.All'],
  lcw: ['LifecycleWorkflows.ReadWrite.All'],
  conditionalAccess: ['Policy.ReadWrite.ConditionalAccess', 'Policy.Read.All'],
  csaDefinition: ['CustomSecAttributeDefinition.ReadWrite.All'],
  csaAssignment: ['CustomSecAttributeAssignment.ReadWrite.All'],
};

/** Microsoft Graph application's appId (constant across all tenants). */
export const MICROSOFT_GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';

/**
 * Microsoft Graph permission ids.
 * - `User.Read.All` (delegated scope id) used for inheritable permissions on
 *   the blueprint.
 * - `Group.Read.All` (application/app-role id) — entitlement-management
 *   catalog role for the Microsoft Graph "API permission" resource is the
 *   application permission so agents (which receive app tokens via FIC)
 *   inherit it as a `roles` claim.
 */
export const GRAPH_DELEGATED_PERMISSIONS = {
  'User.Read.All': 'a154be20-db9c-4678-8ab7-66f6cc099a59',
  'Group.Read.All': '5b567255-7703-4780-807c-7be8301ae99b',
  'Directory.Read.All': '7ab1d382-f21e-4acd-a863-ba3e13f7da61',
};
