/**
 * Catalog of Microsoft Graph delegated scopes used by each feature. The SPA
 * requests these incrementally, only when the user navigates into a feature.
 */
export const Scopes: Record<string, string[]> = {
  signIn: ['User.Read', 'AppRoleAssignment.ReadWrite.All'],
  groupReadWrite: ['Group.ReadWrite.All'],
  agentBlueprint: ['Application.ReadWrite.All', 'Directory.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All'],
  agentId: ['Application.ReadWrite.All', 'Directory.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All'],
  accessPackage: ['EntitlementManagement.ReadWrite.All'],
  lcw: ['LifecycleWorkflows.ReadWrite.All'],
  conditionalAccess: ['Policy.ReadWrite.ConditionalAccess', 'Application.Read.All'],
  csa: [
    'CustomSecAttributeDefinition.ReadWrite.All',
    'CustomSecAttributeAssignment.ReadWrite.All',
  ],
  audit: ['AuditLog.Read.All'],
};

export type ScopeKey = string;
