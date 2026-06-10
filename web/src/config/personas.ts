/**
 * Predefined human identities used to pre-fill the setup role fields so the
 * demo reads with representative names instead of blank UPNs.
 *
 * Defaults target dedicated, fully-licensed demo accounts that exist in the
 * test tenant (the EntraAgentIDDemo1/2/3 UPNs, display names demo_sponsor /
 * demo_approver / demo_manager). Resolution is by UPN, so renaming the display
 * names in Entra does not affect the demo. Each value can be overridden per
 * tenant via web/.env.local (VITE_SPONSOR_UPN, VITE_APPROVER_UPN,
 * VITE_MANAGER_UPN and their *_NAME counterparts).
 *
 * Story: the sponsor sponsors the agent, the approver approves access
 * requests, and the manager is the sponsor's manager in Entra — who receives
 * the transferred sponsorships when the sponsor is offboarded by the Lifecycle
 * Workflow. (Set the sponsor's manager to the manager persona in Entra.)
 */
export interface Persona {
  displayName: string;
  upn: string;
}

const env = import.meta.env;

export const personas: { sponsor: Persona; approver: Persona; manager: Persona } = {
  sponsor: {
    displayName: (env.VITE_SPONSOR_NAME as string) || 'demo_sponsor',
    upn: (env.VITE_SPONSOR_UPN as string) || 'EntraAgentIDDemo1@M365x20582634.onmicrosoft.com',
  },
  approver: {
    displayName: (env.VITE_APPROVER_NAME as string) || 'demo_approver',
    upn: (env.VITE_APPROVER_UPN as string) || 'EntraAgentIDDemo2@M365x20582634.onmicrosoft.com',
  },
  manager: {
    displayName: (env.VITE_MANAGER_NAME as string) || 'demo_manager',
    upn: (env.VITE_MANAGER_UPN as string) || 'EntraAgentIDDemo3@M365x20582634.onmicrosoft.com',
  },
};
