/**
 * Predefined human identities used to pre-fill the setup role fields so the
 * demo reads with representative names instead of blank UPNs.
 *
 * Defaults target the standard Microsoft 365 demo content-pack users that
 * already exist in the test tenant. Each value can be overridden per tenant
 * via web/.env.local (VITE_SPONSOR_UPN, VITE_APPROVER_UPN, VITE_MANAGER_UPN
 * and their *_NAME counterparts).
 *
 * Story: the sponsor sponsors the agent, the approver approves access
 * requests, and the manager is the sponsor's manager in Entra — who receives
 * the transferred sponsorships when the sponsor is offboarded by the Lifecycle
 * Workflow.
 */
export interface Persona {
  displayName: string;
  upn: string;
}

const env = import.meta.env;

export const personas: { sponsor: Persona; approver: Persona; manager: Persona } = {
  sponsor: {
    displayName: (env.VITE_SPONSOR_NAME as string) || 'Adele Vance',
    upn: (env.VITE_SPONSOR_UPN as string) || 'AdeleV@M365x20582634.OnMicrosoft.com',
  },
  approver: {
    displayName: (env.VITE_APPROVER_NAME as string) || 'Megan Bowen',
    upn: (env.VITE_APPROVER_UPN as string) || 'MeganB@M365x20582634.OnMicrosoft.com',
  },
  manager: {
    displayName: (env.VITE_MANAGER_NAME as string) || 'Miriam Graham',
    upn: (env.VITE_MANAGER_UPN as string) || 'MiriamG@M365x20582634.OnMicrosoft.com',
  },
};
