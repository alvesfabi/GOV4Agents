import { SetupStepPage } from './SetupStepPage';
import { PortalLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';
import { personas } from '../../config/personas';

export function AgentPage() {
  return (
    <SetupStepPage
      title="Step 2 — Agent ID"
      intro={
        <>
          Create an Agent ID linked to the blueprint from Step 1. Specify the sponsor below. A
          client secret is generated and stored securely on the backend (in memory) so the
          Govern journey can authenticate as the agent later.
        </>
      }
      fields={[
        { name: 'name', label: 'Agent name', defaultPrefix: 'GOV4Agents Agent ', required: true },
        {
          name: 'sponsorUpn',
          label: 'Sponsor UPN',
          helper: `User principal name of the user who will sponsor this agent (default: ${personas.sponsor.displayName})`,
          defaultValue: personas.sponsor.upn,
          required: true,
        },
      ]}
      endpoint="/api/setup/agent"
      submitLabel="Create agent"
      renderResult={(r) => {
        const id = (r as { agent?: { id?: string } }).agent?.id;
        return id ? <PortalLink href={portal.agentId(id)}>Open in Entra portal</PortalLink> : null;
      }}
    />
  );
}
