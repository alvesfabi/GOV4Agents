import { SetupStepPage } from './SetupStepPage';
import { PortalLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';
import { personas } from '../../config/personas';

export function BlueprintPage() {
  return (
    <SetupStepPage
      title="Step 1 — Agent Blueprint"
      intro={
        <>
          A blueprint standardizes how agents are provisioned. We will create one with the
          sponsor you specify below, and assign <strong>User.Read.All</strong> as an{' '}
          <em>inheritable</em> Microsoft Graph permission so every agent created from this
          blueprint inherits it automatically.
        </>
      }
      fields={[
        { name: 'name', label: 'Blueprint name', defaultPrefix: 'GOV4Agents Blueprint ', required: true },
        {
          name: 'sponsorUpn',
          label: 'Sponsor UPN',
          helper: `User principal name of the user who will sponsor this blueprint (default: ${personas.sponsor.displayName})`,
          defaultValue: personas.sponsor.upn,
          required: true,
        },
      ]}
      endpoint="/api/setup/blueprint"
      submitLabel="Create blueprint"
      renderResult={(r) => {
        const result = r as { blueprintPrincipal?: { id?: string } };
        const principalId = result.blueprintPrincipal?.id;
        return principalId ? (
          <PortalLink href={portal.agentBlueprint(principalId)}>Open in Entra portal</PortalLink>
        ) : null;
      }}
    />
  );
}
