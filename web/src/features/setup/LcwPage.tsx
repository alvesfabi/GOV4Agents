import { SetupStepPage } from './SetupStepPage';
import { PortalLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';
import { personas } from '../../config/personas';

export function LcwPage() {
  return (
    <SetupStepPage
      title="Step 4 — Lifecycle Workflow"
      intro={
        <>
          Creates the <strong>Offboard agent sponsors</strong> leaver workflow. When a sponsor
          user is offboarded (employeeLeaveDateTime), the workflow notifies the manager and
          co-sponsors, then transfers all of the user's agent sponsorships to their manager. In
          this demo the sponsor <strong>{personas.sponsor.displayName}</strong> reports to{' '}
          <strong>{personas.manager.displayName}</strong>, so sponsorships transfer to{' '}
          {personas.manager.displayName} when {personas.sponsor.displayName} leaves.
        </>
      }
      fields={[
        { name: 'name', label: 'Workflow name', defaultPrefix: 'Offboard agent sponsors ', required: true },
      ]}
      endpoint="/api/setup/lcw"
      submitLabel="Create workflow"
      renderResult={(r) => {
        const lcw = (r as { lcw?: { id?: string; displayName?: string } }).lcw;
        return lcw?.id ? (
          <PortalLink href={portal.lcw(lcw.id, lcw.displayName)}>
            Open workflow in Entra
          </PortalLink>
        ) : null;
      }}
    />
  );
}
