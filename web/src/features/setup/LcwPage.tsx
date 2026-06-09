import { SetupStepPage } from './SetupStepPage';
import { PortalLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';

export function LcwPage() {
  return (
    <SetupStepPage
      title="Step 4 — Lifecycle Workflow"
      intro={
        <>
          Creates the <strong>Offboard agent sponsors</strong> leaver workflow. When a sponsor
          user is offboarded (employeeLeaveDateTime), the workflow notifies the manager and
          co-sponsors, then transfers all of the user's agent sponsorships to their manager.
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
