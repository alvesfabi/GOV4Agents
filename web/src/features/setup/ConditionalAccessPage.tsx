import { SetupStepPage } from './SetupStepPage';
import { PortalLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';

export function ConditionalAccessPage() {
  return (
    <SetupStepPage
      title="Step 5 — CSA + Conditional Access policy"
      intro={
        <>
          Creates the <strong>AgentsCSA / TAG</strong> custom security attribute (with allowed
          values <code>approved</code> and <code>notApproved</code>) and the{' '}
          <strong>Block not approved agents</strong> Conditional Access policy. The policy
          targets service principals whose <code>TAG</code> attribute is not{' '}
          <code>approved</code>. It is created in <strong>report-only</strong> mode — the
          Protect journey will turn it on later.
        </>
      }
      endpoint="/api/setup/csa-and-ca"
      submitLabel="Create CSA & policy"
      renderResult={(r) => {
        const result = r as {
          caPolicy?: { id?: string };
          csa?: { setName?: string };
        };
        return (
          <>
            {result.caPolicy?.id && (
              <PortalLink href={portal.conditionalAccessPolicy(result.caPolicy.id)}>
                Open Conditional Access policy
              </PortalLink>
            )}{' '}
            {result.csa?.setName && (
              <PortalLink href={portal.csaSet(result.csa.setName)}>
                Open custom attribute set
              </PortalLink>
            )}
          </>
        );
      }}
    />
  );
}
