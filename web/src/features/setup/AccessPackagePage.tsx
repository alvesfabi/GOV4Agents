import { SetupStepPage } from './SetupStepPage';
import { PortalLink } from '../../components/Links';
import { portal } from '../../lib/portalLinks';

export function AccessPackagePage() {
  return (
    <SetupStepPage
      title="Step 3 — Catalog + Access Package"
      intro={
        <>
          We create a catalog, add <strong>Microsoft Graph</strong> as a resource, and create{' '}
          <strong>two</strong> access packages so the demo can exercise both the happy path
          and a Separation of Duties guardrail:
          <ul style={{ marginTop: 6, marginBottom: 6 }}>
            <li>
              <strong>Access Package #1 — <code>Group.Read.All</code>.</strong> This is the
              package the sponsor will request on behalf of the agent during the Govern
              journey. After approval, the agent gains <code>Group.Read.All</code> and the{' '}
              <em>"query first 10 groups"</em> call starts to succeed — demonstrating
              just-in-time, governed permission elevation for an agent identity.
            </li>
            <li>
              <strong>Access Package #2 — <code>Directory.Read.All</code> (Separation of
              Duties).</strong> This package is declared <em>incompatible</em> with package
              #1. We never expect it to be granted — its sole purpose is to prove that
              Entitlement Management blocks the request at submission time when a requestor
              already has (or has requested) the conflicting package. This is how you enforce
              toxic-combination policies without relying on reviewer vigilance.
            </li>
          </ul>
          Both packages share the same assignment policy shape (open to users, service
          principals, and agents; sponsor requests on behalf of; one approval stage). Specify
          the approver UPN below.
        </>
      }
      fields={[
        { name: 'name', label: 'Catalog name', defaultPrefix: 'GOV4Agents Catalog ', required: true },
        {
          name: 'approverUpn',
          label: 'Approver UPN',
          helper: 'Email of the user who will approve access requests for this package',
          required: true,
        },
      ]}
      endpoint="/api/setup/access-package"
      submitLabel="Create access packages"
      renderResult={(r) => {
        const result = r as {
          accessPackage?: { id?: string; displayName?: string };
          accessPackage2?: { id?: string; displayName?: string };
          catalog?: { id?: string; displayName?: string };
        };
        return (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {result.accessPackage?.id && (
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <PortalLink
                  href={portal.accessPackage(
                    result.accessPackage.id,
                    result.catalog?.id,
                    result.catalog?.displayName,
                    result.accessPackage.displayName,
                  )}
                >
                  Open Group.Read.All package in Entra
                </PortalLink>
                <PortalLink href={portal.myAccessRequest(result.accessPackage.id)}>
                  MyAccess request URL (Group.Read.All)
                </PortalLink>
              </div>
            )}
            {result.accessPackage2?.id && (
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <PortalLink
                  href={portal.accessPackage(
                    result.accessPackage2.id,
                    result.catalog?.id,
                    result.catalog?.displayName,
                    result.accessPackage2.displayName,
                  )}
                >
                  Open Directory.Read.All package in Entra (SoD: incompatible with Group.Read.All)
                </PortalLink>
                <PortalLink href={portal.myAccessRequest(result.accessPackage2.id)}>
                  MyAccess request URL (Directory.Read.All)
                </PortalLink>
              </div>
            )}
          </div>
        );
      }}
    />
  );
}
