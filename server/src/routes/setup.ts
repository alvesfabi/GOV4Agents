import { Router } from 'express';
import { requireBearer, type AuthedRequest } from '../auth/bearer.js';
import { attachSession, type SessionedRequest } from '../store/middleware.js';
import { setSession } from '../store/session.js';
import { GraphClient, GraphError } from '../graph/client.js';
import {
  GraphScopes,
  MICROSOFT_GRAPH_APP_ID,
  GRAPH_DELEGATED_PERMISSIONS,
} from '../graph/constants.js';
import { acquireAgentToken } from '../auth/msal.js';
import { env } from '../config.js';

export const setupRouter = Router();

setupRouter.use(requireBearer, attachSession);

type Req = AuthedRequest & SessionedRequest;

// ---------------------------------------------------------------------------
// 1. Agent Blueprint
// ---------------------------------------------------------------------------

setupRouter.post('/blueprint', async (req: Req, res) => {
  try {
    const { name, sponsorUpn } = req.body ?? {};
    if (!name) return res.status(400).json({ error: 'name required' });
    if (!sponsorUpn) return res.status(400).json({ error: 'sponsorUpn required' });
    const g = new GraphClient(req.userAccessToken!);

    const sponsor = await g.call<{ id: string; userPrincipalName: string }>(
      GraphScopes.user,
      `/users/${encodeURIComponent(sponsorUpn)}?$select=id,userPrincipalName`,
    );

    // Create the blueprint. The signed-in user is auto-added as owner.
    // sponsors must be specified at creation time.
    const blueprint = await g.call<{ id: string; displayName: string; appId: string }>(
      GraphScopes.apps,
      '/applications/microsoft.graph.agentIdentityBlueprint',
      {
        method: 'POST',
        version: 'beta',
        body: {
          displayName: name,
          'sponsors@odata.bind': [`https://graph.microsoft.com/beta/users/${sponsor.id}`],
        },
      },
    );

    // Create the blueprint **principal** (the service principal that backs
    // this blueprint). Without this, agent identities cannot be created from
    // the blueprint.
    const blueprintPrincipal = await g.call<{ id: string; appId: string }>(
      GraphScopes.apps,
      '/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal',
      {
        method: 'POST',
        version: 'beta',
        body: { appId: blueprint.appId },
      },
    );

    // Resolve the Microsoft Graph service principal so we can grant
    // User.Read.All to the blueprint principal directly. Inheritable
    // permissions only flow what the blueprint principal already has.
    const graphSpForBp = await g.call<{
      value: Array<{
        id: string;
        appRoles: Array<{ id: string; value: string }>;
      }>;
    }>(
      GraphScopes.apps,
      `/servicePrincipals?$filter=appId eq '${MICROSOFT_GRAPH_APP_ID}'&$select=id,appRoles`,
    );
    const graphSpId = graphSpForBp.value[0]?.id;
    if (!graphSpId) throw new Error('Microsoft Graph service principal not found');
    const userReadAllRole = graphSpForBp.value[0]?.appRoles.find(
      (r) => r.value === 'User.Read.All',
    );
    if (!userReadAllRole) throw new Error('User.Read.All app role not found on Microsoft Graph');

    // Grant User.Read.All as an APPLICATION permission (app role assignment)
    // on the blueprint principal. The agent ID's FIC flow produces an app
    // token whose authorization comes from app roles ("roles" claim), not
    // from delegated scopes — so the blueprint must hold the role itself.
    try {
      await g.call(
        GraphScopes.apps,
        `/servicePrincipals/${blueprintPrincipal.id}/appRoleAssignments`,
        {
          method: 'POST',
          body: {
            principalId: blueprintPrincipal.id,
            resourceId: graphSpId,
            appRoleId: userReadAllRole.id,
          },
        },
      );
    } catch (err) {
      if (!(err instanceof GraphError && err.status === 409)) throw err;
    }

    // Configure inheritable permissions on the blueprint: agents inherit all
    // app roles the principal has (currently just User.Read.All) and no
    // delegated scopes. Path/body per
    // https://learn.microsoft.com/entra/agent-id/configure-inheritable-permissions-blueprints
    await g.call(
      GraphScopes.apps,
      `/applications/microsoft.graph.agentIdentityBlueprint/${blueprint.id}/inheritablePermissions`,
      {
        method: 'POST',
        body: {
          resourceAppId: MICROSOFT_GRAPH_APP_ID,
          inheritableScopes: {
            '@odata.type': '#microsoft.graph.noScopes',
            kind: 'none',
          },
          inheritableRoles: {
            '@odata.type': '#microsoft.graph.allAllowedRoles',
            kind: 'allAllowed',
          },
        },
      },
    );

    // Add a client secret on the blueprint's application object. Per the
    // agent identity preview, credentials cannot be added to agent service
    // principals — they must live on the blueprint and are inherited by every
    // agent created from it.
    const blueprintSecret = await g.call<{
      secretText: string;
      keyId: string;
      endDateTime: string;
    }>(GraphScopes.apps, `/applications/${blueprint.id}/addPassword`, {
      method: 'POST',
      body: { passwordCredential: { displayName: 'gov4agents demo secret' } },
    });

    setSession(req.sessionId!, {
      blueprint: {
        id: blueprint.id,
        name: blueprint.displayName,
        appId: blueprint.appId,
        principalId: blueprintPrincipal.id,
        clientSecret: blueprintSecret.secretText,
      },
    });

    return res.json({ ok: true, blueprint, blueprintPrincipal, sponsor });
  } catch (err) {
    return failure(res, err);
  }
});

// ---------------------------------------------------------------------------
// 2. Agent ID
// ---------------------------------------------------------------------------

setupRouter.post('/agent', async (req: Req, res) => {
  try {
    const { name, sponsorUpn } = req.body ?? {};
    if (!name) return res.status(400).json({ error: 'name required' });
    if (!sponsorUpn) return res.status(400).json({ error: 'sponsorUpn required' });
    const blueprint = req.sessionData?.blueprint;
    if (!blueprint) return res.status(400).json({ error: 'create the blueprint first' });

    const g = new GraphClient(req.userAccessToken!);
    const sponsor = await g.call<{ id: string; userPrincipalName: string }>(
      GraphScopes.user,
      `/users/${encodeURIComponent(sponsorUpn)}?$select=id,userPrincipalName`,
    );

    // Create the agent identity (a specialized servicePrincipal). The signed-in
    // user is auto-added as owner. Credentials are inherited from the
    // blueprint — agents themselves cannot hold passwordCredentials.
    const agent = await g.call<{
      id: string;
      appId: string;
      displayName: string;
      agentIdentityBlueprintId: string;
    }>(GraphScopes.apps, '/servicePrincipals/microsoft.graph.agentIdentity', {
      method: 'POST',
      version: 'beta',
      body: {
        displayName: name,
        agentIdentityBlueprintId: blueprint.id,
        'sponsors@odata.bind': [`https://graph.microsoft.com/beta/users/${sponsor.id}`],
      },
    });

    setSession(req.sessionId!, {
      agent: {
        id: agent.id,
        clientId: agent.appId,
        name: agent.displayName,
        clientSecret: blueprint.clientSecret,
        tenantId: env.tenantId,
        sponsorId: sponsor.id,
        sponsorUpn: sponsor.userPrincipalName,
      },
    });

    return res.json({
      ok: true,
      agent: {
        id: agent.id,
        appId: agent.appId,
        displayName: agent.displayName,
        agentIdentityBlueprintId: agent.agentIdentityBlueprintId,
      },
      sponsor,
      secretStored: true,
    });
  } catch (err) {
    return failure(res, err);
  }
});

// ---------------------------------------------------------------------------
// 3. Catalog + Access Package (Groups.Read.All)
// ---------------------------------------------------------------------------

setupRouter.post('/access-package', async (req: Req, res) => {
  try {
    const { approverUpn } = req.body ?? {};
    if (!approverUpn) return res.status(400).json({ error: 'approverUpn required' });
    const g = new GraphClient(req.userAccessToken!);

    const suffix = String(Math.floor(Math.random() * 900) + 100);

    // 3a. Catalog (idempotent: reuse if a catalog with this name already exists)
    const CATALOG_NAME = `GOV4Agents Catalog ${suffix}`;
    const existingCatalog = await g.call<{ value: Array<{ id: string; displayName: string }> }>(
      GraphScopes.entitlement,
      `/identityGovernance/entitlementManagement/catalogs?$filter=displayName eq '${CATALOG_NAME}'&$select=id,displayName`,
    );
    const catalog =
      existingCatalog.value[0] ??
      (await g.call<{ id: string; displayName: string }>(
        GraphScopes.entitlement,
        '/identityGovernance/entitlementManagement/catalogs',
        {
          method: 'POST',
          body: {
            displayName: CATALOG_NAME,
            description: 'Catalog created by GOV4Agents PoC',
            isExternallyVisible: true,
            state: 'published',
          },
        },
      ));

    // Ensure catalog is published so the access package shows up in My Access.
    await g.call(
      GraphScopes.entitlement,
      `/identityGovernance/entitlementManagement/catalogs/${catalog.id}`,
      { method: 'PATCH', body: { state: 'published' } },
    );

    // 3b. Resolve the approver user object id.
    const approver = await g.call<{ id: string; userPrincipalName: string }>(
      GraphScopes.directory,
      `/users/${encodeURIComponent(approverUpn)}?$select=id,userPrincipalName`,
    );

    // 3c. Add Microsoft Graph as a catalog resource.
    try {
      await g.call(
        GraphScopes.entitlement,
        '/identityGovernance/entitlementManagement/resourceRequests',
        {
          method: 'POST',
          body: {
            requestType: 'adminAdd',
            catalog: { id: catalog.id },
            resource: {
              originId: MICROSOFT_GRAPH_APP_ID,
              originSystem: 'OAuthApplication',
            },
          },
        },
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      if (!/already|conflict/i.test(msg)) throw err;
    }

    // Wait for the catalog resource entry to materialize (async).
    let graphCatalogResourceId: string | undefined;
    for (let i = 0; i < 12; i++) {
      const found = await g.call<{ value: Array<{ id: string; originId?: string }> }>(
        GraphScopes.entitlement,
        `/identityGovernance/entitlementManagement/catalogs/${catalog.id}/resources?$filter=originId eq '${MICROSOFT_GRAPH_APP_ID}'&$select=id,displayName,originId`,
      );
      if (found.value.length > 0) {
        graphCatalogResourceId = found.value[0].id;
        break;
      }
      await sleep(5000);
    }
    if (!graphCatalogResourceId) {
      throw new Error('Microsoft Graph resource was not found in the catalog after waiting.');
    }

    // Helper: create an access package + role scope (one permission) + policy.
    const createAp = async (params: {
      displayName: string;
      description: string;
      permissionName: 'Group.Read.All' | 'Group.ReadWrite.All';
      policyDisplayName: string;
    }) => {
      const permissionId = GRAPH_DELEGATED_PERMISSIONS[params.permissionName];

      const ap = await g.call<{ id: string; displayName: string }>(
        GraphScopes.entitlement,
        '/identityGovernance/entitlementManagement/accessPackages',
        {
          method: 'POST',
          body: {
            displayName: params.displayName,
            description: params.description,
            catalog: { id: catalog.id },
            isHidden: false,
          },
        },
      );

      await g.call(
        GraphScopes.entitlement,
        `/identityGovernance/entitlementManagement/accessPackages/${ap.id}/resourceRoleScopes`,
        {
          method: 'POST',
          body: {
            role: {
              displayName: params.permissionName,
              originId: permissionId,
              originSystem: 'OauthApplication',
              resource: {
                id: graphCatalogResourceId,
                originId: MICROSOFT_GRAPH_APP_ID,
                originSystem: 'OauthApplication',
              },
            },
            scope: {
              displayName: 'Root',
              description: 'Root Scope',
              originId: MICROSOFT_GRAPH_APP_ID,
              originSystem: 'OauthApplication',
              isRootScope: true,
            },
          },
        },
      );

      const policyBody = {
        displayName: params.policyDisplayName,
        description: 'Sponsor requests access on behalf of an agent.',
        accessPackage: { id: ap.id },
        allowedTargetScope: 'allDirectoryAgentIdentities',
        specificAllowedTargets: [],
        expiration: { type: 'afterDuration', duration: 'P365D', endDateTime: null },
        requestorSettings: {
          enableTargetsToSelfAddAccess: false,
          enableTargetsToSelfUpdateAccess: false,
          enableTargetsToSelfRemoveAccess: true,
          allowCustomAssignmentSchedule: true,
          enableOnBehalfRequestorsToAddAccess: true,
          enableOnBehalfRequestorsToUpdateAccess: true,
          enableOnBehalfRequestorsToRemoveAccess: true,
          onBehalfRequestors: [
            {
              '@odata.type': '#microsoft.graph.targetAgentIdentitySponsorsOrOwners',
              isBackup: false,
            },
          ],
        },
        requestApprovalSettings: {
          isApprovalRequiredForAdd: true,
          isApprovalRequiredForUpdate: false,
          isRequestorJustificationRequired: true,
          stages: [
            {
              durationBeforeAutomaticDenial: 'P14D',
              isApproverJustificationRequired: true,
              isEscalationEnabled: false,
              durationBeforeEscalation: 'PT0S',
              primaryApprovers: [
                {
                  '@odata.type': '#microsoft.graph.singleUser',
                  userId: approver.id,
                },
              ],
              fallbackPrimaryApprovers: [],
              escalationApprovers: [],
              fallbackEscalationApprovers: [],
            },
          ],
        },
        reviewSettings: null,
        questions: [],
        notificationSettings: { isAssignmentNotificationDisabled: false },
      };

      const policy = await g.call<{ id: string; displayName: string }>(
        GraphScopes.entitlement,
        '/identityGovernance/entitlementManagement/assignmentPolicies',
        { method: 'POST', body: policyBody },
      );

      // PUT re-index workaround so "Requesting for Sponsored agent" appears.
      await g.call(
        GraphScopes.entitlement,
        `/identityGovernance/entitlementManagement/assignmentPolicies/${policy.id}`,
        { method: 'PUT', body: policyBody },
      );

      return { ap, policy };
    };

    // 3d. First access package — Group.Read.All
    const ap1 = await createAp({
      displayName: `Agent — Group.Read.All ${suffix}`,
      description: 'Grants Group.Read.All to agents on demand.',
      permissionName: 'Group.Read.All',
      policyDisplayName: `Agents — sponsor requested (Group.Read.All) ${suffix}`,
    });

    // 3e. Second access package — Group.ReadWrite.All (incompatible with ap1)
    const ap2 = await createAp({
      displayName: `Agent — Group.ReadWrite.All ${suffix}`,
      description:
        'Grants Group.ReadWrite.All to agents on demand. Incompatible with the Group.Read.All package (Separation of Duties): an agent that reads group membership should not also be able to modify it.',
      permissionName: 'Group.ReadWrite.All',
      policyDisplayName: `Agents — sponsor requested (Group.ReadWrite.All) ${suffix}`,
    });

    // 3f. Separation of Duties — declare ap1 as incompatible with ap2 so any
    // agent already assigned to ap1 (or with an open request) cannot request ap2.
    await g.call(
      GraphScopes.entitlement,
      `/identityGovernance/entitlementManagement/accessPackages/${ap2.ap.id}/incompatibleAccessPackages/$ref`,
      {
        method: 'POST',
        body: {
          '@odata.id': `https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages/${ap1.ap.id}`,
        },
      },
    );

    setSession(req.sessionId!, {
      catalog: { id: catalog.id, name: catalog.displayName },
      accessPackage: {
        id: ap1.ap.id,
        name: ap1.ap.displayName,
        approverUpn: approver.userPrincipalName,
      },
      accessPackage2: {
        id: ap2.ap.id,
        name: ap2.ap.displayName,
        approverUpn: approver.userPrincipalName,
      },
    });

    return res.json({
      ok: true,
      catalog,
      accessPackage: ap1.ap,
      accessPackage2: ap2.ap,
      approver,
      policy: ap1.policy,
      policy2: ap2.policy,
      separationOfDuties: { incompatibleWith: ap1.ap.id },
    });
  } catch (err) {
    return failure(res, err);
  }
});

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

// ---------------------------------------------------------------------------
// 4. Lifecycle Workflow — Offboard agent sponsors (leaver / time-based)
// ---------------------------------------------------------------------------

const LCW_TASKS = {
  sendEmailToManager: 'b8c4e1f9-3a7d-4b2e-9c5f-8d6a9b1c2e3f',
  sendEmailToCoSponsors: 'ad3b85cd-75b1-43e7-b4b9-0e52faba3944',
  transferAgentSponsorshipsToManager: 'b8f4c3d5-9e7a-4b1c-8f2d-6a5e8b9c7f4a',
};

setupRouter.post('/lcw', async (req: Req, res) => {
  try {
    const g = new GraphClient(req.userAccessToken!);
    const suffix = String(Math.floor(Math.random() * 900) + 100);
    const wf = await g.call<{ id: string; displayName: string }>(
      GraphScopes.lcw,
      '/identityGovernance/lifecycleWorkflows/workflows',
      {
        method: 'POST',
        body: {
          category: 'leaver',
          displayName: `Offboard agent sponsors ${suffix}`,
          description: 'Execute sponsorship transition tasks when an agent sponsor leaves',
          isEnabled: true,
          isSchedulingEnabled: true,
          executionConditions: {
            '@odata.type': '#microsoft.graph.identityGovernance.triggerAndScopeBasedConditions',
            scope: {
              '@odata.type': '#microsoft.graph.identityGovernance.ruleBasedSubjectSet',
              rule: '(accountEnabled eq true)',
            },
            trigger: {
              '@odata.type': '#microsoft.graph.identityGovernance.timeBasedAttributeTrigger',
              timeBasedAttribute: 'employeeLeaveDateTime',
              offsetInDays: 0,
            },
          },
          tasks: [
            {
              category: 'leaver',
              continueOnError: false,
              description: 'Transfer agent sponsorships to manager',
              displayName: 'Transfer agent sponsorships to manager',
              isEnabled: true,
              taskDefinitionId: LCW_TASKS.transferAgentSponsorshipsToManager,
              arguments: [],
            },
          ],
        },
      },
    );

    setSession(req.sessionId!, { lcw: { id: wf.id, name: wf.displayName } });
    return res.json({ ok: true, lcw: wf });
  } catch (err) {
    return failure(res, err);
  }
});

// ---------------------------------------------------------------------------
// 5. Custom security attribute + Conditional Access policy
// ---------------------------------------------------------------------------

setupRouter.post('/csa-and-ca', async (req: Req, res) => {
  try {
    const g = new GraphClient(req.userAccessToken!);
    const suffix = String(Math.floor(Math.random() * 900) + 100);
    const setName = `AgentsCSA${suffix}`;
    const attributeName = `TAG${suffix}`;

    // Try to create the attribute set; if it already exists Graph returns 409
    // and we can move on.
    await ensureCreated(
      g.call(GraphScopes.csaDefinition, '/directory/attributeSets', {
        method: 'POST',
        body: { id: setName, description: 'Custom attributes for agent governance', maxAttributesPerSet: 10 },
        expectStatuses: [409],
      }),
    );

    // Create the TAG attribute on the set with allowed values approved/notApproved.
    await ensureCreated(
      g.call(GraphScopes.csaDefinition, '/directory/customSecurityAttributeDefinitions', {
        method: 'POST',
        body: {
          attributeSet: setName,
          name: attributeName,
          description: 'Approval tag used by Conditional Access for agents',
          type: 'String',
          status: 'Available',
          isCollection: false,
          isSearchable: true,
          usePreDefinedValuesOnly: true,
          allowedValues: [
            { id: 'approved', isActive: true },
            { id: 'notApproved', isActive: true },
          ],
        },
        expectStatuses: [409],
      }),
    );

    // Create the Conditional Access policy targeting Agent ID service
    // principals filtered by the CSA we just created. Excluded = those with
    // CSA.TAG = "approved"; included = everything else (default deny).
    const caBody = {
      displayName: `Default deny agents except approved (Agent ID only) ${suffix}`,
      state: 'enabledForReportingButNotEnforced',
      conditions: {
        userRiskLevels: [],
        signInRiskLevels: [],
        clientAppTypes: ['all'],
        platforms: null,
        locations: null,
        devices: null,
        authenticationFlows: null,
        applications: {
          includeApplications: ['All'],
          excludeApplications: [],
          includeUserActions: [],
          includeAuthenticationContextClassReferences: [],
          applicationFilter: null,
        },
        users: {
          includeUsers: ['None'],
          excludeUsers: [],
          includeGroups: [],
          excludeGroups: [],
          includeRoles: [],
          excludeRoles: [],
          includeGuestsOrExternalUsers: null,
          excludeGuestsOrExternalUsers: null,
        },
        clientApplications: {
          includeServicePrincipals: [],
          excludeServicePrincipals: [],
          includeAgentIdServicePrincipals: ['All'],
          agentIdServicePrincipalFilter: {
            mode: 'exclude',
            rule: `CustomSecurityAttribute.${setName}_${attributeName} -eq "approved"`,
          },
        },
      },
      grantControls: {
        operator: 'OR',
        builtInControls: ['block'],
        customAuthenticationFactors: [],
        termsOfUse: [],
        authenticationStrength: null,
      },
      sessionControls: null,
    };

    const policy = await g.call<{ id: string; displayName: string }>(
      GraphScopes.conditionalAccess,
      '/identity/conditionalAccess/policies',
      { method: 'POST', body: caBody, version: 'beta' },
    );

    setSession(req.sessionId!, {
      csa: { setName, attributeName },
      caPolicy: { id: policy.id, displayName: policy.displayName },
    });

    return res.json({ ok: true, csa: { setName, attributeName }, caPolicy: policy });
  } catch (err) {
    return failure(res, err);
  }
});

async function ensureCreated<T>(p: Promise<T>): Promise<T | undefined> {
  try {
    return await p;
  } catch (err) {
    if (err instanceof GraphError && err.status === 409) return undefined;
    throw err;
  }
}

// ---------------------------------------------------------------------------
// 6. Summary
// ---------------------------------------------------------------------------

setupRouter.get('/summary', (req: Req, res) => {
  const s = req.sessionData ?? {};
  res.json({
    ok: true,
    blueprint: s.blueprint ?? null,
    agent: s.agent
      ? { id: s.agent.id, clientId: s.agent.clientId, name: s.agent.name, hasSecret: Boolean(s.agent.clientSecret) }
      : null,
    catalog: s.catalog ?? null,
    accessPackage: s.accessPackage ?? null,
    accessPackage2: s.accessPackage2 ?? null,
    lcw: s.lcw ?? null,
    csa: s.csa ?? null,
    caPolicy: s.caPolicy ?? null,
    tenantId: env.tenantId,
  });
});

// ---------------------------------------------------------------------------
// Diagnostic: agent token via client credentials. Used by Govern journey.
// ---------------------------------------------------------------------------

setupRouter.post('/agent/token', async (req: Req, res) => {
  try {
    const agent = req.sessionData?.agent;
    if (!agent?.clientSecret) return res.status(400).json({ error: 'no agent secret in session' });
    const result = await acquireAgentToken({
      tenantId: agent.tenantId ?? env.tenantId,
      clientId: agent.clientId,
      clientSecret: agent.clientSecret,
    });
    res.json({
      accessToken: result.accessToken,
      expiresOn: result.expiresOn,
      tokenType: result.tokenType,
      scopes: result.scopes,
    });
  } catch (err) {
    return failure(res, err);
  }
});

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function failure(res: import('express').Response, err: unknown): void {
  if (err instanceof GraphError) {
    res.status(err.status).json({ error: err.message, body: err.body });
    return;
  }
  if (err instanceof Error) {
    res.status(500).json({ error: err.message });
    return;
  }
  res.status(500).json({ error: 'unknown error' });
}
