/**
 * Helpers to build deep links into the Entra portal for resources we create.
 * The links use feature flags / blade routes that surface the *exact* resource
 * (not a generic list).
 */

const ENTRA = 'https://entra.microsoft.com';
const AZURE = 'https://portal.azure.com';
const MYACCESS = 'https://myaccess.microsoft.com';

export const portal = {
  // Agent Blueprint blade for a specific blueprint.
  agentBlueprint(blueprintId: string): string {
    return `${ENTRA}/#view/Microsoft_AAD_RegisteredApps/AgentBlueprintDetails.MenuView/~/overview/objectId/${encodeURIComponent(
      blueprintId,
    )}`;
  },
  // Agent ID (agent identity) blade.
  agentId(objectId: string): string {
    return `${ENTRA}/#view/Microsoft_AAD_RegisteredApps/AgentIdentity.MenuView/~/overview/objectId/${encodeURIComponent(
      objectId,
    )}/menuId/overview`;
  },
  // Lifecycle Workflows: specific workflow. Optional name makes the breadcrumb
  // render correctly in the portal.
  lcw(workflowId: string, workflowName?: string): string {
    const base = `${ENTRA}/#view/Microsoft_AAD_LifecycleManagement/DetailedWorkflowMenuBlade/~/overview/workflowId/${encodeURIComponent(
      workflowId,
    )}`;
    return workflowName
      ? `${base}/workflowName/${encodeURIComponent(workflowName)}`
      : base;
  },
  // Lifecycle policies (agent IDs) blade.
  lifecyclePolicies(): string {
    return `${ENTRA}/?feature.lifecyclePolicies=true#view/Microsoft_AAD_LifecycleManagement/CommonMenuBlade/~/lifecyclePolicies`;
  },
  // Conditional Access policy by id.
  conditionalAccessPolicy(policyId: string): string {
    return `${ENTRA}/#view/Microsoft_AAD_ConditionalAccess/PolicyBlade/policyId/${encodeURIComponent(
      policyId,
    )}`;
  },
  // Custom security attribute set.
  csaSet(setName: string): string {
    return `${ENTRA}/#view/Microsoft_AAD_IAM/AttributeSetBlade/attributeSetName/${encodeURIComponent(
      setName,
    )}`;
  },
  // Entitlement Management: catalog overview.
  catalog(catalogId: string, catalogName?: string): string {
    const base = `${ENTRA}/#view/Microsoft_Azure_ELMAdmin/CatalogMenuBlade/~/overview/catalogId/${encodeURIComponent(
      catalogId,
    )}`;
    return catalogName
      ? `${base}/catalogName/${encodeURIComponent(catalogName)}`
      : base;
  },
  // Access package overview. catalog/name parameters give the portal full
  // context so the breadcrumb and side menu populate correctly.
  accessPackage(
    accessPackageId: string,
    catalogId?: string,
    catalogName?: string,
    accessPackageName?: string,
  ): string {
    let url = `${ENTRA}/#view/Microsoft_Azure_ELMAdmin/EntitlementMenuBlade/~/overview/entitlementId/${encodeURIComponent(
      accessPackageId,
    )}`;
    if (catalogId) url += `/catalogId/${encodeURIComponent(catalogId)}`;
    if (catalogName) url += `/catalogName/${encodeURIComponent(catalogName)}`;
    if (accessPackageName)
      url += `/entitlementName/${encodeURIComponent(accessPackageName)}`;
    return url;
  },
  // MyAccess request URL for a specific access package.
  myAccessRequest(accessPackageId: string): string {
    return `${MYACCESS}/@?#/access-packages/${encodeURIComponent(accessPackageId)}`;
  },
  // MyAccount "Agents" blade for the signed-in sponsor.
  myAccessAgents(): string {
    return 'https://myaccount.microsoft.com/agents';
  },
  azurePortal(): string {
    return AZURE;
  },
};
