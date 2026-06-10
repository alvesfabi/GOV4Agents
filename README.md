# GOV4Agents

Guided **Manage / Govern / Protect** Proof of Concept SPA for Microsoft Entra
**Agent ID**, **Governance for Agents**, and **Conditional Access for Agents**.

The app walks a customer through real component creation in their own Entra
tenant (Agent Blueprint, Agent ID, Catalog + Access Package, Lifecycle
Workflow, Custom Security Attribute + Conditional Access policy) and then
guides them through the corresponding governance and protection journeys
against the live components.

## Repo layout

```
/web      Vite + React + TypeScript SPA (Fluent UI v9, MSAL)
/server   Express + TypeScript API (OBO to Microsoft Graph; agent token broker)
/CA.json  Conditional Access policy template (used by the setup wizard)
```

## Prerequisites

- Node.js >= 20
- An Entra tenant where you have privileges to create app registrations,
  custom security attributes, lifecycle workflows, access packages, and
  conditional access policies.

## Getting started (local dev)

```powershell
npm install
npm run dev
```

This launches the SPA on `http://localhost:5173` and the API on
`http://localhost:8787`.

## App registration

Create a single multi-tenant app registration in the test tenant (e.g.
`GOV4Agents`) and configure it as follows.

### 1. Authentication

Add **two** redirect URIs at `http://localhost:5173`:

- one under the **SPA** platform (used by MSAL in the browser)
- one under the **Web** platform (used by the backend for the OBO confidential
  client)

### 2. Expose an API

Add an Application ID URI of `api://<clientId>` and a delegated scope named
`access_as_user`. This is the audience the SPA requests and the backend
validates before performing OBO.

### 3. Client secret

Create a client secret and store it in `server/.env` along with the tenant id
and client id:

```
TENANT_ID=<tenant-guid>
CLIENT_ID=<app-clientId>
CLIENT_SECRET=<secret-value>
PORT=8787
```

Also create `web/.env.local`:

```
VITE_CLIENT_ID=<app-clientId>
VITE_AUTHORITY=https://login.microsoftonline.com/common
```

### Demo personas (optional)

The setup steps pre-fill the sponsor and approver UPN fields with
representative human identities so the demo reads naturally. Defaults target
dedicated, fully-licensed demo accounts in the test tenant (display names
`demo_sponsor` / `demo_approver` / `demo_manager`, UPNs `EntraAgentIDDemo1/2/3`).
Resolution is by **UPN**, so renaming the accounts' display names in Entra does
not affect the demo — only the UPN matters. Override them per tenant in
`web/.env.local`:

```
VITE_SPONSOR_UPN=EntraAgentIDDemo1@<tenant>.onmicrosoft.com
VITE_SPONSOR_NAME=demo_sponsor
VITE_APPROVER_UPN=EntraAgentIDDemo2@<tenant>.onmicrosoft.com
VITE_APPROVER_NAME=demo_approver
VITE_MANAGER_UPN=EntraAgentIDDemo3@<tenant>.onmicrosoft.com
VITE_MANAGER_NAME=demo_manager
```

The sponsor must have a `manager` set in Entra for the Lifecycle Workflow
offboarding demo to transfer sponsorships; set the sponsor's manager to the
manager persona above (e.g. demo_sponsor → demo_manager). These accounts are
used to sign in to MyAccess during the demo, so make sure you can authenticate
as them (reset password or issue a Temporary Access Pass).

To provision all three personas (and the sponsor→manager relationship) in a new
tenant, run the helper script — it prints the `VITE_*` values to paste into
`web/.env.local`:

```powershell
./scripts/New-DemoUsers.ps1            # uses the signed-in admin's domain
./scripts/New-DemoUsers.ps1 -CreateGroup   # also makes a CA-exclusion group
```

The script is idempotent (looks up users by UPN, creates if missing) and uses
device code flow with the `User.ReadWrite.All` scope (plus `Group.ReadWrite.All`
when `-CreateGroup` is passed).

### 4. API permissions

Add the following **delegated** Microsoft Graph permissions and grant admin
consent in the test tenant. The SPA requests these incrementally per feature
(see `web/src/auth/scopes.ts`), and the backend exchanges the user token for a
Graph token via OBO using `.default` (see `server/src/graph/constants.ts`).

Core scopes (required for sign-in and most journeys):

- `User.Read`
- `Directory.Read.All`
- `Directory.ReadWrite.All`
- `Application.Read.All`
- `Application.ReadWrite.All`
- `AppRoleAssignment.ReadWrite.All`
- `EntitlementManagement.ReadWrite.All`
- `LifecycleWorkflows.ReadWrite.All`
- `Policy.Read.All`
- `Policy.ReadWrite.ConditionalAccess`
- `CustomSecAttributeDefinition.ReadWrite.All`
- `CustomSecAttributeAssignment.ReadWrite.All`
- `AuditLog.Read.All`

Agent Identity (preview) scopes — required for the Blueprint / Agent ID setup
steps. These are still in preview in Entra; if they are not yet consentable in
your tenant, the backend falls back to `.default` so the rest of the flow still
works:

- `AgentIdentityBlueprint.Create`
- `AgentIdentityBlueprint.UpdateAuthProperties.All`
- `AgentIdentityBlueprint.AddRemoveCreds.All`
- `AgentIdentityBlueprintPrincipal.Create`

> Note: `User.Read.All` is also used, but as an **application** role assigned
> to the Agent Blueprint principal (inheritable to created Agent IDs) — not as
> a delegated scope on this app registration.

## Flow

1. **Intro** — explains Manage / Govern / Protect.
2. **Setup wizard** — creates a Blueprint, an Agent ID, a Catalog + Access
   Package (Groups.Read.All), an offboarding Lifecycle Workflow, and a
   Conditional Access policy + Custom Security Attribute.
3. **Manage journey** — deep links to the created blueprint and agent.
4. **Govern journey** — agent client-credentials auth, Graph calls, MyAccess
   request flow, LCW offboarding run, lifecycle policies overview.
5. **Protect journey** — enable CA policy, observe block, set CSA tag to
   `approved`, observe success.

## Security notes

The agent client secret is stored only in the backend's in-memory session
store (keyed by the `x-gov4-session` header sent by the SPA). It is never
returned to the browser. The PoC is intended to be run locally; do not deploy
the in-memory secret store to a multi-instance environment without changes.

### Resuming a session

The server session store is in-memory and is cleared whenever the API server
restarts. To let you reopen the app and resume where you left off, the SPA
persists a **non-secret** snapshot of the demo context (created resource ids,
names, and the demo suffix) in the browser's `localStorage`, keyed by the
session id (also stored in `localStorage`). On load the SPA rehydrates the
server session from this snapshot via `POST /api/setup/manual-session`, so
deep links and the Manage/Govern/Protect journeys keep working after a restart.

Client secrets are deliberately excluded from the snapshot (they never leave
the backend), so the agent client-credential token used in the Govern journey
cannot be restored after a server restart — re-run the Agent ID setup step if
you need it again.

## Scripts

| Command | Description |
| --- | --- |
| `npm run dev` | Run web + server in parallel |
| `npm run build` | Build server then web |
| `npm run lint` | Lint the workspace |
| `npm run format` | Format the workspace with Prettier |
