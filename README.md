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

### 4. API permissions

Add the following **delegated** Microsoft Graph permissions and grant admin
consent in the test tenant:

- `User.Read`
- `Directory.ReadWrite.All`
- `Application.ReadWrite.All`
- `EntitlementManagement.ReadWrite.All`
- `LifecycleWorkflows.ReadWrite.All`
- `Policy.ReadWrite.ConditionalAccess`
- `CustomSecAttributeAssignment.ReadWrite.All`
- `CustomSecAttributeDefinition.ReadWrite.All`
- `AuditLog.Read.All`

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

## Scripts

| Command | Description |
| --- | --- |
| `npm run dev` | Run web + server in parallel |
| `npm run build` | Build server then web |
| `npm run lint` | Lint the workspace |
| `npm run format` | Format the workspace with Prettier |
