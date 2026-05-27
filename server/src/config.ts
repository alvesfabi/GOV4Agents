export const env = {
  port: Number(process.env.PORT ?? 8787),
  spaOrigin: process.env.SPA_ORIGIN ?? 'http://localhost:5173',
  tenantId: required('TENANT_ID'),
  clientId: required('CLIENT_ID'),
  clientSecret: required('CLIENT_SECRET'),
};

function required(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Missing required env var: ${name}`);
  return v;
}
