function requireEnv(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function parseNumber(name: string, fallback: number): number {
  const raw = process.env[name]?.trim();
  if (!raw) {
    return fallback;
  }

  const parsed = Number(raw);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new Error(`Invalid number in ${name}: ${raw}`);
  }

  return parsed;
}

export interface OutlookConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  refreshToken?: string;
  mailbox: string;
  batchSize: number;
  /** MCP transport: "stdio" (default, local) or "http" (remote, for hosting). */
  transport: "stdio" | "http";
  httpPort: number;
  /** Shared secret required as a Bearer token when transport is "http". */
  httpToken?: string;
}

function parseTransport(): "stdio" | "http" {
  const raw = process.env.MCP_TRANSPORT?.trim().toLowerCase();
  if (!raw || raw === "stdio") {
    return "stdio";
  }
  if (raw === "http") {
    return "http";
  }
  throw new Error(`Invalid MCP_TRANSPORT: ${raw} (expected "stdio" or "http")`);
}

export function getConfig(): OutlookConfig {
  return {
    clientId: requireEnv("OUTLOOK_CLIENT_ID"),
    clientSecret: requireEnv("OUTLOOK_CLIENT_SECRET"),
    tenantId: process.env.OUTLOOK_TENANT_ID || "common",
    refreshToken: process.env.OUTLOOK_REFRESH_TOKEN,
    mailbox: process.env.OUTLOOK_MAILBOX?.trim() || "Inbox",
    batchSize: parseNumber("OUTLOOK_BATCH_SIZE", 20),
    transport: parseTransport(),
    httpPort: parseNumber("MCP_HTTP_PORT", 3000),
    httpToken: process.env.MCP_HTTP_TOKEN?.trim() || undefined
  };
}
