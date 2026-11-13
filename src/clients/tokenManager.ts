import fs from "node:fs";
import https from "node:https";
import path from "node:path";

/**
 * OAuth2 refresh-token handling for the delegated flow.
 *
 * Responsibilities:
 * - Cache access tokens in memory and reuse them until shortly before
 *   they expire (instead of a token POST on every Graph call).
 * - Deduplicate concurrent refreshes for the same scope.
 * - Adopt rotated refresh tokens: Microsoft issues a new refresh token
 *   with most responses and eventually invalidates the old one, so the
 *   newest token is kept in memory and persisted to a state file. The
 *   OUTLOOK_REFRESH_TOKEN env value is only the seed for the first run.
 * - Retry transient failures (network errors, HTTP 429, HTTP 5xx) with
 *   backoff, but fail fast on OAuth errors: a dead refresh token or a
 *   missing consent will not fix itself by retrying.
 *
 * Nothing in this module ever logs token values.
 */

export interface TokenHttpResponse {
  statusCode: number;
  body: string;
}

export interface TokenHttpRequest {
  hostname: string;
  path: string;
  method: string;
  headers: Record<string, string | number>;
}

/** Performs the raw token POST. Injectable so tests can fake the endpoint. */
export type TokenRequestFn = (
  request: TokenHttpRequest,
  postData: string
) => Promise<TokenHttpResponse>;

export const httpsTokenRequest: TokenRequestFn = (request, postData) =>
  new Promise((resolve, reject) => {
    const req = https.request(
      {
        ...request,
        headers: {
          ...request.headers,
          "Content-Length": Buffer.byteLength(postData)
        }
      },
      (res) => {
        let body = "";
        res.on("data", (chunk) => {
          body += chunk;
        });
        res.on("end", () => {
          resolve({ statusCode: res.statusCode ?? 0, body });
        });
      }
    );
    req.on("error", reject);
    req.write(postData);
    req.end();
  });

export class TokenError extends Error {
  /** True when retrying the same request may succeed (network error, 429, 5xx). */
  readonly retryable: boolean;
  /** True when the failure points at a scope/consent problem rather than
   * a transient one; used by the caller to fall back to the base scope. */
  readonly isScopeProblem: boolean;
  /** The OAuth `error` code returned by the token endpoint, when any. */
  readonly oauthError?: string;

  constructor(
    message: string,
    options: { retryable?: boolean; isScopeProblem?: boolean; oauthError?: string } = {}
  ) {
    super(message);
    this.name = "TokenError";
    this.retryable = options.retryable ?? false;
    this.isScopeProblem = options.isScopeProblem ?? false;
    this.oauthError = options.oauthError;
  }
}

/** The stored refresh token (or the client credentials) is dead; only
 * minting a fresh token with scripts/get-auth-token.js recovers. */
export class DeadRefreshTokenError extends TokenError {
  constructor(message: string, oauthError?: string) {
    super(message, { oauthError });
    this.name = "DeadRefreshTokenError";
  }
}

export function isScopeConsentError(error: unknown): boolean {
  return error instanceof TokenError && error.isScopeProblem;
}

export function defaultTokenStatePath(): string {
  return path.join(process.cwd(), "data", "refresh-token.json");
}

const EXPIRY_SKEW_MS = 5 * 60 * 1000;
const MAX_ATTEMPTS = 3;
const BACKOFF_MS = [500, 2000];

function log(message: string): void {
  console.error(`[outlook-mcp] ${message}`);
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** Scope/consent problems surface as invalid_scope, as AADSTS70011
 * (invalid scope value) / AADSTS65001 (no consent), or with an AADSTS
 * description that talks about scope or consent. */
function looksLikeScopeProblem(oauthError: string | undefined, description: string): boolean {
  if (oauthError === "invalid_scope" || oauthError === "consent_required") {
    return true;
  }
  if (/AADSTS(70011|65001)\b/.test(description)) {
    return true;
  }
  return /AADSTS\d+/.test(description) && /scope|consent/i.test(description);
}

export interface TokenManagerOptions {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  /** Seed refresh token from config (OUTLOOK_REFRESH_TOKEN). May be
   * undefined when a token state file from a previous run provides one. */
  seedRefreshToken?: string;
  /** Where rotated refresh tokens are persisted (JSON, mode 0600). */
  stateFilePath: string;
  /** Injectable token-endpoint transport; defaults to real HTTPS. */
  requestFn?: TokenRequestFn;
  /** Injectable delay; defaults to a real timer. */
  sleep?: (ms: number) => Promise<void>;
  /** Injectable clock (milliseconds); defaults to Date.now. */
  now?: () => number;
}

interface CachedToken {
  accessToken: string;
  scope: string;
  expiresAt: number;
}

export class TokenManager {
  private readonly clientId: string;
  private readonly clientSecret: string;
  private readonly tenantId: string;
  private readonly stateFilePath: string;
  private readonly requestFn: TokenRequestFn;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly now: () => number;

  private refreshToken: string | undefined;
  private cached: CachedToken | null = null;
  private readonly inFlight = new Map<string, Promise<string>>();

  constructor(options: TokenManagerOptions) {
    this.clientId = options.clientId;
    this.clientSecret = options.clientSecret;
    this.tenantId = options.tenantId;
    this.stateFilePath = options.stateFilePath;
    this.requestFn = options.requestFn ?? httpsTokenRequest;
    this.sleep =
      options.sleep ??
      ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.now = options.now ?? (() => Date.now());
    this.refreshToken = this.loadRefreshToken(options.seedRefreshToken);
  }

  /** True when any refresh token is available (state file or seed). */
  hasRefreshToken(): boolean {
    return !!this.refreshToken;
  }

  async getAccessToken(scope: string): Promise<string> {
    const cached = this.cached;
    if (
      cached &&
      cached.scope === scope &&
      this.now() < cached.expiresAt - EXPIRY_SKEW_MS
    ) {
      return cached.accessToken;
    }
    const pending = this.inFlight.get(scope);
    if (pending) {
      return pending;
    }
    const promise = this.refresh(scope).finally(() => {
      this.inFlight.delete(scope);
    });
    this.inFlight.set(scope, promise);
    return promise;
  }

  /** The state file wins over the .env seed: it holds the freshest
   * rotated token. A missing or corrupt file falls back to the seed. */
  private loadRefreshToken(seed: string | undefined): string | undefined {
    try {
      if (fs.existsSync(this.stateFilePath)) {
        const parsed = JSON.parse(fs.readFileSync(this.stateFilePath, "utf8"));
        if (typeof parsed?.refreshToken === "string" && parsed.refreshToken.trim()) {
          log(`token: using refresh token from state file ${this.stateFilePath}`);
          return parsed.refreshToken;
        }
        log(
          `token: warning: state file ${this.stateFilePath} holds no usable refresh token; falling back to OUTLOOK_REFRESH_TOKEN`
        );
      }
    } catch (error) {
      log(
        `token: warning: could not read state file ${this.stateFilePath} (${errorText(error)}); falling back to OUTLOOK_REFRESH_TOKEN`
      );
    }
    return seed;
  }

  private async refresh(scope: string): Promise<string> {
    const refreshToken = this.refreshToken;
    if (!refreshToken) {
      throw new Error(
        "No refresh token available. Set OUTLOOK_REFRESH_TOKEN or restore the token state file; mint a token with scripts/get-auth-token.js"
      );
    }

    let lastError: TokenError | null = null;
    for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
      try {
        const { statusCode, body } = await this.requestFn(
          this.buildRequest(),
          this.buildPostData(scope, refreshToken)
        );
        if (statusCode === 200) {
          const parsed = this.parseJsonBody(body);
          if (parsed?.access_token) {
            this.adoptRotation(parsed.refresh_token);
            const expiresIn =
              typeof parsed.expires_in === "number" ? parsed.expires_in : 3600;
            this.cached = {
              accessToken: parsed.access_token,
              scope,
              expiresAt: this.now() + expiresIn * 1000
            };
            return parsed.access_token;
          }
          throw new TokenError(
            `Failed to acquire access token: ${parsed?.error_description || body}`
          );
        }
        throw this.classifyFailure(statusCode, body);
      } catch (error) {
        const tokenError =
          error instanceof TokenError
            ? error
            : new TokenError(`Token request failed: ${errorText(error)}`, {
                retryable: true
              });
        if (!tokenError.retryable) {
          throw tokenError;
        }
        lastError = tokenError;
        if (attempt < MAX_ATTEMPTS) {
          const base = BACKOFF_MS[attempt - 1];
          const delay = Math.round(base * (0.75 + Math.random() * 0.5));
          log(
            `token: attempt ${attempt} of ${MAX_ATTEMPTS} failed (${tokenError.message}); retrying in ${delay}ms`
          );
          await this.sleep(delay);
        }
      }
    }
    throw lastError ?? new TokenError("Failed to acquire access token");
  }

  private buildRequest(): TokenHttpRequest {
    return {
      hostname: "login.microsoftonline.com",
      path: `/${this.tenantId}/oauth2/v2.0/token`,
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      }
    };
  }

  private buildPostData(scope: string, refreshToken: string): string {
    return new URLSearchParams({
      client_id: this.clientId,
      client_secret: this.clientSecret,
      refresh_token: refreshToken,
      grant_type: "refresh_token",
      scope
    }).toString();
  }

  private parseJsonBody(body: string): any {
    try {
      return JSON.parse(body);
    } catch {
      throw new TokenError(
        "Failed to acquire access token: token endpoint returned an unreadable response"
      );
    }
  }

  private classifyFailure(statusCode: number, body: string): TokenError {
    let parsed: any = null;
    try {
      parsed = JSON.parse(body);
    } catch {
      // Not JSON; fall through and classify by status code only.
    }
    const oauthError: string | undefined =
      typeof parsed?.error === "string" ? parsed.error : undefined;
    const description: string = parsed?.error_description || body || `HTTP ${statusCode}`;

    if (statusCode === 429 || statusCode >= 500) {
      return new TokenError(
        `Failed to acquire access token (HTTP ${statusCode}): ${description}`,
        { retryable: true, oauthError }
      );
    }
    if (oauthError === "invalid_grant" || oauthError === "invalid_client") {
      return new DeadRefreshTokenError(
        `The stored refresh token is no longer accepted (${oauthError}: ${description}). ` +
          "It was revoked, expired, or superseded by token rotation, or the client credentials are wrong. " +
          "Mint a new refresh token with scripts/get-auth-token.js and update OUTLOOK_REFRESH_TOKEN in .env; " +
          `you can also delete the token state file (${this.stateFilePath}) so the .env seed is used again.`,
        oauthError
      );
    }
    if (looksLikeScopeProblem(oauthError, description)) {
      return new TokenError(`Failed to acquire access token: ${description}`, {
        isScopeProblem: true,
        oauthError
      });
    }
    return new TokenError(`Failed to acquire access token: ${description}`, {
      oauthError
    });
  }

  /** Microsoft rotates refresh tokens: keep the newest one in memory and
   * persist it so a restart does not resurrect a superseded token. */
  private adoptRotation(candidate: unknown): void {
    if (typeof candidate !== "string" || !candidate || candidate === this.refreshToken) {
      return;
    }
    this.refreshToken = candidate;
    try {
      fs.mkdirSync(path.dirname(this.stateFilePath), { recursive: true });
      const tempPath = `${this.stateFilePath}.tmp-${process.pid}-${Math.random()
        .toString(36)
        .slice(2)}`;
      fs.writeFileSync(
        tempPath,
        `${JSON.stringify(
          { refreshToken: candidate, updatedAt: new Date().toISOString() },
          null,
          2
        )}\n`,
        { mode: 0o600 }
      );
      fs.renameSync(tempPath, this.stateFilePath);
      fs.chmodSync(this.stateFilePath, 0o600);
      log(`token: refresh token rotated; new token persisted to ${this.stateFilePath}`);
    } catch (error) {
      log(
        `token: warning: could not persist rotated refresh token to ${this.stateFilePath} (${errorText(error)}); continuing with the in-memory token`
      );
    }
  }
}
