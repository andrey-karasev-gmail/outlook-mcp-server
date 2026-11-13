import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";

import {
  DeadRefreshTokenError,
  TokenManager,
  isScopeConsentError,
  type TokenHttpRequest,
  type TokenHttpResponse,
  type TokenManagerOptions,
  type TokenRequestFn
} from "../src/clients/tokenManager.js";
import { GraphEmailClient } from "../src/clients/graphClient.js";
import type { OutlookConfig } from "../src/config.js";

const EXTENDED_SCOPE =
  "offline_access mail.read mail.readwrite mailboxsettings.readwrite";

function tempStatePath(): { dir: string; statePath: string } {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "outlook-token-test-"));
  return { dir, statePath: path.join(dir, "refresh-token.json") };
}

type Handler = (
  request: TokenHttpRequest,
  body: string,
  callNumber: number
) => TokenHttpResponse | Promise<TokenHttpResponse>;

function fakeRequest(handler: Handler) {
  const calls: { request: TokenHttpRequest; body: string }[] = [];
  const fn: TokenRequestFn = async (request, body) => {
    calls.push({ request, body });
    return handler(request, body, calls.length);
  };
  return { fn, calls };
}

function sleepRecorder() {
  const sleeps: number[] = [];
  const sleep = async (ms: number) => {
    sleeps.push(ms);
  };
  return { sleep, sleeps };
}

function options(
  statePath: string,
  overrides: Partial<TokenManagerOptions> = {}
): TokenManagerOptions {
  return {
    clientId: "test-client-id",
    clientSecret: "test-client-secret",
    tenantId: "common",
    seedRefreshToken: "seed-refresh-token",
    stateFilePath: statePath,
    sleep: async () => {},
    ...overrides
  };
}

function sentRefreshToken(body: string): string | null {
  return new URLSearchParams(body).get("refresh_token");
}

function ok(body: Record<string, unknown>): TokenHttpResponse {
  return { statusCode: 200, body: JSON.stringify(body) };
}

test("(a) access token is cached: two calls, one HTTP request", async () => {
  const { statePath } = tempStatePath();
  const { fn, calls } = fakeRequest(() =>
    ok({ access_token: "access-1", expires_in: 3600 })
  );
  const manager = new TokenManager(options(statePath, { requestFn: fn }));

  assert.equal(await manager.getAccessToken(EXTENDED_SCOPE), "access-1");
  assert.equal(await manager.getAccessToken(EXTENDED_SCOPE), "access-1");
  assert.equal(calls.length, 1);
  assert.equal(sentRefreshToken(calls[0].body), "seed-refresh-token");
});

test("concurrent refreshes are deduplicated into one HTTP request", async () => {
  const { statePath } = tempStatePath();
  let release!: (response: TokenHttpResponse) => void;
  const gate = new Promise<TokenHttpResponse>((resolve) => {
    release = resolve;
  });
  const { fn, calls } = fakeRequest(() => gate);
  const manager = new TokenManager(options(statePath, { requestFn: fn }));

  const first = manager.getAccessToken(EXTENDED_SCOPE);
  const second = manager.getAccessToken(EXTENDED_SCOPE);
  release(ok({ access_token: "shared-access", expires_in: 3600 }));

  assert.equal(await first, "shared-access");
  assert.equal(await second, "shared-access");
  assert.equal(calls.length, 1);
});

test("(b) rotated refresh token is persisted (0600) and used next time", async () => {
  const { statePath } = tempStatePath();
  const { fn, calls } = fakeRequest((_req, _body, call) =>
    call === 1
      ? // expires_in below the 5-minute skew: the next call must refresh.
        ok({ access_token: "access-1", refresh_token: "rotated-1", expires_in: 60 })
      : ok({ access_token: "access-2", refresh_token: "rotated-2", expires_in: 3600 })
  );
  const manager = new TokenManager(options(statePath, { requestFn: fn }));

  assert.equal(await manager.getAccessToken(EXTENDED_SCOPE), "access-1");
  const persisted = JSON.parse(fs.readFileSync(statePath, "utf8"));
  assert.equal(persisted.refreshToken, "rotated-1");
  assert.equal(typeof persisted.updatedAt, "string");
  assert.equal(fs.statSync(statePath).mode & 0o777, 0o600);

  assert.equal(await manager.getAccessToken(EXTENDED_SCOPE), "access-2");
  assert.equal(calls.length, 2);
  assert.equal(sentRefreshToken(calls[1].body), "rotated-1");
  assert.equal(
    JSON.parse(fs.readFileSync(statePath, "utf8")).refreshToken,
    "rotated-2"
  );
});

test("(c) a transient HTTP 500 is retried and then succeeds", async () => {
  const { statePath } = tempStatePath();
  const { sleep, sleeps } = sleepRecorder();
  const { fn, calls } = fakeRequest((_req, _body, call) =>
    call === 1
      ? { statusCode: 500, body: "internal server error" }
      : ok({ access_token: "after-retry", expires_in: 3600 })
  );
  const manager = new TokenManager(options(statePath, { requestFn: fn, sleep }));

  assert.equal(await manager.getAccessToken(EXTENDED_SCOPE), "after-retry");
  assert.equal(calls.length, 2);
  assert.equal(sleeps.length, 1);
  // ~500ms backoff with jitter
  assert.ok(sleeps[0] >= 250 && sleeps[0] <= 1000, `backoff was ${sleeps[0]}ms`);
});

test("a network error is retried and then succeeds", async () => {
  const { statePath } = tempStatePath();
  const { fn, calls } = fakeRequest((_req, _body, call) => {
    if (call === 1) {
      throw new Error("read ECONNRESET");
    }
    return ok({ access_token: "after-network-error", expires_in: 3600 });
  });
  const manager = new TokenManager(options(statePath, { requestFn: fn }));

  assert.equal(await manager.getAccessToken(EXTENDED_SCOPE), "after-network-error");
  assert.equal(calls.length, 2);
});

test("persistent 5xx exhausts retries and does not look like a scope problem", async () => {
  const { statePath } = tempStatePath();
  const { sleep, sleeps } = sleepRecorder();
  const { fn, calls } = fakeRequest(() => ({ statusCode: 503, body: "unavailable" }));
  const manager = new TokenManager(options(statePath, { requestFn: fn, sleep }));

  await assert.rejects(manager.getAccessToken(EXTENDED_SCOPE), (error: unknown) => {
    assert.equal(isScopeConsentError(error), false);
    return true;
  });
  assert.equal(calls.length, 3);
  assert.equal(sleeps.length, 2);
});

test("(d) invalid_grant throws immediately with the re-mint instructions", async () => {
  const { statePath } = tempStatePath();
  const { sleep, sleeps } = sleepRecorder();
  const { fn, calls } = fakeRequest(() => ({
    statusCode: 400,
    body: JSON.stringify({
      error: "invalid_grant",
      error_description:
        "AADSTS70000: The provided value for the input parameter 'refresh_token' is not valid."
    })
  }));
  const manager = new TokenManager(options(statePath, { requestFn: fn, sleep }));

  await assert.rejects(manager.getAccessToken(EXTENDED_SCOPE), (error: unknown) => {
    assert.ok(error instanceof DeadRefreshTokenError);
    assert.match((error as Error).message, /scripts\/get-auth-token\.js/);
    return true;
  });
  assert.equal(calls.length, 1);
  assert.equal(sleeps.length, 0);
});

test("invalid_scope is classified as a scope/consent problem, not retried", async () => {
  const { statePath } = tempStatePath();
  const { fn, calls } = fakeRequest(() => ({
    statusCode: 400,
    body: JSON.stringify({
      error: "invalid_scope",
      error_description:
        "AADSTS70011: The provided value for the input parameter 'scope' is not valid."
    })
  }));
  const manager = new TokenManager(options(statePath, { requestFn: fn }));

  await assert.rejects(manager.getAccessToken(EXTENDED_SCOPE), (error: unknown) => {
    assert.equal(isScopeConsentError(error), true);
    return true;
  });
  assert.equal(calls.length, 1);
});

test("(e) startup prefers the state file over the .env seed", async () => {
  const { statePath } = tempStatePath();
  fs.writeFileSync(
    statePath,
    JSON.stringify({ refreshToken: "state-file-token", updatedAt: "2026-01-01T00:00:00.000Z" })
  );
  const { fn, calls } = fakeRequest(() =>
    ok({ access_token: "from-state", expires_in: 3600 })
  );
  const manager = new TokenManager(options(statePath, { requestFn: fn }));

  assert.equal(await manager.getAccessToken(EXTENDED_SCOPE), "from-state");
  assert.equal(sentRefreshToken(calls[0].body), "state-file-token");
});

test("a corrupt state file falls back to the .env seed without crashing", async () => {
  const { statePath } = tempStatePath();
  fs.writeFileSync(statePath, "this is not json {{{");
  const { fn, calls } = fakeRequest(() =>
    ok({ access_token: "from-seed", expires_in: 3600 })
  );
  const manager = new TokenManager(options(statePath, { requestFn: fn }));

  assert.equal(manager.hasRefreshToken(), true);
  assert.equal(await manager.getAccessToken(EXTENDED_SCOPE), "from-seed");
  assert.equal(sentRefreshToken(calls[0].body), "seed-refresh-token");
});

test("client starts from the state file alone when .env has no refresh token", async () => {
  const { statePath } = tempStatePath();
  fs.writeFileSync(
    statePath,
    JSON.stringify({ refreshToken: "state-only-token", updatedAt: "2026-01-01T00:00:00.000Z" })
  );
  const config: OutlookConfig = {
    clientId: "test-client-id",
    clientSecret: "test-client-secret",
    tenantId: "common",
    refreshToken: undefined,
    tokenStatePath: statePath,
    mailbox: "Inbox",
    batchSize: 20,
    transport: "stdio",
    httpPort: 3000
  };
  const { fn } = fakeRequest(() => ok({ access_token: "x", expires_in: 3600 }));
  assert.doesNotThrow(() => new GraphEmailClient(config, fn));
});

test("client keeps the missing-token error when no source has a token", async () => {
  const { statePath } = tempStatePath();
  const config: OutlookConfig = {
    clientId: "test-client-id",
    clientSecret: "test-client-secret",
    tenantId: "common",
    refreshToken: undefined,
    tokenStatePath: statePath,
    mailbox: "Inbox",
    batchSize: 20,
    transport: "stdio",
    httpPort: 3000
  };
  assert.throws(() => new GraphEmailClient(config), /Missing OUTLOOK_REFRESH_TOKEN/);
});
