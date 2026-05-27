// Smoke test for vault Bearer token validation (connector-graph).
//
// Combines two test approaches:
//   1. Real Graph-MCP-User token from local MSAL cache — exercises the live
//      claim shape end-to-end.
//   2. Synthetic JWTs with tampered claims — covers issuer/audience/appid/exp
//      rejection paths without needing live JWKS.
//
// Run after `npm run build`:
//   GRAPH_TENANT_ID=d54794b1-f598-4c0f-a276-6039a39774ac \
//     GRAPH_API_KEY=server-api-key \
//     node test-vault-auth.mjs
//
// Both env vars must be set BEFORE node starts — config.ts reads them at
// module-load and the X-API-Key precedence cases need GRAPH_API_KEY populated.
//
// Requires a fresh graph token cache (run `python get-user-token.py` first).

import { readFileSync, existsSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { SignJWT, generateKeyPair } from "jose";
import { validateVaultBearerToken, looksLikeJwt } from "./dist/utils/jwt-validator.js";
import { authenticateRequest } from "./dist/utils/auth.js";

const TENANT = "d54794b1-f598-4c0f-a276-6039a39774ac";
const ALLOWED_APP = "02fa0ea1-4b30-4bd9-9c4a-483f97d63b21";
const OTHER_APP = "11111111-2222-3333-4444-555555555555";
const GRAPH_AUD = "00000003-0000-0000-c000-000000000000";

const cases = [];
function record(label, ok) {
  console.log(`  [${ok ? "PASS" : "FAIL"}] ${label}`);
  cases.push(ok);
}

// ── Real token from MSAL cache ─────────────────────────────────────────────
const cacheFile = join(homedir(), ".connector-graph", "token_cache.bin");
if (!existsSync(cacheFile)) {
  console.error(`No token cache at ${cacheFile} — run get-user-token.py first`);
  process.exit(1);
}
const cache = JSON.parse(readFileSync(cacheFile, "utf8"));
const tokenEntry = Object.values(cache.AccessToken ?? {}).find(
  (t) => t?.secret?.split(".").length === 3
);
if (!tokenEntry) {
  console.error("No JWT-shaped access token in cache");
  process.exit(1);
}
const realToken = tokenEntry.secret;

console.log("Token preview:", realToken.slice(0, 60) + "...");
console.log("looksLikeJwt:", looksLikeJwt(realToken));
console.log();
console.log("Test cases:");

try {
  const result = await validateVaultBearerToken(realToken);
  record(`real Graph-MCP-User token accepted (appid=${result.appid}, upn=${result.upn})`, true);
} catch (err) {
  record(`real Graph-MCP-User token REJECTED: ${err.message}`, false);
}

// ── Synthetic JWTs ──────────────────────────────────────────────────────────
// We sign with a throwaway key — the validator skips signature verification
// for Graph tokens (Microsoft's nonce-hashing prevents third-party signature
// validation), so the test focuses on claim checks.
const { privateKey } = await generateKeyPair("RS256");
const issuerV1 = `https://sts.windows.net/${TENANT}/`;
const issuerV2 = `https://login.microsoftonline.com/${TENANT}/v2.0`;

async function mintToken({ aud, iss, appid, expSec }) {
  return await new SignJWT({
    appid,
    azp: appid,
    oid: "user-oid-12345",
    upn: "user@majans.com",
  })
    .setProtectedHeader({ alg: "RS256" })
    .setIssuer(iss)
    .setAudience(aud)
    .setIssuedAt()
    .setNotBefore("0s")
    .setExpirationTime(expSec ?? "1h")
    .sign(privateKey);
}

// 2a. valid v1 token (allowed app)
try {
  const t = await mintToken({ aud: GRAPH_AUD, iss: issuerV1, appid: ALLOWED_APP });
  await validateVaultBearerToken(t);
  record("synthetic v1 token (allowed app) accepted", true);
} catch (err) {
  record(`synthetic v1 token REJECTED: ${err.message}`, false);
}

// 2b. valid v2 token (allowed app)
try {
  const t = await mintToken({ aud: GRAPH_AUD, iss: issuerV2, appid: ALLOWED_APP });
  await validateVaultBearerToken(t);
  record("synthetic v2 token (allowed app) accepted", true);
} catch (err) {
  record(`synthetic v2 token REJECTED: ${err.message}`, false);
}

// 2c. wrong appid
try {
  const t = await mintToken({ aud: GRAPH_AUD, iss: issuerV1, appid: OTHER_APP });
  await validateVaultBearerToken(t);
  record("wrong appid INCORRECTLY ACCEPTED", false);
} catch (err) {
  record(`wrong appid rejected: ${err.message}`, true);
}

// 2d. wrong audience (D365 instead of Graph)
try {
  const t = await mintToken({
    aud: "https://majans.operations.dynamics.com",
    iss: issuerV1,
    appid: ALLOWED_APP,
  });
  await validateVaultBearerToken(t);
  record("wrong audience INCORRECTLY ACCEPTED", false);
} catch (err) {
  record(`wrong audience rejected: ${err.message}`, true);
}

// 2e. wrong tenant
try {
  const t = await mintToken({
    aud: GRAPH_AUD,
    iss: "https://sts.windows.net/00000000-0000-0000-0000-000000000000/",
    appid: ALLOWED_APP,
  });
  await validateVaultBearerToken(t);
  record("wrong tenant INCORRECTLY ACCEPTED", false);
} catch (err) {
  record(`wrong tenant rejected: ${err.message}`, true);
}

// 2f. expired token
try {
  const t = await mintToken({
    aud: GRAPH_AUD,
    iss: issuerV1,
    appid: ALLOWED_APP,
    expSec: Math.floor(Date.now() / 1000) - 60,
  });
  await validateVaultBearerToken(t);
  record("expired token INCORRECTLY ACCEPTED", false);
} catch (err) {
  record(`expired token rejected: ${err.message}`, true);
}

// 2g. garbage / non-JWT
try {
  await validateVaultBearerToken("not.a.token");
  record("garbage INCORRECTLY ACCEPTED", false);
} catch (err) {
  record(`garbage rejected: ${err.message.slice(0, 80)}`, true);
}

// ── authenticateRequest precedence checks ──────────────────────────────────
// GRAPH_API_KEY must be set in the parent shell environment for these to work
// (see usage notes at top of file).
const SERVER_KEY = process.env.GRAPH_API_KEY ?? "server-api-key";
const goodBearer = await mintToken({ aud: GRAPH_AUD, iss: issuerV1, appid: ALLOWED_APP });
const badBearer = await mintToken({ aud: GRAPH_AUD, iss: issuerV1, appid: OTHER_APP });

function fakeReq(headers) {
  return { headers };
}

// X-API-Key wins when correct (good API key + good Bearer -> SP path)
{
  const r = await authenticateRequest(
    fakeReq({ "x-api-key": SERVER_KEY, authorization: `Bearer ${goodBearer}` })
  );
  record(
    `X-API-Key precedence: good key + good Bearer -> SP path (no userToken): allowed=${r.allowed}, userToken=${r.userToken}`,
    r.allowed && r.userToken === undefined
  );
}

// Bad X-API-Key + good Bearer -> rejected (no fallback)
{
  const r = await authenticateRequest(
    fakeReq({ "x-api-key": "wrong", authorization: `Bearer ${goodBearer}` })
  );
  record(
    `X-API-Key precedence: bad key + good Bearer -> REJECT (no fallback): ${r.error}`,
    !r.allowed
  );
}

// Bearer-only good
{
  const r = await authenticateRequest(fakeReq({ authorization: `Bearer ${goodBearer}` }));
  record(
    `Bearer-only good token -> user path (userToken set): allowed=${r.allowed}, userToken set=${r.userToken === goodBearer}`,
    r.allowed && r.userToken === goodBearer
  );
}

// Bearer-only bad appid
{
  const r = await authenticateRequest(fakeReq({ authorization: `Bearer ${badBearer}` }));
  record(`Bearer-only bad appid -> REJECT: ${r.error}`, !r.allowed);
}

// X-API-Key only (regression — agent path)
{
  const r = await authenticateRequest(fakeReq({ "x-api-key": SERVER_KEY }));
  record(`X-API-Key only (regression): allowed=${r.allowed}`, r.allowed && r.userToken === undefined);
}

const passed = cases.filter(Boolean).length;
const total = cases.length;
console.log(`\nResult: ${passed}/${total} passed`);
process.exit(passed === total ? 0 : 1);
