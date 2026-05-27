/**
 * Production HTTP client primitives for connector-graph.
 *
 * Wraps `fetch` (via undici under the hood, with our keep-alive Agent already
 * configured in client.ts) to add:
 *
 * - Per-request timeout via AbortController (native fetch has NO default
 *   timeout — a hung Graph endpoint would block forever otherwise)
 * - Exponential backoff + jitter on 429 / 503 / 504 / network errors
 * - Honors `Retry-After` (Graph returns it for throttling)
 * - Per-host circuit breaker — fast-fail when Graph is unhealthy rather
 *   than dragging down all 14 calling agents
 * - Slow-call logging (>5s) so we can find latency outliers
 *
 * This module is intentionally small (~200 lines) and dependency-free — we
 * don't pull in a retry/breaker library; the logic is straightforward enough
 * that adding deps to a deployed service is the worse trade.
 */

import { log, logError } from "../utils/logger.js";

// ── Timeouts ─────────────────────────────────────────────────────────────
// Native fetch has no default timeout. These caps are intentionally generous:
// Graph search endpoints can take 30+ seconds under load. We do NOT want to
// be tighter than Graph's own latency budget.
export const DEFAULT_TIMEOUT_MS = 60_000;       // 60s for normal Graph calls
export const SEARCH_TIMEOUT_MS = 120_000;       // 2m for $search / large pages
export const ATTACHMENT_TIMEOUT_MS = 180_000;   // 3m for attachment downloads

// ── Retry policy ─────────────────────────────────────────────────────────
export const RETRYABLE_STATUSES = new Set([429, 503, 504]);
export const DEFAULT_MAX_RETRIES = 5;
const INITIAL_BACKOFF_MS = 1_000;
const MAX_BACKOFF_MS = 60_000;
const JITTER_MS = 2_000;

// ── Circuit breaker policy ───────────────────────────────────────────────
const BREAKER_FAILURE_THRESHOLD = 5;       // consecutive failures to open
const BREAKER_COOL_DOWN_MS = 60_000;       // stay open this long
const BREAKER_FAILURE_WINDOW_MS = 30_000;  // older failures don't count

export class CircuitOpenError extends Error {
  readonly host: string;
  readonly transient = true;
  constructor(host: string) {
    super(
      `Circuit breaker open for ${host}: downstream is degraded; retry in ~${BREAKER_COOL_DOWN_MS / 1000}s`
    );
    this.name = "CircuitOpenError";
    this.host = host;
  }
}

interface BreakerState {
  failures: number;
  firstFailureAt: number;  // ms epoch of first failure in window (0 if none)
  openedAt: number;        // ms epoch when circuit opened (0 if closed)
}

const breakers = new Map<string, BreakerState>();

function breakerFor(host: string): BreakerState {
  let b = breakers.get(host);
  if (!b) {
    b = { failures: 0, firstFailureAt: 0, openedAt: 0 };
    breakers.set(host, b);
  }
  return b;
}

function recordSuccess(host: string): void {
  const b = breakerFor(host);
  b.failures = 0;
  b.firstFailureAt = 0;
  b.openedAt = 0;
}

function recordFailure(host: string): void {
  const b = breakerFor(host);
  const now = Date.now();
  if (b.firstFailureAt && now - b.firstFailureAt > BREAKER_FAILURE_WINDOW_MS) {
    b.failures = 0;
    b.firstFailureAt = now;
  }
  if (!b.firstFailureAt) {
    b.firstFailureAt = now;
  }
  b.failures += 1;
  if (b.failures >= BREAKER_FAILURE_THRESHOLD) {
    b.openedAt = now;
  }
}

function isOpen(host: string): boolean {
  const b = breakerFor(host);
  if (!b.openedAt) return false;
  if (Date.now() - b.openedAt > BREAKER_COOL_DOWN_MS) {
    // Half-open — let one probe through. A failed probe stays open via
    // recordFailure; a successful one fully resets via recordSuccess.
    return false;
  }
  return true;
}

export function getBreakerState(): Record<
  string,
  { state: "open" | "closed"; failures: number; openedForMs: number }
> {
  const snapshot: Record<
    string,
    { state: "open" | "closed"; failures: number; openedForMs: number }
  > = {};
  const now = Date.now();
  for (const [host, b] of breakers) {
    snapshot[host] = {
      state: isOpen(host) ? "open" : "closed",
      failures: b.failures,
      openedForMs: b.openedAt ? now - b.openedAt : 0,
    };
  }
  return snapshot;
}

// ── Retry helpers ────────────────────────────────────────────────────────

/** Parse Retry-After header. Returns ms, or undefined if absent/invalid. */
function parseRetryAfter(value: string | null): number | undefined {
  if (!value) return undefined;
  const n = parseInt(value, 10);
  if (!Number.isNaN(n) && n >= 0) return n * 1000;
  // HTTP-date form is rare for these APIs — try Date.parse, else give up.
  const ts = Date.parse(value);
  if (!Number.isNaN(ts)) {
    return Math.max(0, ts - Date.now());
  }
  return undefined;
}

function computeBackoffMs(attempt: number, retryAfterMs?: number): number {
  if (retryAfterMs !== undefined) {
    return Math.min(MAX_BACKOFF_MS, retryAfterMs + Math.random() * JITTER_MS);
  }
  const base = INITIAL_BACKOFF_MS * Math.pow(2, attempt - 1);
  return Math.min(MAX_BACKOFF_MS, base + Math.random() * JITTER_MS);
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

function logPath(url: URL): string {
  return `${url.pathname}`;
}

// ── Core request ─────────────────────────────────────────────────────────

export interface HttpRequestOptions {
  method: string;
  headers?: Record<string, string>;
  body?: string;
  /** Per-request total timeout in ms. Default DEFAULT_TIMEOUT_MS. */
  timeoutMs?: number;
  /** Override max retry attempts. Default DEFAULT_MAX_RETRIES. */
  maxRetries?: number;
  /** AbortSignal from the caller — composed with the timeout signal. */
  externalSignal?: AbortSignal;
}

/**
 * Send an HTTP request with retries, Retry-After, circuit breaker, and timing.
 *
 * Returns the final Response. Caller decides how to handle non-2xx that
 * survived our retry policy (e.g., 401, 404, 400 are not retried).
 *
 * Throws:
 *   - CircuitOpenError: when target host's breaker is open
 *   - Error("request timeout"): when our AbortController fires
 *   - Other fetch errors when all retries exhausted
 */
export async function httpRequest(
  url: URL,
  opts: HttpRequestOptions
): Promise<Response> {
  const host = url.host;
  if (isOpen(host)) {
    log(`Circuit open for ${host} — failing fast`);
    throw new CircuitOpenError(host);
  }

  const timeoutMs = opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const maxRetries = opts.maxRetries ?? DEFAULT_MAX_RETRIES;
  const started = Date.now();
  let lastError: unknown;
  let response: Response | null = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    // Per-attempt AbortController — keeps each attempt under timeoutMs even
    // if previous attempts ate part of the wall-clock budget.
    const ctl = new AbortController();
    const timer = setTimeout(() => ctl.abort(), timeoutMs);
    if (opts.externalSignal) {
      if (opts.externalSignal.aborted) {
        ctl.abort();
      } else {
        opts.externalSignal.addEventListener("abort", () => ctl.abort(), {
          once: true,
        });
      }
    }

    try {
      response = await fetch(url.toString(), {
        method: opts.method,
        headers: opts.headers,
        body: opts.body,
        signal: ctl.signal,
      });

      // Retryable status?
      if (RETRYABLE_STATUSES.has(response.status) && attempt < maxRetries) {
        const retryAfter = parseRetryAfter(response.headers.get("retry-after"));
        const wait = computeBackoffMs(attempt, retryAfter);
        log(
          `${opts.method} ${logPath(url)} -> ${response.status} (retry ${attempt}/${maxRetries} in ${Math.round(wait)}ms${retryAfter !== undefined ? `, Retry-After=${retryAfter}ms` : ""})`
        );
        // Drain the body so the connection can be reused.
        await response.text().catch(() => undefined);
        await sleep(wait);
        continue;
      }

      // Done — success or non-retryable error
      const duration = Date.now() - started;
      if (response.ok) {
        recordSuccess(host);
        if (duration > 5_000) {
          log(
            `SLOW ${opts.method} ${logPath(url)} -> ${response.status} in ${duration}ms (attempt ${attempt})`
          );
        }
      } else {
        if (response.status >= 500) {
          recordFailure(host);
        }
        log(`${opts.method} ${logPath(url)} -> ${response.status} in ${duration}ms`);
      }
      return response;
    } catch (err) {
      lastError = err;
      // Distinguish abort-from-timeout vs other transport errors. Both are
      // retryable; we just want clear logging.
      const isAbort =
        (err as { name?: string })?.name === "AbortError" || ctl.signal.aborted;
      const errLabel = isAbort ? "timeout" : (err as Error).message || "transport-error";

      if (attempt < maxRetries) {
        const wait = computeBackoffMs(attempt);
        log(
          `${opts.method} ${logPath(url)} (${errLabel}) — retry ${attempt}/${maxRetries} in ${Math.round(wait)}ms`
        );
        await sleep(wait);
        continue;
      }

      // Out of retries — count toward breaker
      recordFailure(host);
      const duration = Date.now() - started;
      logError(
        `${opts.method} ${logPath(url)} FAILED after ${attempt} attempts in ${duration}ms`,
        err
      );
      throw err;
    } finally {
      clearTimeout(timer);
    }
  }

  // Exhausted retries on retryable statuses — return last response so caller
  // can decide what to surface.
  if (response) {
    recordFailure(host);
    return response;
  }
  // Shouldn't be reachable, but for type safety:
  throw lastError instanceof Error ? lastError : new Error("Max retries exceeded");
}
