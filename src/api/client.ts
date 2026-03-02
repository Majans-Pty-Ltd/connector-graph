import {
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET,
  GRAPH_API_BASE,
} from "../utils/config.js";
import { log, logError } from "../utils/logger.js";
import type { ODataResponse } from "./types.js";

const MAX_RETRIES = 2;

interface TokenCache {
  token: string;
  expiresAt: number;
}

let tokenCache: TokenCache | null = null;

async function getToken(): Promise<string> {
  const now = Date.now();
  if (tokenCache && tokenCache.expiresAt > now + 60_000) {
    return tokenCache.token;
  }

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: GRAPH_CLIENT_ID,
    client_secret: GRAPH_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    }
  );

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Token acquisition failed (${res.status}): ${text}`);
  }

  const data = (await res.json()) as { access_token: string; expires_in: number };
  tokenCache = {
    token: data.access_token,
    expiresAt: now + data.expires_in * 1000,
  };
  log("Token acquired");
  return tokenCache.token;
}

export class GraphClient {
  /** Make an authenticated GET request to the Graph API. */
  async get<T>(path: string, params?: Record<string, string>): Promise<T> {
    const url = new URL(path, GRAPH_API_BASE);
    if (params) {
      for (const [k, v] of Object.entries(params)) {
        if (v !== undefined && v !== "") url.searchParams.set(k, v);
      }
    }
    return this.request<T>("GET", url);
  }

  /** Make an authenticated PATCH request. */
  async patch<T>(path: string, body: Record<string, unknown>): Promise<T> {
    const url = new URL(path, GRAPH_API_BASE);
    return this.request<T>("PATCH", url, body);
  }

  /** Make an authenticated POST request. */
  async post<T>(path: string, body: Record<string, unknown>): Promise<T> {
    const url = new URL(path, GRAPH_API_BASE);
    return this.request<T>("POST", url, body);
  }

  /** Make an authenticated DELETE request. */
  async delete(path: string): Promise<void> {
    const url = new URL(path, GRAPH_API_BASE);
    await this.request<void>("DELETE", url);
  }

  /** Fetch all pages of a paginated OData response. */
  async getAll<T>(path: string, params?: Record<string, string>): Promise<T[]> {
    const all: T[] = [];
    let url = new URL(path, GRAPH_API_BASE);
    if (params) {
      for (const [k, v] of Object.entries(params)) {
        if (v !== undefined && v !== "") url.searchParams.set(k, v);
      }
    }

    let nextLink: string | undefined;
    do {
      const fetchUrl = nextLink ? new URL(nextLink) : url;
      const result = await this.request<ODataResponse<T>>("GET", fetchUrl);
      all.push(...(result.value ?? []));
      nextLink = result["@odata.nextLink"];
    } while (nextLink && all.length < 5000); // Safety cap

    return all;
  }

  private async request<T>(
    method: string,
    url: URL,
    body?: Record<string, unknown>
  ): Promise<T> {
    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      const token = await getToken();

      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      };
      const init: RequestInit = { method, headers };

      if (body) {
        init.body = JSON.stringify(body);
      }

      log(`${method} ${url.pathname}${url.search}`);
      const res = await fetch(url.toString(), init);

      if (res.status === 429) {
        const retryAfter = parseInt(res.headers.get("Retry-After") ?? "10", 10);
        logError(`Rate limited, waiting ${retryAfter}s`);
        await new Promise((r) => setTimeout(r, retryAfter * 1000));
        continue;
      }

      if (res.status >= 500 && attempt < MAX_RETRIES) {
        logError(`Server error ${res.status}, retrying in 1s`);
        await new Promise((r) => setTimeout(r, 1000));
        continue;
      }

      // 204 No Content (DELETE responses)
      if (res.status === 204) {
        return undefined as T;
      }

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Graph API error ${res.status}: ${text}`);
      }

      return (await res.json()) as T;
    }

    throw new Error("Max retries exceeded");
  }
}
