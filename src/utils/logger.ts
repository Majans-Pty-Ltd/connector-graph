/** Log to stderr so stdout stays clean for MCP stdio transport. */
export function log(message: string, ...args: unknown[]): void {
  console.error(`[GRAPH-MCP] ${message}`, ...args);
}

export function logError(message: string, error?: unknown): void {
  const detail = error instanceof Error ? error.message : String(error ?? "");
  console.error(`[GRAPH-MCP ERROR] ${message}${detail ? `: ${detail}` : ""}`);
}
