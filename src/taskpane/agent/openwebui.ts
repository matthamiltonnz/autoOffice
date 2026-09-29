/**
 * Open WebUI model discovery over its OpenAI-compatible API.
 *
 * Open WebUI mounts an OpenAI-compatible surface under `<host>/api`:
 *   1. `GET <baseUrl>/models`     — OpenAI-shaped list (`{ data: [{ id, name }] }`).
 *   2. `GET <baseUrl>/v1/models`  — fallback for builds/proxies that only expose the versioned route.
 *
 * `baseUrl` is the OpenAI-compatible base (typically `http://localhost:8080/api`).
 * When the instance has auth enabled, the user's Open WebUI API key is sent as
 * `Authorization: Bearer <key>`.
 *
 * Discovery never throws: transport problems are reported through
 * `OpenWebUiDiscovery.status`/`error` so the settings panel can show a hint
 * without an error boundary.
 */

export const DEFAULT_OPENWEBUI_BASE_URL = 'http://localhost:8080/api';

export interface OpenWebUiModel {
  id: string;
  displayName?: string;
  ownedBy?: string;
}

export type OpenWebUiStatus = 'idle' | 'connecting' | 'connected' | 'unreachable';

export interface OpenWebUiDiscovery {
  status: OpenWebUiStatus;
  models: OpenWebUiModel[];
  error?: string;
}

export interface OpenWebUiDiscoveryOptions {
  signal?: AbortSignal;
  timeoutMs?: number;
  force?: boolean;
  apiKey?: string;
}

const DEFAULT_TIMEOUT_MS = 2500;
const CACHE_TTL_MS = 30_000;

const cache = new Map<string, { result: OpenWebUiDiscovery; expiresAt: number }>();

/**
 * Normalises a user-entered Open WebUI base URL.
 *
 * - blank input falls back to {@link DEFAULT_OPENWEBUI_BASE_URL}
 * - trailing slashes are removed
 * - a bare origin (`http://localhost:8080`) gets the `/api` suffix appended,
 *   because that is where Open WebUI mounts its OpenAI-compatible routes
 */
export function normalizeOpenWebUiBaseUrl(raw: string): string {
  const trimmed = (raw || '').trim().replace(/\/+$/, '');
  if (!trimmed) return DEFAULT_OPENWEBUI_BASE_URL;
  try {
    const url = new URL(trimmed);
    if (!url.pathname || url.pathname === '/') {
      return `${trimmed}/api`;
    }
  } catch {
    // Not a parseable absolute URL — leave the value untouched.
  }
  return trimmed;
}

export function clearOpenWebUiCache(): void {
  cache.clear();
}

export async function discoverOpenWebUiModels(
  rawBaseUrl: string,
  opts: OpenWebUiDiscoveryOptions = {},
): Promise<OpenWebUiDiscovery> {
  const baseUrl = normalizeOpenWebUiBaseUrl(rawBaseUrl);
  const apiKey = (opts.apiKey ?? '').trim();
  const cacheKey = `${baseUrl}::${apiKey}`;
  const now = Date.now();

  if (!opts.force) {
    const hit = cache.get(cacheKey);
    if (hit && hit.expiresAt > now) return hit.result;
  }

  const timeoutMs = opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;

  // Primary route: `/api/models` (OpenAI-compatible).
  const primary = await tryFetch(`${baseUrl}/models`, opts.signal, timeoutMs, apiKey);
  if (primary.ok && primary.body) {
    return remember(cacheKey, now, parseModels(primary.body));
  }

  // Fallback route: `/api/v1/models` on builds that version the API.
  const fallback = await tryFetch(`${baseUrl}/v1/models`, opts.signal, timeoutMs, apiKey);
  if (fallback.ok && fallback.body) {
    return remember(cacheKey, now, parseModels(fallback.body));
  }

  const result: OpenWebUiDiscovery = {
    status: 'unreachable',
    models: [],
    error: fallback.error ?? primary.error ?? 'Unknown error',
  };
  cache.set(cacheKey, { result, expiresAt: now + CACHE_TTL_MS });
  return result;
}

function remember(cacheKey: string, now: number, models: OpenWebUiModel[]): OpenWebUiDiscovery {
  const result: OpenWebUiDiscovery = {
    status: 'connected',
    models: sortModels(models),
  };
  cache.set(cacheKey, { result, expiresAt: now + CACHE_TTL_MS });
  return result;
}

interface FetchAttempt {
  ok: boolean;
  body?: unknown;
  error?: string;
}

async function tryFetch(
  url: string,
  externalSignal: AbortSignal | undefined,
  timeoutMs: number,
  apiKey: string,
): Promise<FetchAttempt> {
  const ctrl = new AbortController();
  const onAbort = () => ctrl.abort(externalSignal?.reason);
  if (externalSignal) {
    if (externalSignal.aborted) {
      throw new DOMException('Aborted', 'AbortError');
    }
    externalSignal.addEventListener('abort', onAbort, { once: true });
  }
  const timer = setTimeout(() => ctrl.abort(new DOMException('Timeout', 'TimeoutError')), timeoutMs);
  try {
    const headers: Record<string, string> = { Accept: 'application/json' };
    if (apiKey) headers.Authorization = `Bearer ${apiKey}`;
    const response = await fetch(url, {
      method: 'GET',
      headers,
      signal: ctrl.signal,
    });
    if (!response.ok) {
      return { ok: false, error: `HTTP ${response.status} at ${url}` };
    }
    const body = await response.json();
    return { ok: true, body };
  } catch (err) {
    if (ctrl.signal.aborted && !externalSignal?.aborted) {
      // Internal timeout — treat as unreachable.
      return { ok: false, error: `Timeout after ${timeoutMs}ms at ${url}` };
    }
    if (err instanceof TypeError) {
      // CORS / DNS / connection refused all surface as `TypeError: Failed to fetch`.
      return { ok: false, error: `Network error at ${url}: ${err.message}` };
    }
    return { ok: false, error: `Fetch failed at ${url}: ${(err as Error).message}` };
  } finally {
    clearTimeout(timer);
    externalSignal?.removeEventListener('abort', onAbort);
  }
}



/** Accepts `{ data: [...] }` (OpenAI), `{ models: [...] }` and bare arrays. */
function extractModels(body: unknown): unknown[] {
  if (Array.isArray(body)) return body;
  if (body && typeof body === 'object') {
    const record = body as Record<string, unknown>;
    if (Array.isArray(record.data)) return record.data;
    if (Array.isArray(record.models)) return record.models;
  }
  return [];
}

function parseModels(body: unknown): OpenWebUiModel[] {
  const models: OpenWebUiModel[] = [];
  const seen = new Set<string>();
  for (const entry of extractModels(body)) {
    if (!entry || typeof entry !== 'object') continue;
    const record = entry as Record<string, unknown>;
    const id = typeof record.id === 'string' ? record.id.trim() : '';
    if (!id || seen.has(id)) continue;
    seen.add(id);
    const name = typeof record.name === 'string' ? record.name.trim() : '';
    const ownedBy = typeof record.owned_by === 'string' ? record.owned_by : '';
    models.push({
      id,
      ...(name && name !== id ? { displayName: name } : {}),
      ...(ownedBy ? { ownedBy } : {}),
    });
  }
  return models;
}

function sortModels(models: OpenWebUiModel[]): OpenWebUiModel[] {
  return [...models].sort((a, b) =>
    (a.displayName ?? a.id).localeCompare(b.displayName ?? b.id)
  );
}
