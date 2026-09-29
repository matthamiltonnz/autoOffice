import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  DEFAULT_OPENWEBUI_BASE_URL,
  clearOpenWebUiCache,
  discoverOpenWebUiModels,
  normalizeOpenWebUiBaseUrl,
} from './openwebui.ts';

const fetchMock = vi.fn();

beforeEach(() => {
  fetchMock.mockReset();
  vi.stubGlobal('fetch', fetchMock);
  clearOpenWebUiCache();
});

afterEach(() => {
  vi.unstubAllGlobals();
});

function jsonResponse(body: unknown, init: { status?: number } = {}): Response {
  return new Response(JSON.stringify(body), {
    status: init.status ?? 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

describe('normalizeOpenWebUiBaseUrl', () => {
  it('falls back to the default base URL when blank', () => {
    expect(normalizeOpenWebUiBaseUrl('')).toBe(DEFAULT_OPENWEBUI_BASE_URL);
    expect(normalizeOpenWebUiBaseUrl('   ')).toBe(DEFAULT_OPENWEBUI_BASE_URL);
  });

  it('strips trailing slashes and surrounding whitespace', () => {
    expect(normalizeOpenWebUiBaseUrl(' http://localhost:8080/api/ ')).toBe(
      'http://localhost:8080/api'
    );
  });

  it('appends /api to a bare origin', () => {
    expect(normalizeOpenWebUiBaseUrl('http://localhost:8080')).toBe('http://localhost:8080/api');
    expect(normalizeOpenWebUiBaseUrl('https://chat.example.com/')).toBe(
      'https://chat.example.com/api'
    );
  });

  it('keeps an explicit path untouched', () => {
    expect(normalizeOpenWebUiBaseUrl('https://example.com/openwebui/api')).toBe(
      'https://example.com/openwebui/api'
    );
  });
});

describe('discoverOpenWebUiModels', () => {
  it('parses the OpenAI-shaped model list and sorts by display name', async () => {
    fetchMock.mockResolvedValue(
      jsonResponse({
        object: 'list',
        data: [
          { id: 'llama3.2', name: 'Llama 3.2', object: 'model' },
          { id: 'gpt-4o-mini', name: 'GPT-4o mini', object: 'model' },
        ],
      })
    );

    const result = await discoverOpenWebUiModels('http://localhost:8080/api');

    expect(result.status).toBe('connected');
    expect(result.models).toEqual([
      { id: 'gpt-4o-mini', displayName: 'GPT-4o mini' },
      { id: 'llama3.2', displayName: 'Llama 3.2' },
    ]);
  });

  it('requests the OpenAI-compatible models endpoint', async () => {
    fetchMock.mockResolvedValue(jsonResponse({ data: [] }));

    await discoverOpenWebUiModels('http://localhost:8080');

    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0][0]).toBe('http://localhost:8080/api/models');
  });

  it('sends the Open WebUI API key as a bearer token', async () => {
    fetchMock.mockResolvedValue(jsonResponse({ data: [] }));

    await discoverOpenWebUiModels('http://localhost:8080/api', { apiKey: 'sk-owui' });

    const init = fetchMock.mock.calls[0][1] as { headers: Record<string, string> };
    expect(init.headers.Authorization).toBe('Bearer sk-owui');
  });

  it('does not send an Authorization header when no key is configured', async () => {
    fetchMock.mockResolvedValue(jsonResponse({ data: [] }));

    await discoverOpenWebUiModels('http://localhost:8080/api');

    const init = fetchMock.mock.calls[0][1] as { headers: Record<string, string> };
    expect(init.headers.Authorization).toBeUndefined();
  });

  it('falls back to the versioned /v1/models route when /models is missing', async () => {
    fetchMock
      .mockResolvedValueOnce(jsonResponse({ detail: 'Not Found' }, { status: 404 }))
      .mockResolvedValueOnce(jsonResponse({ data: [{ id: 'mistral' }] }));

    const result = await discoverOpenWebUiModels('http://localhost:8080/api');

    expect(result.status).toBe('connected');
    expect(result.models).toEqual([{ id: 'mistral' }]);
    expect(fetchMock.mock.calls[1][0]).toBe('http://localhost:8080/api/v1/models');
  });

  it('caches results per base URL and key', async () => {
    fetchMock.mockResolvedValue(jsonResponse({ data: [{ id: 'qwen2.5' }] }));

    await discoverOpenWebUiModels('http://localhost:8080/api', { apiKey: 'sk-owui' });
    await discoverOpenWebUiModels('http://localhost:8080/api', { apiKey: 'sk-owui' });

    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('bypasses the cache when force is set', async () => {
    fetchMock.mockImplementation(() => Promise.resolve(jsonResponse({ data: [{ id: 'qwen2.5' }] })));

    await discoverOpenWebUiModels('http://localhost:8080/api');
    await discoverOpenWebUiModels('http://localhost:8080/api', { force: true });

    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('reports unreachable with the HTTP error when both routes fail', async () => {
    fetchMock.mockResolvedValue(jsonResponse({ detail: 'Unauthorized' }, { status: 401 }));

    const result = await discoverOpenWebUiModels('http://localhost:8080/api');

    expect(result.status).toBe('unreachable');
    expect(result.models).toEqual([]);
    expect(result.error).toContain('HTTP 401');
  });

  it('reports unreachable when the server is not reachable at all', async () => {
    fetchMock.mockRejectedValue(new TypeError('Failed to fetch'));

    const result = await discoverOpenWebUiModels('http://localhost:8080/api');

    expect(result.status).toBe('unreachable');
    expect(result.error).toContain('Network error');
  });

  it('ignores malformed entries in the model list', async () => {
    fetchMock.mockResolvedValue(
      jsonResponse({ data: [{ name: 'no id' }, { id: 'valid' }, { id: 'valid' }] })
    );

    const result = await discoverOpenWebUiModels('http://localhost:8080/api');

    expect(result.models).toEqual([{ id: 'valid' }]);
  });
});

