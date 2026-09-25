import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  describeOutlookCapabilities,
  detectOutlookItemMode,
  supportedMailboxVersions,
} from './capabilities.ts';

const READ_ITEM = {
  subject: { getAsync: () => {} },
  displayReplyFormAsync: () => {},
  displayReplyAllFormAsync: () => {},
  displayNewMessageFormAsync: () => {},
};

const COMPOSE_ITEM = {
  subject: { getAsync: () => {}, setAsync: () => {} },
  body: { setAsync: () => {} },
  close: () => {},
};

function stubOffice(item: unknown, versions: readonly string[] = ['1.1', '1.9', '1.10']) {
  vi.stubGlobal('Office', {
    context: {
      mailbox: {
        item,
        diagnostics: { hostName: 'OutlookWebApp', hostVersion: '16.0.17830.4' },
      },
      requirements: {
        isSetSupported: (_name: string, version?: string) =>
          version !== undefined && versions.includes(version),
      },
    },
  });
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('detectOutlookItemMode', () => {
  it('detects read mode from the reply-form APIs', () => {
    stubOffice(READ_ITEM);
    expect(detectOutlookItemMode()).toBe('read');
  });

  it('detects compose mode from a settable subject', () => {
    stubOffice(COMPOSE_ITEM);
    expect(detectOutlookItemMode()).toBe('compose');
  });

  it('reports unknown when there is no Mailbox context', () => {
    vi.stubGlobal('Office', undefined);
    expect(detectOutlookItemMode()).toBe('unknown');
  });
});

describe('supportedMailboxVersions', () => {
  it('lists only the versions the client reports', () => {
    stubOffice(READ_ITEM, ['1.1', '1.2', '1.3']);
    expect(supportedMailboxVersions()).toEqual(['1.1', '1.2', '1.3']);
  });

  it('returns nothing when the Office global is missing', () => {
    vi.stubGlobal('Office', undefined);
    expect(supportedMailboxVersions()).toEqual([]);
  });
});

describe('describeOutlookCapabilities', () => {
  it('names the client, the Mailbox support and the available reply APIs in read mode', () => {
    stubOffice(READ_ITEM);

    const text = describeOutlookCapabilities();

    expect(text).toContain('OutlookWebApp 16.0.17830.4');
    expect(text).toContain('1.1-1.10');
    expect(text).toContain('Current item: READ');
    expect(text).toContain('displayReplyFormAsync=available');
  });

  it('sends the agent to body.setAsync in compose mode', () => {
    stubOffice(COMPOSE_ITEM);

    const text = describeOutlookCapabilities();

    expect(text).toContain('Current item: COMPOSE');
    expect(text).toContain('body.setAsync');
    expect(text).not.toContain('displayReplyFormAsync=available');
  });

  it('documents the "hand the draft to the user" fallback when no reply API exists', () => {
    stubOffice({ subject: { getAsync: () => {} } });

    const text = describeOutlookCapabilities();

    expect(text).toContain('Current item: READ');
    expect(text).toContain('cannot open a reply form');
    expect(text).toContain('displayReplyFormAsync=NOT available');
  });

  it('degrades safely without a Mailbox context', () => {
    vi.stubGlobal('Office', undefined);

    expect(describeOutlookCapabilities()).toContain('capabilities are unavailable');
  });
});
