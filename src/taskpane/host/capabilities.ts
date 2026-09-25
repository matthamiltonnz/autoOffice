// src/taskpane/host/capabilities.ts
/**
 * Runtime capability report for the Outlook host.
 *
 * The agent must not discover what a client supports by probing for it: on a
 * client that does not implement Mailbox 1.9, `item.displayReplyFormAsync` is
 * simply absent, the probe reads that as "API missing", and the run burns turns
 * guessing. So we report the facts once, up front, inside the system prompt.
 */

const MAILBOX_VERSIONS = [
  '1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8', '1.9',
  '1.10', '1.11', '1.12', '1.13', '1.14', '1.15',
] as const;

export type OutlookItemMode = 'read' | 'compose' | 'unknown';

interface MailboxItemLike {
  subject?: { getAsync?: unknown; setAsync?: unknown };
  displayReplyFormAsync?: unknown;
  displayReplyAllFormAsync?: unknown;
  displayNewMessageFormAsync?: unknown;
  close?: unknown;
}

const office = (): typeof Office | undefined =>
  typeof Office === 'undefined' ? undefined : Office;

function mailboxItem(): MailboxItemLike | undefined {
  return office()?.context?.mailbox?.item as MailboxItemLike | undefined;
}

const isFn = (value: unknown): boolean => typeof value === 'function';

/**
 * The Mailbox APIs that write a draft are read-mode only, and the subject on a
 * compose item is settable, so the presence of those members identifies the
 * mode even on clients where the reply-form APIs are missing.
 */
export function detectOutlookItemMode(): OutlookItemMode {
  const item = mailboxItem();
  if (!item) return 'unknown';
  if (isFn(item.subject?.setAsync)) return 'compose';
  if (isFn(item.displayReplyFormAsync) || isFn(item.displayReplyAllFormAsync)) {
    return 'read';
  }
  if (isFn(item.close)) return 'compose';
  // Nothing on the item is writable and there is no close(): this is a read
  // item, even if the client is too old to expose the reply-form APIs.
  return 'read';
}

export function supportedMailboxVersions(): string[] {
  const requirements = office()?.context?.requirements;
  if (!requirements?.isSetSupported) return [];
  return MAILBOX_VERSIONS.filter((version) => {
    try {
      return requirements.isSetSupported('Mailbox', version);
    } catch {
      return false;
    }
  });
}

function diagnostics(): string {
  const diagnostics = office()?.context?.mailbox?.diagnostics;
  const name = diagnostics?.hostName;
  const version = diagnostics?.hostVersion;
  if (!name) return 'unknown client';
  return version ? `${name} ${version}` : String(name);
}

export function describeOutlookCapabilities(): string {
  const item = mailboxItem();
  if (!item) {
    return 'Outlook client capabilities are unavailable (no Mailbox context). '
      + 'Check Office.context.requirements.isSetSupported() before using an API.';
  }

  const versions = supportedMailboxVersions();
  const versionsText = versions.length > 0
    ? `${versions[0]}-${versions[versions.length - 1]} (${versions.join(', ')})`
    : 'unknown — call Office.context.requirements.isSetSupported() before using an API and never probe the object model';
  const mode = detectOutlookItemMode();

  const lines = [
    `Client: ${diagnostics()}`,
    `Mailbox requirement sets supported: ${versionsText}`,
    mode === 'compose'
      ? 'Current item: COMPOSE — a draft is already open. The reply-form APIs do not exist in this mode; write into the open message (body.setAsync for the text, subject.setAsync if a subject is needed).'
      : mode === 'read'
        ? `Current item: READ — an email is open and it is read-only. The reply-form APIs are the ones to use: displayReplyFormAsync=${isFn(item.displayReplyFormAsync) ? 'available' : 'NOT available'}, displayReplyAllFormAsync=${isFn(item.displayReplyAllFormAsync) ? 'available' : 'NOT available'}, displayNewMessageFormAsync=${isFn(item.displayNewMessageFormAsync) ? 'available' : 'NOT available'}.`
        : 'Current item: mode could not be determined — do not call a reply-form API without checking typeof first, and never guess other API names.',
  ];

  if (mode !== 'compose' && !isFn(item.displayReplyFormAsync)
    && !isFn(item.displayReplyAllFormAsync) && !isFn(item.displayNewMessageFormAsync)) {
    lines.push(
      'This client cannot open a reply form. Put the finished draft in the chat, tell the user to click Reply and paste it, and say that the client does not support compose-for-me.',
    );
  }

  return lines.join('\n');
}
