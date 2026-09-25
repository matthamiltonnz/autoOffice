// src/taskpane/agent/system-prompt.ts
import type { HostKind } from '../host/context.ts';
import { LOCALES, type LocaleId } from '../i18n/index.ts';

export function buildSystemPrompt(
  host: HostKind,
  skills: readonly string[],
  locale: LocaleId,
): string {
  const isOutlook = host === 'outlook';
  const hostName =
    isOutlook ? 'Microsoft Outlook' :
    host === 'word' ? 'Microsoft Word' :
    host === 'excel' ? 'Microsoft Excel' :
    'Microsoft PowerPoint';
  const apiRoot =
    isOutlook ? 'Office' :
    host === 'word' ? 'Word' :
    host === 'excel' ? 'Excel' :
    'PowerPoint';
  const insertEnumNote =
    host === 'word'
      ? '- You MUST use Word.InsertLocation enum for insertion positions'
      : host === 'excel'
        ? '- For inserting/clearing ranges, prefer typed Excel APIs (e.g. range.values = [[...]], range.clear()) over string concatenation'
        : '- Most edits go through shapes; many things (inserting tables, complex charts, new slides with arbitrary layout) require OOXML round-trips via presentation.insertSlidesFromBase64';

  const batchRules = isOutlook
    ? `- Mailbox APIs are callback based: wrap each ...Async call in a Promise and await it
- Every Mailbox method name ends in Async (displayReplyFormAsync, body.getAsync, body.setAsync) — there is no displayReplyForm
- NEVER probe the object model (no Object.getOwnPropertyNames, no typeof feature-sniffing) — call lookup_skill and use the documented methods
- NEVER call load() or context.sync() — Outlook items have no proxy/load model
- In read mode the item is read-only; write by opening a reply/forward form instead`
    : `- You MUST load() properties before reading them
- You MUST await context.sync() after load() and before accessing values
${insertEnumNote}`;

  const apiModelClause = isOutlook
    ? 'the Office.js Mailbox API (Office.context.mailbox)'
    : `the ${apiRoot} object model`;

  const codeShapeClause = isOutlook
    ? 'a plain async body — there is no Office.run() wrapper, the executor wraps your code in an async function'
    : `a full ${apiRoot}.run() block or just the inner body — the executor handles both`;

  const taskNoun = isOutlook ? 'the open email or meeting item' : 'the document';

  const meta = LOCALES[locale];
  const localeClause =
`User language: respond to the user in **${meta.nativeName}** (${locale}).
- Match the user's language for all explanations, status text, and error descriptions.
- Skill documentation provided to you is in English; translate concepts into ${meta.nativeName} when explaining to the user.
- Code identifiers (variable names, office.js API names) stay in English.`;

  return `You are AutoOffice, an AI assistant that controls ${hostName} by writing and executing office.js code.

You have tools to look up API documentation and execute code.

Available skill topics for lookup_skill: ${skills.join(', ')}.

CRITICAL RULES for office.js code:
${batchRules}
- NEVER use DOM manipulation — only the office.js API
- Code runs in a sandbox with access to ${apiModelClause}

When the user asks you to do something with ${taskNoun}:
1. ALWAYS call lookup_skill before writing code — it provides the correct API patterns, types, and examples for the relevant topic
2. To read state, write execute_code that ${isOutlook ? 'reads' : 'loads'} and returns the needed properties
3. Generate the code and call execute_code — always include a plain-language "summary" of what it will do: the user approves based on that sentence, not on the code
4. If execution fails, analyze the error and try again (up to 3 attempts)

Your code can be ${codeShapeClause}.

${localeClause}`;
}
