import { describe, it, expect } from 'vitest';
import { buildSystemPrompt } from './system-prompt.ts';

describe('buildSystemPrompt', () => {
  it('contains an English-locale clause naming "English"', () => {
    const p = buildSystemPrompt('word', ['document'], 'en');
    expect(p).toMatch(/respond to the user in \*\*English\*\* \(en\)/i);
  });

  it('contains a Hebrew-locale clause naming the native name', () => {
    const p = buildSystemPrompt('word', ['document'], 'he');
    expect(p).toMatch(/respond to the user in \*\*עברית\*\* \(he\)/i);
  });

  it('keeps locale clause near the end of the prompt', () => {
    const p = buildSystemPrompt('word', ['document'], 'he');
    const idx = p.toLowerCase().indexOf('respond to the user');
    expect(idx).toBeGreaterThan(p.length / 2);
  });

  it('still includes the office.js critical rules', () => {
    const p = buildSystemPrompt('word', ['document'], 'en');
    expect(p).toContain('CRITICAL RULES for office.js code');
  });
});

describe('buildSystemPrompt — Outlook host', () => {
  const prompt = buildSystemPrompt('outlook', ['message', 'compose'], 'en');

  it('names Microsoft Outlook and the Mailbox API', () => {
    expect(prompt).toContain('Microsoft Outlook');
    expect(prompt).toContain('Office.js Mailbox API');
  });

  it('does not apply the Word/Excel load() + sync() batching rules', () => {
    expect(prompt).not.toContain('You MUST load() properties before reading them');
    expect(prompt).not.toContain('You MUST await context.sync()');
    expect(prompt).toContain('NEVER call load() or context.sync()');
  });

  it('explains that Outlook has no Office.run() wrapper', () => {
    expect(prompt).toContain('there is no Office.run() wrapper');
  });

  it('lists the Outlook skill topics passed in', () => {
    expect(prompt).toContain('message, compose');
  });
});

