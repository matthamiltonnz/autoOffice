// src/taskpane/skills/outlook/index.ts
import message         from './message.md?raw';
import compose         from './compose.md?raw';
import recipients      from './recipients.md?raw';

export const OUTLOOK_SKILL_NAMES = [
  'message', 'compose', 'recipients',
] as const;

export type OutlookSkillName = (typeof OUTLOOK_SKILL_NAMES)[number];

export const OUTLOOK_SKILLS: Record<OutlookSkillName, string> = {
  'message': message,
  'compose': compose,
  'recipients': recipients,
};
