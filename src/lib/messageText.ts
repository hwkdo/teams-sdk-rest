import type { Activity } from '@microsoft/teams.api';

export function normalizeActivityText(activity: Activity): string {
  const rawText = 'text' in activity ? activity.text : undefined;

  if (typeof rawText !== 'string') {
    return '';
  }

  return rawText
    .replace(/<at>.*?<\/at>\s*/gi, '')
    .replace(/<[^>]+>/g, '')
    .trim();
}

export function isHiCommand(text: string): boolean {
  const normalized = text
    .trim()
    .toLowerCase()
    .replace(/^\/+/, '')
    .replace(/[!?.]+$/g, '');

  return normalized === 'hi' || normalized === 'hello' || normalized === 'hallo';
}

export function isFromBot(activity: Activity, botClientId: string): boolean {
  const fromId = activity.from?.id;

  if (!fromId || !botClientId) {
    return false;
  }

  if (fromId === botClientId) {
    return true;
  }

  return fromId.includes(botClientId);
}
