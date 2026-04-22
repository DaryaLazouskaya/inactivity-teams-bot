'use strict';

import { App } from '@microsoft/teams.apps';

interface MentionRef {
  id: string;
  name: string;
  text: string; // e.g. "<at>Anna</at>"
}

interface TrackedItem {
  chatId: string;
  channelId: string;
  serviceUrl: string;
  tenantId?: string;
  reminderMs: number;
  agreedText: string;
  mentions: MentionRef[];
  lastMessageAt: string;
  lastNotifiedAt: string | null;
  isActive: boolean;
}

interface ParsedTrackResult {
  reminderMs: number;
  agreedText: string;
  usedDefault: boolean;
}

const botApp = new App({});

const DEFAULT_REMINDER_MS = 2 * 24 * 60 * 60 * 1000; // 2 days
const CHECK_EVERY_MS = 1000; // use 1s for testing, increase later

const MIN_REMINDER_MS = 5 * 1000; // 5 seconds
const MAX_REMINDER_MS = 30 * 24 * 60 * 60 * 1000; // 30 days
const MAX_AGREED_TEXT_LENGTH = 2000;

const tracked = new Map<string, TrackedItem>();

function nowIso(): string {
  return new Date().toISOString();
}

function formatDuration(ms: number): string {
  if (ms >= 24 * 60 * 60 * 1000) {
    const days = Math.round(ms / (24 * 60 * 60 * 1000));
    return `${days} day${days === 1 ? '' : 's'}`;
  }
  if (ms >= 60 * 60 * 1000) {
    const hours = Math.round(ms / (60 * 60 * 1000));
    return `${hours} hour${hours === 1 ? '' : 's'}`;
  }
  if (ms >= 60 * 1000) {
    const mins = Math.round(ms / (60 * 1000));
    return `${mins} minute${mins === 1 ? '' : 's'}`;
  }
  const secs = Math.round(ms / 1000);
  return `${secs} second${secs === 1 ? '' : 's'}`;
}

function isOlderThan(dateIso: string, thresholdMs: number): boolean {
  if (!dateIso) return false;

  const timestamp = new Date(dateIso).getTime();
  if (!Number.isFinite(timestamp)) return false;

  return Date.now() - timestamp >= thresholdMs;
}

function sanitizeAgreedText(value: unknown): string {
  if (typeof value !== 'string') return '';

  return value
    .replace(/\r/g, '')
    .replace(/\0/g, '')
    .trim()
    .slice(0, MAX_AGREED_TEXT_LENGTH);
}

function normalizeBoolean(value: unknown): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') return value.trim().toLowerCase() === 'true';
  if (typeof value === 'number') return value !== 0;
  return false;
}

function isValidReminderMs(value: unknown): value is number {
  return (
    typeof value === 'number' &&
    Number.isFinite(value) &&
    value >= MIN_REMINDER_MS &&
    value <= MAX_REMINDER_MS
  );
}

function normalizeIncomingTrackText(text: string): string {
  return text.replace(/<at>(.*?)<\/at>/gi, '@$1');
}

function extractMentions(activity: any): MentionRef[] {
  const entities = Array.isArray(activity?.entities) ? activity.entities : [];
  const seen = new Set<string>();

  return entities
    .filter((entity: any) => entity?.type === 'mention' && entity?.mentioned?.id)
    .map((entity: any) => {
      const name = String(entity?.mentioned?.name || entity?.text || '')
        .replace(/<\/?at>/gi, '')
        .trim();

      return {
        id: String(entity.mentioned.id),
        name,
        text: `<at>${name}</at>`
      };
    })
    .filter((mention: MentionRef) => {
      if (!mention.id || !mention.name) return false;
      if (seen.has(mention.id)) return false;

      seen.add(mention.id);
      return true;
    });
}

function applyMentionsToText(text: string, mentions: MentionRef[]): string {
  let result = text;

  for (const mention of mentions) {
    result = result.split(`<at>${mention.name}</at>`).join(mention.text);
    result = result.split(`@${mention.name}`).join(mention.text);
  }

  return result;
}

function buildMentionEntities(mentions: MentionRef[]) {
  return mentions.map((mention) => ({
    type: 'mention',
    text: mention.text,
    mentioned: {
      id: mention.id,
      name: mention.name
    }
  }));
}

function startTracking(activity: any, reminderMs: number, agreedText: string): void {
  const chatId = activity.conversation.id;
  const mentions = extractMentions(activity);

  tracked.set(chatId, {
    chatId,
    channelId: activity.channelId,
    serviceUrl: activity.serviceUrl,
    tenantId: activity.conversation?.tenantId ?? activity.channelData?.tenant?.id,
    reminderMs,
    agreedText,
    mentions,
    lastMessageAt: nowIso(),
    lastNotifiedAt: null,
    isActive: true
  });
}

function stopTracking(chatId: string): boolean {
  const item = tracked.get(chatId);
  if (!item || !item.isActive) return false;

  item.isActive = false;
  tracked.set(chatId, item);
  return true;
}

function refreshLastMessage(chatId: string): void {
  const item = tracked.get(chatId);
  if (!item || !item.isActive) return;

  item.lastMessageAt = nowIso();
  tracked.set(chatId, item);
}

function markReminderSent(chatId: string): void {
  const item = tracked.get(chatId);
  if (!item) return;

  const now = nowIso();
  item.lastNotifiedAt = now;
  item.lastMessageAt = now;
  tracked.set(chatId, item);
}

async function parseTrackCommandWithFlow(text: string): Promise<ParsedTrackResult> {
  const url = process.env.PA_PARSE_URL;

  if (!url) {
    throw new Error('PA_PARSE_URL is not configured');
  }

  const trimmedText = text.trim();

  if (!trimmedText) {
    return {
      reminderMs: DEFAULT_REMINDER_MS,
      agreedText: '',
      usedDefault: true
    };
  }

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15000);

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ text: trimmedText }),
      signal: controller.signal
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Flow request failed: ${response.status} ${errorText}`);
    }

    const data = await response.json();

    console.log('flow response:', JSON.stringify(data));

    const rawReminderMs = Number(data?.reminderMs);
    const rawAgreedText = sanitizeAgreedText(data?.agreedText);
    const rawUsedDefault = normalizeBoolean(data?.usedDefault);

    let reminderMs = rawReminderMs;
    let agreedText = rawAgreedText;
    let usedDefault = rawUsedDefault;

    if (!Number.isFinite(reminderMs)) {
      reminderMs = DEFAULT_REMINDER_MS;
      usedDefault = true;
    }

    if (reminderMs < MIN_REMINDER_MS || reminderMs > MAX_REMINDER_MS) {
      reminderMs = DEFAULT_REMINDER_MS;
      usedDefault = true;
    }

    if (!agreedText) {
      agreedText = trimmedText.slice(0, MAX_AGREED_TEXT_LENGTH);
    }

    return {
      reminderMs,
      agreedText,
      usedDefault
    };
  } finally {
    clearTimeout(timeout);
  }
}

botApp.on('message', async (ctx) => {
  const { send, activity } = ctx;
  const chatId = activity.conversation.id;
  const text = (activity.text || '').trim();

  console.log('message received');
  console.log('chatId:', chatId);
  console.log('text:', text);
  console.log('entities:', JSON.stringify(activity.entities || []));

  if (!text) {
    return;
  }

  if (/^done$/i.test(text)) {
    const stopped = stopTracking(chatId);

    if (stopped) {
      await send('Tracking has been stopped for this chat. No further reminders will be sent.');
    } else {
      await send('There is no active tracking configured for this chat.');
    }
    return;
  }

  if (/^\/track\b/i.test(text)) {
    const rawTrackText = text.replace(/^\/track\b/i, '').trim();
    const freeText = normalizeIncomingTrackText(rawTrackText);

    if (!freeText) {
      await send(
        'Please provide a reminder instruction, for example:\n\n' +
          '/track in 2 hours send PR update'
      );
      return;
    }

    try {
      const parsed = await parseTrackCommandWithFlow(freeText);

      if (!isValidReminderMs(parsed.reminderMs)) {
        await send(
          'Unable to start tracking because the parsed reminder interval is invalid. ' +
            'Please try something like:\n\n' +
            '/track in 2 hours send PR update'
        );
        return;
      }

      const agreedText = sanitizeAgreedText(parsed.agreedText);

      startTracking(activity, parsed.reminderMs, agreedText);

      const item = tracked.get(chatId);
      const readable = formatDuration(parsed.reminderMs);
      const formattedAgreedText = applyMentionsToText(
        agreedText || '<no message provided>',
        item?.mentions || []
      );

      const entities = buildMentionEntities(item?.mentions || []);

      await send({
        type: 'message',
        textFormat: 'markdown',
        text:
          `Tracking has been enabled for this chat.\n\n` +
          `Agreed action:\n\n` +
          `${formattedAgreedText}\n\n` +
          `Reminder interval: ${readable}\n\n` +
          `${parsed.usedDefault ? `Note: default reminder interval was used.\n\n` : ''}` +
          `I will continue sending reminders in this chat after each period of inactivity until someone types DONE.`,
        entities: entities as any
      });
    } catch (error) {
      console.error('Failed to parse /track command with flow', error);

      await send(
        'Unable to understand the reminder command. Try something like:\n\n' +
          '/track in 2 hours send PR update\n' +
          '/track tomorrow morning follow up with support\n' +
          '/track in 15 minutes ping the team'
      );
    }

    return;
  }

  const existing = tracked.get(chatId);
  if (existing?.isActive) {
    refreshLastMessage(chatId);
  }
});

async function checkInactiveChats() {
  console.log('checker tick, tracked chats count =', tracked.size);

  for (const item of tracked.values()) {
    if (!item.isActive) continue;
    if (item.channelId !== 'msteams') continue;

    if (!isValidReminderMs(item.reminderMs)) {
      console.warn('Skipping tracked item with invalid reminderMs for chat', item.chatId);
      continue;
    }

    if (!isOlderThan(item.lastMessageAt, item.reminderMs)) {
      continue;
    }

    const formattedAgreedText = applyMentionsToText(
      sanitizeAgreedText(item.agreedText) || '<no message provided>',
      item.mentions
    );

    try {
      if (item.agreedText) {
        await botApp.send(item.chatId, {
          type: 'message',
          textFormat: 'markdown',
          text:
            `Quick reminder about the agreed actions:\n\n` +
            `${formattedAgreedText}\n\n` +
            `I will continue sending reminders in this chat after each period of inactivity until someone types DONE.`,
          entities: buildMentionEntities(item.mentions) as any
        });
      } else {
        await botApp.send(
          item.chatId,
          'Reminder: this chat has been inactive for the configured period.'
        );
      }

      markReminderSent(item.chatId);
    } catch (error) {
      console.error('Failed to send reminder to', item.chatId, error);
    }
  }
}

setInterval(() => {
  void checkInactiveChats();
}, CHECK_EVERY_MS);

const PORT = Number(process.env.PORT) || 3978;
botApp.start(PORT).catch(console.error);