// src/storage.ts

export type MentionRecord = {
  id: string;
  name: string;
  text: string;
};

export type ChatRecord = {
  chatId: string;
  conversationId: string;
  channelId: string;
  serviceUrl: string;
  tenantId?: string;
  reminderMs: number;
  agreedText: string;
  mentions: MentionRecord[];
  isActive: boolean;
  lastMessageAt: string | null;
  lastNotifiedAt: string | null;
};

const chats = new Map<string, ChatRecord>();

export function upsertChat(record: ChatRecord) {
  chats.set(record.chatId, record);
}

export function getChat(chatId: string) {
  return chats.get(chatId);
}

export function getAllChats() {
  return [...chats.values()];
}

export function updateLastMessageAt(chatId: string, iso: string) {
  const item = chats.get(chatId);
  if (!item) return;

  item.lastMessageAt = iso;
  chats.set(chatId, item);
}

export function updateLastNotifiedAt(chatId: string, iso: string) {
  const item = chats.get(chatId);
  if (!item) return;

  item.lastNotifiedAt = iso;
  chats.set(chatId, item);
}

export function setChatActive(chatId: string, isActive: boolean) {
  const item = chats.get(chatId);
  if (!item) return;

  item.isActive = isActive;
  chats.set(chatId, item);
}

export function updateTrackedConfig(
  chatId: string,
  updates: Partial<Pick<ChatRecord, 'reminderMs' | 'agreedText' | 'mentions' | 'channelId' | 'serviceUrl' | 'tenantId'>>
) {
  const item = chats.get(chatId);
  if (!item) return;

  const next: ChatRecord = {
    ...item,
    ...updates
  };

  chats.set(chatId, next);
}