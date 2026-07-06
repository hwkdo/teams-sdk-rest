import type { IAdaptiveCard } from '@microsoft/teams.cards';

export type ApiSuccessResponse<T = Record<string, unknown>> = {
  success: true;
} & T;

export type ApiErrorResponse = {
  success: false;
  error: string;
  details?: unknown;
};

export type SendMessageRequest = {
  conversationId?: string;
  userAadId?: string;
  teamId?: string;
  channelId?: string;
  text?: string;
  html?: string;
  card?: IAdaptiveCard;
};

export type ReplyMessageRequest = {
  conversationId: string;
  messageId?: string;
  text?: string;
  html?: string;
  card?: IAdaptiveCard;
};

export type TypingRequest = {
  conversationId: string;
};

export type StoredConversation = {
  id: number;
  type: 'user' | 'channel';
  userAadId: string | null;
  teamId: string | null;
  channelId: string | null;
  conversationId: string;
  serviceUrl: string | null;
  tenantId: string | null;
  updatedAt: string;
};

export type RegisterConversationRequest = {
  userAadId?: string;
  teamId?: string;
  channelId?: string;
  conversationId: string;
  serviceUrl?: string;
  tenantId?: string;
};

export type WebhookPayload = {
  event: string;
  timestamp: string;
  activity: unknown;
  conversationRef: {
    conversationId?: string;
    userAadId?: string;
    teamId?: string;
    channelId?: string;
    serviceUrl?: string;
    tenantId?: string;
  };
};
