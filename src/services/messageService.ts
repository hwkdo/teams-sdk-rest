import { MessageActivity, TypingActivity } from '@microsoft/teams.api';
import type { IAdaptiveCard } from '@microsoft/teams.cards';
import type { App } from '@microsoft/teams.apps';
import type { ConversationStore } from '../storage/conversationStore.js';
import type { ReplyMessageRequest, SendMessageRequest } from '../types/api.js';

export class MessageService {
  constructor(
    private readonly getApp: () => App,
    private readonly conversationStore: ConversationStore,
  ) {}

  async send(request: SendMessageRequest): Promise<{ messageId: string; conversationId: string }> {
    const conversationId = this.resolveConversationId(request);
    const activity = this.buildMessageActivity(request.text, request.html, request.card);
    const result = await this.getApp().send(conversationId, activity);

    return {
      messageId: result.id,
      conversationId,
    };
  }

  async reply(request: ReplyMessageRequest): Promise<{ messageId?: string; conversationId: string }> {
    const conversationId = this.normalizeConversationId(request.conversationId);
    const activity = this.buildMessageActivity(request.text, request.html, request.card);
    const result = await this.getApp().send(conversationId, activity);

    return {
      messageId: result.id,
      conversationId,
    };
  }

  async typing(conversationId: string): Promise<void> {
    await this.getApp().send(conversationId, new TypingActivity());
  }

  private resolveConversationId(request: SendMessageRequest): string {
    if (request.conversationId) {
      return request.conversationId;
    }

    if (request.userAadId) {
      const stored = this.conversationStore.getByUserAadId(request.userAadId);
      if (!stored) {
        throw new Error(
          `No conversation found for user ${request.userAadId}. Install the bot for this user first.`,
        );
      }
      return stored.conversationId;
    }

    if (request.channelId) {
      const stored = request.teamId
        ? this.conversationStore.getByTeamChannel(request.teamId, request.channelId)
        : null;

      // For a new top-level channel post the conversation id equals the channel
      // thread id. A stored reference is only needed to reply into an existing
      // thread, so fall back to the channel id when nothing is stored.
      return stored?.conversationId ?? request.channelId;
    }

    throw new Error('Provide conversationId, userAadId, or teamId+channelId');
  }

  private normalizeConversationId(conversationId: string): string {
    const separatorIndex = conversationId.indexOf(';messageid=');

    if (separatorIndex === -1) {
      return conversationId;
    }

    return conversationId.slice(0, separatorIndex);
  }

  private buildMessageActivity(text?: string, html?: string, card?: IAdaptiveCard): MessageActivity {
    if (!text && !html && !card) {
      throw new Error('Provide at least one of text, html, or card');
    }

    const content = html ?? text ?? ' ';
    const activity = new MessageActivity(content);

    if (html) {
      activity.withTextFormat('xml');
    }

    if (card) {
      activity.addCard('adaptive', card);
    }

    return activity;
  }
}
