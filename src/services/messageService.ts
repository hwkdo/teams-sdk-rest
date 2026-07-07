import { Client as ApiClient, MessageActivity, TypingActivity, toActivityParams } from '@microsoft/teams.api';
import type { IAdaptiveCard } from '@microsoft/teams.cards';
import type { App } from '@microsoft/teams.apps';
import type { ConversationStore } from '../storage/conversationStore.js';
import type { ReplyMessageRequest, SendMessageRequest } from '../types/api.js';

type ResolvedTarget = {
  conversationId: string;
  serviceUrl?: string;
};

export class MessageService {
  constructor(
    private readonly getApp: () => App,
    private readonly conversationStore: ConversationStore,
  ) {}

  async send(request: SendMessageRequest): Promise<{ messageId: string; conversationId: string }> {
    // A channel target without an explicit conversationId means we start a new
    // top-level thread. Sending directly to the channel id triggers
    // "BotNotInConversationRoster"; the Bot Connector requires creating a new
    // conversation with channelData instead.
    if (request.channelId && !request.conversationId) {
      return this.sendToChannel(request);
    }

    const target = this.resolveTarget(request);
    const activity = this.buildMessageActivity(request.text, request.html, request.card);
    return this.dispatch(target, activity);
  }

  private async sendToChannel(
    request: SendMessageRequest,
  ): Promise<{ messageId: string; conversationId: string }> {
    const channelId = request.channelId!;
    const activity = this.buildMessageActivity(request.text, request.html, request.card);
    const app = this.getApp();

    // Prefer a stored regional serviceUrl for this team/channel; otherwise fall
    // back to the app's default API client.
    const stored = request.teamId
      ? this.conversationStore.getByTeamChannel(request.teamId, channelId)
      : null;
    const api = stored?.serviceUrl ? new ApiClient(stored.serviceUrl, app.api.http) : app.api;

    const resource = await api.conversations.create({
      isGroup: true,
      channelData: { channel: { id: channelId } },
      activity: toActivityParams(activity),
    });

    return { messageId: resource.activityId, conversationId: resource.id };
  }

  async reply(request: ReplyMessageRequest): Promise<{ messageId: string; conversationId: string }> {
    const baseConversationId = this.normalizeConversationId(request.conversationId);
    const activity = this.buildMessageActivity(request.text, request.html, request.card);
    const serviceUrl = this.conversationStore.getByConversationId(baseConversationId)?.serviceUrl ?? undefined;

    // Post into an existing thread when a messageId is supplied, otherwise the
    // reply becomes a new top-level message in the conversation.
    const conversationId = request.messageId
      ? `${baseConversationId};messageid=${request.messageId}`
      : baseConversationId;

    return this.dispatch({ conversationId, serviceUrl }, activity);
  }

  async typing(conversationId: string): Promise<void> {
    const normalized = this.normalizeConversationId(conversationId);
    const serviceUrl = this.conversationStore.getByConversationId(normalized)?.serviceUrl ?? undefined;
    await this.dispatch({ conversationId: normalized, serviceUrl }, new TypingActivity());
  }

  private async dispatch(
    target: ResolvedTarget,
    activity: MessageActivity | TypingActivity,
  ): Promise<{ messageId: string; conversationId: string }> {
    const app = this.getApp();

    // When we know the conversation's regional serviceUrl, send through an API
    // client bound to it. The global app.send() default (SERVICE_URL env) does
    // not always match the tenant's region, which yields 403s from the Bot Connector.
    if (target.serviceUrl) {
      const botId = app.id;
      if (!botId) {
        throw new Error('App has no credentials set up');
      }

      const api = new ApiClient(target.serviceUrl, app.api.http);
      const params = toActivityParams(activity);
      // The conversation id is already encoded in the request path; the Bot
      // Connector resolves the target from there, so we only need to set the sender.
      const res = await api.conversations.activities(target.conversationId).create({
        ...params,
        from: { id: botId, role: 'bot' },
      });

      return { messageId: res.id, conversationId: target.conversationId };
    }

    const res = await app.send(target.conversationId, activity);
    return { messageId: res.id, conversationId: target.conversationId };
  }

  private resolveTarget(request: SendMessageRequest): ResolvedTarget {
    if (request.conversationId) {
      const stored = this.conversationStore.getByConversationId(
        this.normalizeConversationId(request.conversationId),
      );
      return {
        conversationId: request.conversationId,
        serviceUrl: stored?.serviceUrl ?? undefined,
      };
    }

    if (request.userAadId) {
      const stored = this.conversationStore.getByUserAadId(request.userAadId);
      if (!stored) {
        throw new Error(
          `No conversation found for user ${request.userAadId}. Install the bot for this user first.`,
        );
      }
      return { conversationId: stored.conversationId, serviceUrl: stored.serviceUrl ?? undefined };
    }

    if (request.channelId) {
      const stored = request.teamId
        ? this.conversationStore.getByTeamChannel(request.teamId, request.channelId)
        : null;

      // For a new top-level channel post the conversation id equals the channel
      // thread id. A stored reference is only needed to reply into an existing
      // thread, so fall back to the channel id when nothing is stored.
      return {
        conversationId: stored?.conversationId ?? request.channelId,
        serviceUrl: stored?.serviceUrl ?? undefined,
      };
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
