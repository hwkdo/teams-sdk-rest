import crypto from 'node:crypto';
import type { Activity } from '@microsoft/teams.api';
import type { Config } from '../config.js';
import type { WebhookPayload } from '../types/api.js';
import { logWebhookForward } from '../lib/eventLogger.js';

export type WebhookForwardResult = {
  ok: boolean;
  status: number | null;
  body: unknown;
};

export class WebhookForwarder {
  constructor(private readonly config: Config) {}

  private serializeActivity(activity: Activity): unknown {
    if (typeof activity === 'object' && activity !== null && 'toInterface' in activity) {
      const toInterface = (activity as { toInterface?: () => unknown }).toInterface;
      if (typeof toInterface === 'function') {
        return toInterface.call(activity);
      }
    }

    return JSON.parse(JSON.stringify(activity));
  }

  async forward(
    event: string,
    activity: Activity,
    options: { timeoutMs?: number } = {},
  ): Promise<WebhookForwardResult> {
    if (!this.config.LARAVEL_WEBHOOK_URL) {
      logWebhookForward(event, '(not configured)', 'skipped', 'LARAVEL_WEBHOOK_URL missing');
      return { ok: false, status: null, body: null };
    }

    const webhookUrl = this.config.LARAVEL_WEBHOOK_URL;
    const payload: WebhookPayload = {
      event,
      timestamp: new Date().toISOString(),
      activity: this.serializeActivity(activity),
      conversationRef: {
        conversationId: activity.conversation?.id,
        userAadId: activity.from?.aadObjectId,
        teamId: activity.channelData?.team?.id,
        channelId: activity.channelData?.channel?.id,
        serviceUrl: activity.serviceUrl,
        tenantId: activity.conversation?.tenantId ?? activity.channelData?.tenant?.id,
      },
    };

    const body = JSON.stringify(payload);
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'X-Teams-Event': event,
    };

    if (this.config.LARAVEL_WEBHOOK_SECRET) {
      const signature = crypto
        .createHmac('sha256', this.config.LARAVEL_WEBHOOK_SECRET)
        .update(body)
        .digest('hex');
      headers['X-Teams-Signature'] = `sha256=${signature}`;
    }

    const timeoutMs = options.timeoutMs ?? 10_000;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);

    try {
      const response = await fetch(this.config.LARAVEL_WEBHOOK_URL, {
        method: 'POST',
        headers,
        body,
        signal: controller.signal,
      });

      const responseText = await response.text();
      let responseBody: unknown = null;

      if (responseText !== '') {
        try {
          responseBody = JSON.parse(responseText);
        } catch {
          responseBody = responseText;
        }
      }

      if (!response.ok) {
        logWebhookForward(
          event,
          webhookUrl,
          'failed',
          `HTTP ${response.status}: ${typeof responseBody === 'string' ? responseBody : JSON.stringify(responseBody)}`,
        );
        return { ok: false, status: response.status, body: responseBody };
      }

      logWebhookForward(event, webhookUrl, 'ok', `HTTP ${response.status}`);
      return { ok: true, status: response.status, body: responseBody };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logWebhookForward(event, webhookUrl, 'failed', message);
      return { ok: false, status: null, body: null };
    } finally {
      clearTimeout(timeout);
    }
  }
}
