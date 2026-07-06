import crypto from 'node:crypto';
import type { Activity } from '@microsoft/teams.api';
import type { Config } from '../config.js';
import type { WebhookPayload } from '../types/api.js';
import { logWebhookForward } from '../lib/eventLogger.js';

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

  async forward(event: string, activity: Activity): Promise<void> {
    if (!this.config.LARAVEL_WEBHOOK_URL) {
      logWebhookForward(event, '(not configured)', 'skipped', 'LARAVEL_WEBHOOK_URL missing');
      return;
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

    try {
      const response = await fetch(this.config.LARAVEL_WEBHOOK_URL, {
        method: 'POST',
        headers,
        body,
      });

      if (!response.ok) {
        const responseBody = await response.text();
        logWebhookForward(event, webhookUrl, 'failed', `HTTP ${response.status}: ${responseBody}`);
        return;
      }

      logWebhookForward(event, webhookUrl, 'ok', `HTTP ${response.status}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logWebhookForward(event, webhookUrl, 'failed', message);
    }
  }
}
