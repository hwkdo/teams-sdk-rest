export function logInboundEvent(event: string, details: Record<string, unknown> = {}): void {
  const payload = Object.entries(details)
    .filter(([, value]) => value !== undefined && value !== null && value !== '')
    .map(([key, value]) => `${key}=${String(value)}`)
    .join(' ');

  console.log(`[teams-sdk-rest] inbound event=${event}${payload ? ` ${payload}` : ''}`);
}

export function logWebhookForward(
  event: string,
  url: string,
  status: 'skipped' | 'ok' | 'failed',
  detail?: string,
): void {
  const suffix = detail ? ` (${detail})` : '';

  console.log(`[teams-sdk-rest] webhook forward event=${event} url=${url} status=${status}${suffix}`);
}
