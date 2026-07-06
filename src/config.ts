import { z } from 'zod';

const configSchema = z.object({
  CLIENT_ID: z.string().min(1),
  CLIENT_SECRET: z.string().min(1),
  TENANT_ID: z.string().min(1),
  PORT: z.coerce.number().default(3978),
  API_KEY: z.string().min(1),
  LARAVEL_WEBHOOK_URL: z.string().url().optional(),
  LARAVEL_WEBHOOK_SECRET: z.string().optional(),
  DATA_DIR: z.string().default('./data'),
  CLOUD: z.enum(['Public', 'USGov', 'USGovDoD', 'China']).optional(),
  WELCOME_MESSAGE: z
    .string()
    .default('Bot installiert — bereit für Benachrichtigungen.'),
  HI_REPLY_MESSAGE: z
    .string()
    .default(
      'Hallo! Schön, dass du da bist. Ich sende dir Benachrichtigungen aus dem HWKDO Intranet.',
    ),
});

export type Config = z.infer<typeof configSchema>;

export function loadConfig(): Config {
  const result = configSchema.safeParse(process.env);

  if (!result.success) {
    const missing = result.error.issues.map((issue) => issue.path.join('.')).join(', ');
    throw new Error(`Invalid configuration: ${missing}`);
  }

  return result.data;
}
