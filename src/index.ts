import express from 'express';
import { loadConfig } from './config.js';
import { errorHandler } from './middleware/errorHandler.js';
import { createTeamsApplication } from './teams/app.js';

async function main(): Promise<void> {
  const config = loadConfig();
  const server = express();

  server.use(express.json());

  const { teamsApp } = createTeamsApplication(config, server);

  server.use(errorHandler);

  await teamsApp.initialize();

  server.listen(config.PORT, '0.0.0.0', () => {
    console.log(`teams-sdk-rest listening on 0.0.0.0:${config.PORT}`);
    console.log(`Health: GET /health`);
    console.log(`Debug:  POST /debug/echo (no auth, logs request body)`);
    console.log(`Teams webhook: POST /api/messages (Bot Framework auth required)`);
    console.log(`REST API: /v1/* (Bearer token required)`);
    console.log(
      `Laravel webhook: ${config.LARAVEL_WEBHOOK_URL ?? '(not configured — inbound replies disabled)'}`,
    );
  });
}

main().catch((error) => {
  console.error('Failed to start:', error);
  process.exit(1);
});
