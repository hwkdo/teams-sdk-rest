import { Router } from 'express';

export function createHealthRouter(): Router {
  const router = Router();

  router.get('/', (_req, res) => {
    res.json({
      status: 'ok',
      service: 'teams-sdk-rest',
      endpoints: {
        health: 'GET /health',
        debug: 'POST /debug/echo',
        teamsWebhook: 'POST /api/messages',
        restApi: '/v1/* (Bearer API_KEY)',
      },
    });
  });

  router.get('/health', (_req, res) => {
    res.json({ status: 'ok', service: 'teams-sdk-rest' });
  });

  router.post('/debug/echo', (req, res) => {
    console.log('[teams-sdk-rest] debug echo body=', JSON.stringify(req.body));

    res.json({
      success: true,
      service: 'teams-sdk-rest',
      received: req.body ?? null,
    });
  });

  return router;
}
