import { Router } from 'express';
import type { GraphService } from '../services/graphService.js';

export function createGraphRouter(graphService: GraphService): Router {
  const router = Router();

  router.get('/me', async (_req, res, next) => {
    try {
      const data = await graphService.getMe();
      res.json({ success: true, data });
    } catch (error) {
      next(error);
    }
  });

  router.get('/users', async (req, res, next) => {
    try {
      const search = String(req.query.search ?? '');

      if (!search) {
        res.status(400).json({ success: false, error: 'Query parameter search is required' });
        return;
      }

      const top = Number(req.query.top ?? 25);
      const data = await graphService.searchUsers(search, top);
      res.json({ success: true, data });
    } catch (error) {
      next(error);
    }
  });

  router.get('/users/:idOrUpn', async (req, res, next) => {
    try {
      const data = await graphService.getUser(req.params.idOrUpn);
      res.json({ success: true, data });
    } catch (error) {
      next(error);
    }
  });

  router.get('/teams', async (req, res, next) => {
    try {
      const top = Number(req.query.top ?? 50);
      const data = await graphService.listTeams(top);
      res.json({ success: true, data });
    } catch (error) {
      next(error);
    }
  });

  router.get('/teams/:teamId/channels', async (req, res, next) => {
    try {
      const data = await graphService.listTeamChannels(req.params.teamId);
      res.json({ success: true, data });
    } catch (error) {
      next(error);
    }
  });

  return router;
}
