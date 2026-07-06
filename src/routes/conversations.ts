import { Router } from 'express';
import { z } from 'zod';
import type { ConversationStore } from '../storage/conversationStore.js';

const registerConversationSchema = z
  .object({
    userAadId: z.string().min(1).optional(),
    teamId: z.string().min(1).optional(),
    channelId: z.string().min(1).optional(),
    conversationId: z.string().min(1),
    serviceUrl: z.string().url().optional(),
    tenantId: z.string().min(1).optional(),
  })
  .refine(
    (data) =>
      Boolean(data.userAadId) ||
      (Boolean(data.teamId) && Boolean(data.channelId)),
    { message: 'Provide userAadId or teamId+channelId' },
  );

export function createConversationsRouter(conversationStore: ConversationStore): Router {
  const router = Router();

  router.post('/', (req, res) => {
    const parsed = registerConversationSchema.safeParse(req.body);

    if (!parsed.success) {
      res.status(400).json({
        success: false,
        error: parsed.error.issues.map((issue) => issue.message).join(', '),
      });
      return;
    }

    const data = parsed.data;

    if (data.userAadId) {
      conversationStore.saveUser(
        data.userAadId,
        data.conversationId,
        data.serviceUrl ?? null,
        data.tenantId ?? null,
      );
    } else if (data.teamId && data.channelId) {
      conversationStore.saveChannel(
        data.teamId,
        data.channelId,
        data.conversationId,
        data.serviceUrl ?? null,
        data.tenantId ?? null,
      );
    }

    res.json({ success: true, conversationId: data.conversationId });
  });

  router.get('/', (req, res) => {
    const limit = Math.min(Number(req.query.limit ?? 50), 200);
    const offset = Number(req.query.offset ?? 0);
    const items = conversationStore.list(limit, offset);

    res.json({
      success: true,
      total: conversationStore.count(),
      limit,
      offset,
      items,
    });
  });

  router.get('/users/:aadObjectId', (req, res) => {
    const conversation = conversationStore.getByUserAadId(req.params.aadObjectId);

    if (!conversation) {
      res.status(404).json({ success: false, error: 'Conversation not found' });
      return;
    }

    res.json({ success: true, conversation });
  });

  router.get('/channels/:teamId/:channelId', (req, res) => {
    const conversation = conversationStore.getByTeamChannel(req.params.teamId, req.params.channelId);

    if (!conversation) {
      res.status(404).json({ success: false, error: 'Conversation not found' });
      return;
    }

    res.json({ success: true, conversation });
  });

  return router;
}
