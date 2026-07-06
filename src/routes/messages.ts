import { Router } from 'express';
import { z } from 'zod';
import type { MessageService } from '../services/messageService.js';
import type { ReplyMessageRequest, SendMessageRequest } from '../types/api.js';

const sendSchema = z
  .object({
    conversationId: z.string().optional(),
    userAadId: z.string().optional(),
    teamId: z.string().optional(),
    channelId: z.string().optional(),
    text: z.string().optional(),
    html: z.string().optional(),
    card: z.record(z.string(), z.unknown()).optional(),
  })
  .refine((data) => data.conversationId || data.userAadId || (data.teamId && data.channelId), {
    message: 'Provide conversationId, userAadId, or teamId+channelId',
  })
  .refine((data) => data.text || data.html || data.card, {
    message: 'Provide at least one of text, html, or card',
  });

const replySchema = z
  .object({
    conversationId: z.string(),
    messageId: z.string().optional(),
    text: z.string().optional(),
    html: z.string().optional(),
    card: z.record(z.string(), z.unknown()).optional(),
  })
  .refine((data) => data.text || data.html || data.card, {
    message: 'Provide at least one of text, html, or card',
  });

const typingSchema = z.object({
  conversationId: z.string(),
});

export function createMessagesRouter(messageService: MessageService): Router {
  const router = Router();

  router.post('/', async (req, res, next) => {
    try {
      const body = sendSchema.parse(req.body);
      const result = await messageService.send({
        ...body,
        card: body.card as SendMessageRequest['card'],
      });
      res.json({ success: true, ...result });
    } catch (error) {
      next(error);
    }
  });

  router.post('/reply', async (req, res, next) => {
    try {
      const body = replySchema.parse(req.body);
      const result = await messageService.reply({
        ...body,
        card: body.card as ReplyMessageRequest['card'],
      });
      res.json({ success: true, ...result });
    } catch (error) {
      next(error);
    }
  });

  router.post('/typing', async (req, res, next) => {
    try {
      const body = typingSchema.parse(req.body);
      await messageService.typing(body.conversationId);
      res.json({ success: true, conversationId: body.conversationId });
    } catch (error) {
      next(error);
    }
  });

  return router;
}
