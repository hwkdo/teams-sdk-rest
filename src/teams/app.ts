import express from 'express';
import { cloudFromName } from '@microsoft/teams.api';
import { App, ExpressAdapter } from '@microsoft/teams.apps';
import type { Config } from '../config.js';
import { ConversationStore } from '../storage/conversationStore.js';
import { WebhookForwarder } from '../services/webhookForwarder.js';
import { MessageService } from '../services/messageService.js';
import { GraphService } from '../services/graphService.js';
import { createApiAuth } from '../middleware/apiAuth.js';
import { createHealthRouter } from '../routes/health.js';
import { createMessagesRouter } from '../routes/messages.js';
import { createConversationsRouter } from '../routes/conversations.js';
import { createGraphRouter } from '../routes/graph.js';
import { logInboundEvent } from '../lib/eventLogger.js';

export type TeamsServices = {
  conversationStore: ConversationStore;
  webhookForwarder: WebhookForwarder;
  messageService: MessageService;
  graphService: GraphService;
};

export function createTeamsApplication(config: Config, server: express.Express): {
  teamsApp: App;
  services: TeamsServices;
} {
  const conversationStore = new ConversationStore(config.DATA_DIR);
  const webhookForwarder = new WebhookForwarder(config);

  const teamsApp = new App({
    httpServerAdapter: new ExpressAdapter(server),
    ...(config.CLOUD ? { cloud: cloudFromName(config.CLOUD) } : {}),
  });

  const getApp = () => teamsApp;
  const messageService = new MessageService(getApp, conversationStore);
  const graphService = new GraphService(getApp);

  teamsApp.on('install.add', async ({ activity, send }) => {
    logInboundEvent('install.add', {
      conversationId: activity.conversation?.id,
      userAadId: activity.from?.aadObjectId,
    });

    conversationStore.saveFromActivity(activity);
    await webhookForwarder.forward('install.add', activity);

    const welcomeMessage = config.WELCOME_MESSAGE.trim();

    if (welcomeMessage !== '') {
      await send(welcomeMessage);
    }
  });

  teamsApp.on('install.remove', async ({ activity }) => {
    await webhookForwarder.forward('install.remove', activity);
  });

  teamsApp.on('message', async ({ activity }) => {
    logInboundEvent('message', {
      conversationId: activity.conversation?.id,
      activityId: activity.id,
      userAadId: activity.from?.aadObjectId,
      fromId: activity.from?.id,
      text: typeof activity.text === 'string' ? activity.text.slice(0, 120) : undefined,
    });

    conversationStore.saveFromActivity(activity);
    await webhookForwarder.forward('message', activity);
  });

  teamsApp.on('mention', async ({ activity }) => {
    conversationStore.saveFromActivity(activity);
    await webhookForwarder.forward('mention', activity);
  });

  teamsApp.on('channelMemberAdded', async ({ activity }) => {
    conversationStore.saveFromActivity(activity);
    await webhookForwarder.forward('conversationUpdate.channelMemberAdded', activity);
  });

  teamsApp.on('invoke', async ({ activity }) => {
    if (activity.name === 'adaptiveCard/action') {
      conversationStore.saveFromActivity(activity);
      await webhookForwarder.forward('adaptive-card.action', activity);
    }
  });

  const apiAuth = createApiAuth(config);

  server.use(createHealthRouter());
  server.use('/v1/messages', apiAuth, createMessagesRouter(messageService));
  server.use('/v1/conversations', apiAuth, createConversationsRouter(conversationStore));
  server.use('/v1/graph', apiAuth, createGraphRouter(graphService));

  return {
    teamsApp,
    services: {
      conversationStore,
      webhookForwarder,
      messageService,
      graphService,
    },
  };
}
