import * as endpoints from '@microsoft/teams.graph-endpoints';
import type { App } from '@microsoft/teams.apps';

export class GraphService {
  constructor(private readonly getApp: () => App) {}

  async getMe(): Promise<unknown> {
    return this.getApp().graph.call(endpoints.me.get);
  }

  async getUser(idOrUpn: string): Promise<unknown> {
    return this.getApp().graph.call(endpoints.users.get, {
      'user-id': idOrUpn,
    });
  }

  async searchUsers(search: string, top = 25): Promise<unknown> {
    return this.getApp().graph.call(endpoints.users.list, {
      $search: `"displayName:${search}" OR "mail:${search}" OR "userPrincipalName:${search}"`,
      $top: top,
      $select: ['id', 'displayName', 'mail', 'userPrincipalName', 'jobTitle'],
      ConsistencyLevel: 'eventual',
    });
  }

  async listTeams(top = 50): Promise<unknown> {
    return this.getApp().graph.call(endpoints.teams.list, {
      $top: top,
      $select: ['id', 'displayName', 'description'],
    });
  }

  async listTeamChannels(teamId: string): Promise<unknown> {
    return this.getApp().graph.call(endpoints.teams.channels.list, {
      'team-id': teamId,
    });
  }
}
