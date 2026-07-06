import fs from 'node:fs';
import path from 'node:path';
import Database from 'better-sqlite3';
import type { Activity } from '@microsoft/teams.api';
import type { StoredConversation } from '../types/api.js';

type ConversationRow = {
  id: number;
  type: 'user' | 'channel';
  user_aad_id: string | null;
  team_id: string | null;
  channel_id: string | null;
  conversation_id: string;
  service_url: string | null;
  tenant_id: string | null;
  updated_at: string;
};

type MemberRecord = {
  aadObjectId?: string;
  id?: string;
};

export class ConversationStore {
  private db: Database.Database;

  constructor(dataDir: string) {
    fs.mkdirSync(dataDir, { recursive: true });
    const dbPath = path.join(dataDir, 'conversations.db');
    this.db = new Database(dbPath);
    this.db.pragma('journal_mode = WAL');
    this.initSchema();
  }

  private initSchema(): void {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS conversations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        type TEXT NOT NULL CHECK(type IN ('user', 'channel')),
        user_aad_id TEXT,
        team_id TEXT,
        channel_id TEXT,
        conversation_id TEXT NOT NULL,
        service_url TEXT,
        tenant_id TEXT,
        updated_at TEXT NOT NULL,
        UNIQUE(user_aad_id),
        UNIQUE(team_id, channel_id)
      );
      CREATE INDEX IF NOT EXISTS idx_conversations_type ON conversations(type);
    `);
  }

  saveFromActivity(activity: Activity): void {
    const conversationId = activity.conversation?.id;
    if (!conversationId) {
      return;
    }

    const serviceUrl = activity.serviceUrl ?? null;
    const tenantId = activity.conversation?.tenantId ?? activity.channelData?.tenant?.id ?? null;
    const teamId = activity.channelData?.team?.id ?? null;
    const channelId = activity.channelData?.channel?.id ?? null;
    const userAadId = activity.from?.aadObjectId ?? null;
    const isPersonal = activity.conversation?.conversationType === 'personal';

    if (teamId && channelId) {
      this.saveChannel(teamId, channelId, conversationId, serviceUrl, tenantId);
    }

    if (userAadId && isPersonal) {
      this.saveUser(userAadId, conversationId, serviceUrl, tenantId);
    }

    const membersAdded = this.readMembersAdded(activity);
    if (membersAdded.length > 0 && isPersonal) {
      for (const member of membersAdded) {
        const memberAadId = member.aadObjectId ?? this.extractAadIdFromMemberId(member.id);

        if (memberAadId) {
          this.saveUser(memberAadId, conversationId, serviceUrl, tenantId);
        }
      }
    }
  }

  private readMembersAdded(activity: Activity): MemberRecord[] {
    if (!('membersAdded' in activity)) {
      return [];
    }

    const membersAdded = (activity as Activity & { membersAdded?: unknown }).membersAdded;

    if (!Array.isArray(membersAdded)) {
      return [];
    }

    return membersAdded.filter(
      (member): member is MemberRecord => typeof member === 'object' && member !== null,
    );
  }

  private extractAadIdFromMemberId(memberId?: string): string | null {
    if (!memberId) {
      return null;
    }

    const match = memberId.match(/29:1([0-9a-f-]{36})/i);
    if (match?.[1]) {
      return match[1].toLowerCase();
    }

    if (/^[0-9a-f-]{36}$/i.test(memberId)) {
      return memberId.toLowerCase();
    }

    return null;
  }

  saveUser(userAadId: string, conversationId: string, serviceUrl?: string | null, tenantId?: string | null): void {
    const normalizedUserAadId = userAadId.toLowerCase();
    const now = new Date().toISOString();
    this.db
      .prepare(
        `INSERT INTO conversations (type, user_aad_id, team_id, channel_id, conversation_id, service_url, tenant_id, updated_at)
         VALUES ('user', ?, NULL, NULL, ?, ?, ?, ?)
         ON CONFLICT(user_aad_id) DO UPDATE SET
           conversation_id = excluded.conversation_id,
           service_url = excluded.service_url,
           tenant_id = excluded.tenant_id,
           updated_at = excluded.updated_at`,
      )
      .run(normalizedUserAadId, conversationId, serviceUrl ?? null, tenantId ?? null, now);
  }

  saveChannel(
    teamId: string,
    channelId: string,
    conversationId: string,
    serviceUrl?: string | null,
    tenantId?: string | null,
  ): void {
    const now = new Date().toISOString();
    this.db
      .prepare(
        `INSERT INTO conversations (type, user_aad_id, team_id, channel_id, conversation_id, service_url, tenant_id, updated_at)
         VALUES ('channel', NULL, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(team_id, channel_id) DO UPDATE SET
           conversation_id = excluded.conversation_id,
           service_url = excluded.service_url,
           tenant_id = excluded.tenant_id,
           updated_at = excluded.updated_at`,
      )
      .run(teamId, channelId, conversationId, serviceUrl ?? null, tenantId ?? null, now);
  }

  getByUserAadId(userAadId: string): StoredConversation | null {
    const row = this.db
      .prepare(`SELECT * FROM conversations WHERE user_aad_id = ?`)
      .get(userAadId.toLowerCase()) as ConversationRow | undefined;
    return row ? this.toStoredConversation(row) : null;
  }

  getByTeamChannel(teamId: string, channelId: string): StoredConversation | null {
    const row = this.db
      .prepare(`SELECT * FROM conversations WHERE team_id = ? AND channel_id = ?`)
      .get(teamId, channelId) as ConversationRow | undefined;
    return row ? this.toStoredConversation(row) : null;
  }

  list(limit = 50, offset = 0): StoredConversation[] {
    const rows = this.db
      .prepare(`SELECT * FROM conversations ORDER BY updated_at DESC LIMIT ? OFFSET ?`)
      .all(limit, offset) as ConversationRow[];
    return rows.map((row) => this.toStoredConversation(row));
  }

  count(): number {
    const row = this.db.prepare(`SELECT COUNT(*) as count FROM conversations`).get() as { count: number };
    return row.count;
  }

  private toStoredConversation(row: ConversationRow): StoredConversation {
    return {
      id: row.id,
      type: row.type,
      userAadId: row.user_aad_id,
      teamId: row.team_id,
      channelId: row.channel_id,
      conversationId: row.conversation_id,
      serviceUrl: row.service_url,
      tenantId: row.tenant_id,
      updatedAt: row.updated_at,
    };
  }
}
