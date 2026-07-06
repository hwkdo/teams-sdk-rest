# teams-sdk-rest

REST-API-Wrapper für das [Microsoft Teams SDK](https://microsoft.github.io/teams-sdk/typescript/getting-started/code-basics), damit Laravel-Apps Teams-Nachrichten senden und eingehende Events empfangen können.

## Architektur

- **Ausgehend:** Laravel → `POST /v1/*` (Bearer `API_KEY`) → Teams SDK `app.send` / `app.reply` → Bot Connector API
- **Eingehend:** Teams → `POST /api/messages` (SDK) → Weiterleitung an Laravel-Webhook
- **Graph (read-only):** User, Teams, Channels über `app.graph`

## Debug / Connectivity (Postman)

**502 Bad Gateway** kommt fast immer vom Reverse Proxy (Nginx/Traefik), nicht vom Node-Service.
Der Container antwortet nicht oder der Proxy zeigt auf den falschen Upstream.

### 1. Direkt am Container testen (ohne Proxy)

Port laut `docker compose ps` (Standard: **3978**).

| Request | Auth | Erwartung |
|---------|------|-----------|
| `GET http://localhost:3978/health` | keine | `200` + `{"status":"ok"}` |
| `POST http://localhost:3978/debug/echo` | keine | `200`, Body wird in Container-Logs geloggt |

Postman Body (JSON) für `/debug/echo`:
```json
{ "test": "hallo" }
```

Logs:
```bash
docker logs -f <container-name>
# → [teams-sdk-rest] debug echo body= {"test":"hallo"}
```

### 2. Was **nicht** mit bare Postman funktioniert

`POST /api/messages` ist der **Microsoft Teams / Bot Framework** Endpoint.
Er erwartet ein gültiges Bot-Framework-JWT und eine Activity — ein einfacher Postman-POST liefert
höchstens `401`/`403`, wenn der Request den Container überhaupt erreicht.

### 3. REST-API testen

```
POST http://localhost:3978/v1/messages
Authorization: Bearer <API_KEY aus .env>
Content-Type: application/json

{ "conversationId": "...", "text": "Test" }
```

### 4. Bei 502 über öffentliche URL

Dann trifft Postman den **Proxy**, nicht den Container direkt.

- Container läuft? `docker ps | grep teams-sdk`
- Logs beim Start? `docker logs <container>` — crasht er wegen fehlender `CLIENT_ID`/`API_KEY`?
- Proxy-Upstream korrekt? (Host `teams-sdk-rest` im Docker-Netzwerk, Port `3978`)
- Öffentlich: `https://<host>/api/messages` → Proxy → `http://teams-sdk-rest:3978/api/messages`

## Schnellstart

```bash
cp .env.example .env
# CLIENT_ID, CLIENT_SECRET, TENANT_ID, API_KEY eintragen

npm install
npm run build
npm start

# oder mit Docker
docker compose up --build
```

## CI/CD (Docker Registry)

Bei jedem Push auf `main` baut [`.github/workflows/docker-publish.yml`](.github/workflows/docker-publish.yml) das Image und lädt es nach **GitHub Container Registry** hoch:

```
ghcr.io/hwkdo/teams-sdk-rest:latest
ghcr.io/hwkdo/teams-sdk-rest:main
ghcr.io/hwkdo/teams-sdk-rest:<git-sha>
```

Manuell starten: GitHub → Actions → **Docker Publish** → **Run workflow**.

### Image pullen

```bash
echo $GITHUB_TOKEN | docker login ghcr.io -u USERNAME --password-stdin
docker pull ghcr.io/hwkdo/teams-sdk-rest:latest
```

Für private Packages: unter GitHub → **Packages** → `teams-sdk-rest` → **Package settings** → Visibility anpassen oder Deploy-Token nutzen.

Der Workflow nutzt `GITHUB_TOKEN` mit `packages: write` — keine zusätzlichen Secrets nötig.

## Umgebungsvariablen

| Variable | Beschreibung |
|---|---|
| `CLIENT_ID` | Azure App Registration Client ID |
| `CLIENT_SECRET` | Client Secret |
| `TENANT_ID` | Azure AD Tenant ID |
| `API_KEY` | Bearer-Token für REST-API (Laravel) |
| `LARAVEL_WEBHOOK_URL` | URL für Event-Weiterleitung (optional) |
| `LARAVEL_WEBHOOK_SECRET` | HMAC-Secret für `X-Teams-Signature` (optional) |
| `WELCOME_MESSAGE` | Willkommensnachricht bei `install.add` (leer = deaktiviert) |
| `HI_REPLY_MESSAGE` | Antwort auf `Hi`/`Hello`/`Hallo` (Pflicht für Teams-App-Validierung) |
| `DATA_DIR` | SQLite-Pfad (Default: `./data`, Docker: `/data`) |
| `PORT` | HTTP-Port (Default: `3978`) |

## Azure-Voraussetzungen

1. **Azure Bot Registration** mit Teams-Kanal
2. Messaging Endpoint: `https://<host>/api/messages`
3. Teams App sideloaden/installieren (Bot muss im Team/User installiert sein)
4. **Graph Application Permissions** (Admin Consent):
   - `User.Read.All`
   - `Team.ReadBasic.All`
   - `Channel.ReadBasic.All`

Ohne Bot-Installation schlagen proaktive Nachrichten fehl — das ist eine Teams-Plattform-Regel.

## Teams Developer Portal — App-Validierung („Hi“-Befehl)

Die Validierung prüft, ob der Bot auf den Manifest-Befehl **Hi** antwortet. Die Antwort erfolgt **direkt im Container** (vor dem Laravel-Webhook), damit das kurze Validierungs-Timeout eingehalten wird.

### Manifest (`commandLists`)

Im Teams-App-Manifest muss der Befehl deklariert sein, z. B.:

```json
"commandLists": [
  {
    "scopes": ["personal"],
    "commands": [
      {
        "title": "Hi",
        "description": "Begrüßung"
      }
    ]
  }
]
```

`bots[].botId` und `webApplicationInfo.id` müssen der Bot-App-ID entsprechen.

### Container-Config

```env
HI_REPLY_MESSAGE=Hallo! Schön, dass du da bist. Ich sende dir Benachrichtigungen aus dem HWKDO Intranet.
```

Leerer Wert = keine Hi-Antwort → Validierung schlägt fehl.

### Nach Deploy prüfen

```bash
docker compose up -d --build
docker logs -f <teams-sdk-rest-container>
```

In Teams `Hi` an den Bot senden. Erwartetes Log:

```
[teams-sdk-rest] inbound event=message ... text=Hi
[teams-sdk-rest] inbound event=message.hi-reply conversationId=...
```

Laravel antwortet **nicht** nochmals auf `Hi` (kein Doppel-Reply).

### Bekannte Plattform-Probleme

Microsoft hat wiederholt gemeldet, dass die Validierung manchmal **keine** `message`-Activity an den Bot sendet, trotzdem aber „Bot reagiert nicht auf Hi“ meldet. Wenn manuell alles funktioniert und `message.hi-reply` im Log erscheint, liegt es oft am Validator — ggf. erneut validieren oder Microsoft-Support.

## REST-API

Alle `/v1/*`-Endpunkte erfordern:

```
Authorization: Bearer <API_KEY>
```

### Health

```
GET /health
```

### Nachrichten senden

```
POST /v1/messages
```

```json
{
  "teamId": "19:...@thread.tacv2",
  "channelId": "19:...@thread.tacv2",
  "text": "Neuer Ticket #1234"
}
```

Ziel-Auflösung (eines von):

- `conversationId` — direkt
- `userAadId` — Lookup aus SQLite
- `teamId` + `channelId` — Lookup aus SQLite

Adaptive Card:

```json
{
  "conversationId": "19:...@thread.v2",
  "card": {
    "type": "AdaptiveCard",
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "version": "1.5",
    "body": [{ "type": "TextBlock", "text": "Hallo!" }]
  }
}
```

### Thread-Antwort

```
POST /v1/messages/reply
```

```json
{
  "conversationId": "19:...@thread.tacv2",
  "messageId": "1234567890",
  "text": "Update zum Thread"
}
```

### Typing-Indikator

```
POST /v1/messages/typing
```

```json
{ "conversationId": "19:...@thread.v2" }
```

### Conversations (Storage)

```
POST /v1/conversations
GET /v1/conversations?limit=50&offset=0
GET /v1/conversations/users/{aadObjectId}
GET /v1/conversations/channels/{teamId}/{channelId}
```

Registrieren (z. B. nach Graph-Installation aus Laravel):

```json
{
  "userAadId": "02bd2b59-d49b-44ce-a709-580a54e1eaf8",
  "conversationId": "19:...@thread.v2",
  "serviceUrl": "https://smba.trafficmanager.net/teams/",
  "tenantId": "..."
}
```

### Graph (Read-only)

```
GET /v1/graph/me
GET /v1/graph/users/{idOrUpn}
GET /v1/graph/users?search=Max
GET /v1/graph/teams
GET /v1/graph/teams/{teamId}/channels
```

## Laravel-Integration

### Ausgehend

```php
use Illuminate\Support\Facades\Http;

Http::withToken(config('teams.api_key'))
    ->post(config('teams.base_url') . '/v1/messages', [
        'teamId' => $teamId,
        'channelId' => $channelId,
        'text' => 'Neuer Ticket: #1234',
    ]);
```

### Eingehend (Webhook)

```php
public function handle(Request $request)
{
    $payload = $request->json()->all();
    $signature = $request->header('X-Teams-Signature');
    $event = $request->header('X-Teams-Event');

    if ($secret = config('teams.webhook_secret')) {
        $expected = 'sha256=' . hash_hmac('sha256', $request->getContent(), $secret);
        abort_unless(hash_equals($expected, $signature ?? ''), 403);
    }

    // $payload['event'], $payload['activity'], $payload['conversationRef']
}
```

Webhook-Payload:

```json
{
  "event": "message",
  "timestamp": "2026-07-06T12:00:00.000Z",
  "activity": { },
  "conversationRef": {
    "conversationId": "...",
    "userAadId": "...",
    "teamId": "...",
    "channelId": "..."
  }
}
```

Weitergeleitete Events: `message`, `mention`, `install.add`, `install.remove`, `conversationUpdate.channelMemberAdded`, `adaptive-card.action`

## Entwicklung

```bash
npm run dev
```

## Lizenz

MIT
