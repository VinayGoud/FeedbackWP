# Wiring the real NavigatorService into AiBanner.tsx

## 1. Drop-in files
Replace the mock service with these three (`src/webparts/aiBanner/Services/`):
- `INavigatorTypes.ts`
- `NavigatorStreamParser.ts`
- `NavigatorService.ts`

## 2. Two TODOs before this works
- **`NAVIGATOR_RESOURCE_ID`** in `NavigatorService.ts` — the App ID URI of the
  `aicoe-navigator-api` Azure AD app registration. Get this from onboarding.
  Your SPFx app registration will also need **API permissions** granted (and
  admin-consented) to call that resource, or `aadHttpClientFactory.getClient()`
  will fail.
- **`NAVIGATOR_BASE_URL`** — currently set to the dev URL from the spec
  (`https://navigatordev.ai.3m.com/nav-api`). Swap for prod once issued.

## 3. Replacing `getMockResponse()`
`AiBanner.tsx` currently calls `NavigatorService.getMockResponse(query, siteContext)`
and renders a single response card. The new API is streaming and callback-based:

```ts
const navigatorService = new NavigatorService(this.props.context);

await navigatorService.sendMessage(
  {
    session: this.state.sessionId, // generate once per chat, keep stable (see §4 below)
    chatHistory: this.state.messages, // INavigatorMessage[], oldest first, must end with role:'user'
    knowledgeTool: true, // per BRD default; flip per selected knowledge hub
  },
  {
    onContent: (delta) => this.appendToLastAssistantMessage(delta),
    onToolStart: (evt) => this.setToolStatus(evt.tool_data.message),
    onToolComplete: (evt) => this.setToolStatus(undefined),
    onCitations: (evt) => this.setCitations(evt.citations_data.citations),
    onFollowups: (followups) => this.setFollowupChips(followups),
    onRoutingDetected: (handoff) => this.showSpecialistHandoff(handoff.domain),
    onError: (message, type, retryable) => this.showError(message, retryable),
    onDone: () => this.finishStreamingMessage(),
  }
);
```

This directly satisfies the client's "full multi-turn chat window" ask —
`onContent` deltas append into a running assistant bubble, matching how
Claude/ChatGPT stream.

## 4. Session continuity
`session` is a **vendor-generated** id (§2.2) — you own it, not Navigator.
Generate one GUID per chat when the homepage bar is first opened, keep it in
component state for the life of that conversation, and reuse it on every
`sendMessage` call. This also answers the client's open question:

- If "Continue in Navigator" should carry the conversation over, pass the
  same `session` value as a URL parameter when launching Navigator (assuming
  Navigator's own UI accepts a `session` param — confirm this separately,
  it's outside the scope of this API spec, which only covers third-party
  callers).

## 5. HR/Legal routing keyword (§6)
When `onRoutingDetected` fires:
1. Don't render any of the buffered text (the service already discards it).
2. Show your own UI: a link to the specialist assistant + a "continue in this
   chat" affordance.
3. If the user clicks "continue in this chat", resend the **same** last user
   question with `continueLlmResponse: true` and `ragType` set to the matched
   domain. Do not add the routed reply to `chatHistory` — the service never
   surfaces the bare keyword to you, so this happens naturally as long as you
   only push real `onContent` text into message state.

## 6. Error handling surface
`onError(message, type, retryable)` covers every case in §7:
- `'blocked'` → WAF block, HTML body was never rendered — show a generic
  message, suggest a new chat.
- `'http_error'` → validation (422), throttling (429, already retried with
  backoff internally), or exhausted 5xx retries.
- `'stream_timeout'` → no chunk within 300s.
- `'client_error'` → token acquisition or network failure client-side.

`tool_complete` with `status: 'error'` is **not** passed to `onError` — it's
routed to `onToolComplete` since the assistant still produces an answer; only
surface it as a subtle "couldn't search 3M knowledge" note if you want.

## 7. Still open (not in this spec's scope)
- Link-behavior decision (which of the 4 original options ships) — the API
  itself doesn't dictate this, it's a product/UX call.
- Translation, targeting, analytics, icons — unchanged, still open BRD items.
