import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import {
  INavigatorMessage,
  INavigatorRequest,
  INavigatorStreamHandlers,
  NavigatorStreamEvent,
  NavigatorRagType,
} from './INavigatorTypes';
import { NavigatorStreamParser, RoutingKeywordBuffer } from './NavigatorStreamParser';

// ---- Config - internal values intentionally omitted, fill in locally after cloning ----
// NAVIGATOR_BASE_URL: internal Navigator API base URL (dev/prod) - do not commit
const NAVIGATOR_BASE_URL = ''; // e.g. https://<navigator-host>/nav-api
const NAVIGATOR_PATH = '/3mgo/chat';
/**
 * App ID URI (audience) of the aicoe-navigator-api Azure AD app registration.
 * SPFx's AadHttpClientFactory uses this to silently acquire a token for the
 * signed-in user - it's a per-user token, not a static secret, so calling it
 * directly from SPFx (browser) code is safe. No backend proxy needed for auth.
 * NAVIGATOR_RESOURCE_ID: internal Azure AD App ID URI - do not commit
 */
const NAVIGATOR_RESOURCE_ID = ''; // e.g. api://<navigator-api-app-id-uri>

const HISTORY_MAX_MESSAGES = 30;
const HISTORY_MAX_CHARS_PER_MESSAGE = 12000;
const HISTORY_MAX_TOTAL_CHARS = 60000;

const STREAM_INACTIVITY_TIMEOUT_MS = 300_000; // 300s - matches server-side stream_timeout (§8)
const MAX_RETRIES = 1; // spec: "retry once" for 401/403 JSON and 5xx

export interface ISendMessageOptions {
  session: string;
  chatHistory: INavigatorMessage[];
  knowledgeTool?: boolean;
  tools?: string[]; // only "web_search" is honoured server-side
  customInstructions?: string;
  continueLlmResponse?: boolean;
  ragType?: NavigatorRagType;
}

export class NavigatorService {
  constructor(private context: WebPartContext) {}

  /**
   * Sends one chat turn and streams the response back through `handlers`.
   * Resolves once the stream reaches `done` or a terminal `error`.
   */
  public async sendMessage(
    options: ISendMessageOptions,
    handlers: INavigatorStreamHandlers,
    signal?: AbortSignal
  ): Promise<void> {
    const requestId = this.generateRequestId();
    const body = this.buildRequestBody(requestId, options);
    await this.postWithRetry(body, handlers, signal, 0);
  }

  // ---------------------------------------------------------------------

  private async postWithRetry(
    body: INavigatorRequest,
    handlers: INavigatorStreamHandlers,
    signal: AbortSignal | undefined,
    attempt: number
  ): Promise<void> {
    let client: AadHttpClient;
    try {
      client = await this.context.aadHttpClientFactory.getClient(NAVIGATOR_RESOURCE_ID);
    } catch {
      handlers.onError('Could not acquire an access token for Navigator.', 'client_error', false);
      return;
    }

    let response: HttpClientResponse;
    try {
      response = await client.fetch(NAVIGATOR_BASE_URL + NAVIGATOR_PATH, AadHttpClient.configurations.v1, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'text/event-stream',
        },
        body: JSON.stringify(body),
        signal,
      });
    } catch {
      if (attempt < MAX_RETRIES) {
        return this.postWithRetry({ ...body, request_id: this.generateRequestId() }, handlers, signal, attempt + 1);
      }
      handlers.onError('Network error contacting Navigator.', 'client_error', true);
      return;
    }

    // --- Pre-stream HTTP status handling (§7.1) ---
    if (!response.ok) {
      const contentType = response.headers.get('content-type') || '';

      if ((response.status === 401 || response.status === 403) && contentType.includes('json')) {
        // Missing/expired/wrong-audience token - refresh and retry once.
        if (attempt < MAX_RETRIES) {
          return this.postWithRetry({ ...body, request_id: this.generateRequestId() }, handlers, signal, attempt + 1);
        }
        handlers.onError('Authentication failed after retry.', 'http_error', false);
        return;
      }

      if (response.status === 403) {
        // HTML body = blocked by the gateway WAF - never render the HTML.
        handlers.onError('Request was blocked. Please start a new chat.', 'blocked', false);
        return;
      }

      if (response.status === 422) {
        const detail = await this.safeReadJson(response);
        handlers.onError(`Request validation failed: ${this.summarizeValidationError(detail)}`, 'http_error', false);
        return;
      }

      if (response.status === 429) {
        const retryAfterSec = Number(response.headers.get('retry-after') ?? '1');
        if (attempt < MAX_RETRIES + 3) {
          await this.delay(Math.max(retryAfterSec, 1) * 1000 * Math.pow(2, attempt));
          return this.postWithRetry(body, handlers, signal, attempt + 1);
        }
        handlers.onError('Navigator is throttling requests. Please try again shortly.', 'http_error', true);
        return;
      }

      if (response.status >= 500) {
        if (attempt < MAX_RETRIES) {
          return this.postWithRetry({ ...body, request_id: this.generateRequestId() }, handlers, signal, attempt + 1);
        }
        handlers.onError('Navigator is temporarily unavailable.', 'http_error', true);
        return;
      }

      handlers.onError(`Unexpected error (HTTP ${response.status}).`, 'http_error', false);
      return;
    }

    await this.consumeStream(response, handlers);
  }

  private async consumeStream(response: HttpClientResponse, handlers: INavigatorStreamHandlers): Promise<void> {
    const body = (response as unknown as { body?: ReadableStream<Uint8Array> }).body;
    if (!body) {
      handlers.onError('Streaming is not supported by this browser/response.', 'client_error', false);
      return;
    }

    const reader = body.getReader();
    const decoder = new TextDecoder('utf-8');
    const parser = new NavigatorStreamParser();
    const routingBuffer = new RoutingKeywordBuffer();
    let sawTerminalFrame = false;
    let lastChunkAt = Date.now();

    const inactivityTimer = setInterval(() => {
      if (Date.now() - lastChunkAt > STREAM_INACTIVITY_TIMEOUT_MS) {
        clearInterval(inactivityTimer);
        void reader.cancel();
        if (!sawTerminalFrame) {
          handlers.onError('Connection timed out.', 'stream_timeout', true);
        }
      }
    }, 5000);

    try {
      // eslint-disable-next-line no-constant-condition
      while (true) {
        const { value, done } = await reader.read();
        if (done) {
          break;
        }
        lastChunkAt = Date.now();
        const events = parser.push(decoder.decode(value, { stream: true }));

        for (const evt of events) {
          this.dispatchEvent(evt, handlers, routingBuffer);
          if (evt.type === 'done' || evt.type === 'error') {
            sawTerminalFrame = true;
          }
        }
        if (sawTerminalFrame) {
          break;
        }
      }
    } finally {
      clearInterval(inactivityTimer);
    }

    // Flush anything still held back mid-keyword if the stream ended without a full match.
    const leftover = routingBuffer.flushEnd();
    if (leftover) {
      handlers.onContent(leftover);
    }
    if (!sawTerminalFrame) {
      handlers.onError('Stream ended unexpectedly.', 'stream_timeout', true);
    }
  }

  private dispatchEvent(
    evt: NavigatorStreamEvent,
    handlers: INavigatorStreamHandlers,
    routingBuffer: RoutingKeywordBuffer
  ): void {
    switch (evt.type) {
      case 'content': {
        const { textToEmit, matchedDomain } = routingBuffer.feed(evt.data.content);
        if (matchedDomain) {
          handlers.onRoutingDetected?.({ domain: matchedDomain });
        } else if (textToEmit) {
          handlers.onContent(textToEmit);
        }
        break;
      }
      case 'tool_start':
        handlers.onToolStart?.(evt);
        break;
      case 'tool_progress':
        handlers.onToolProgress?.(evt);
        break;
      case 'tool_complete':
        // status "error" here is NOT fatal - the assistant still answers (§4.4).
        handlers.onToolComplete?.(evt);
        break;
      case 'citations':
        handlers.onCitations?.(evt);
        break;
      case 'followups':
        handlers.onFollowups?.(evt.followups_data.followups);
        break;
      case 'error':
        handlers.onError(
          evt.error_data.message || 'Navigator returned an error.',
          evt.error_data.error_type,
          evt.error_data.error_type === 'stream_timeout'
        );
        break;
      case 'done':
        handlers.onDone();
        break;
      default:
        // Unknown/new event type - ignore per spec.
        break;
    }
  }

  // ---------------------------------------------------------------------

  private buildRequestBody(requestId: string, options: ISendMessageOptions): INavigatorRequest {
    return {
      request_id: requestId,
      user_id: this.context.pageContext.user.email || this.context.pageContext.user.loginName,
      timestamp: new Date().toISOString(),
      session: options.session,
      chat_history: this.capHistory(options.chatHistory),
      knowledge_tool: options.knowledgeTool ?? false,
      tools: options.tools,
      custom_instructions: options.customInstructions?.slice(0, 2000),
      continue_llm_response: options.continueLlmResponse ?? false,
      rag_type: options.ragType,
    };
  }

  /** Enforces the recommended history cap: last 30 messages / 12k chars per message / 60k total (§8). */
  private capHistory(history: INavigatorMessage[]): INavigatorMessage[] {
    const trimmed = history
      .slice(-HISTORY_MAX_MESSAGES)
      .map((m) => ({ ...m, content: m.content.slice(0, HISTORY_MAX_CHARS_PER_MESSAGE) }));

    let totalChars = trimmed.reduce((sum, m) => sum + m.content.length, 0);
    while (totalChars > HISTORY_MAX_TOTAL_CHARS && trimmed.length > 1) {
      const removed = trimmed.shift();
      totalChars -= removed?.content.length ?? 0;
    }

    // Server requires the last message to have role "user" (422 otherwise, §7.1 common causes).
    if (trimmed.length === 0 || trimmed[trimmed.length - 1].role !== 'user') {
      throw new Error('chat_history must end with a user message.');
    }
    return trimmed;
  }

  /** request_id must match ^[a-zA-Z0-9\-_]+$, 1-2048 chars - no dots/colons/slashes/plus (§2.2). */
  private generateRequestId(): string {
    const rand = Array.from(crypto.getRandomValues(new Uint8Array(16)))
      .map((b) => b.toString(16).padStart(2, '0'))
      .join('');
    return `vendor_${rand}`;
  }

  private async safeReadJson(response: HttpClientResponse): Promise<unknown> {
    try {
      return await response.json();
    } catch {
      return undefined;
    }
  }

  private summarizeValidationError(detail: unknown): string {
    const d = detail as { detail?: Array<{ msg?: string }> } | undefined;
    return d?.detail?.[0]?.msg ?? 'invalid request payload';
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
