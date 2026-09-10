/**
 * Types for the Navigator chat API (POST /<chat-path>)
 * Source: vendor API spec (v1, protocol version 1)
 */

export type NavigatorRole = 'user' | 'assistant' | 'system';

export interface INavigatorMessage {
  role: NavigatorRole;
  content: string;
}

export type NavigatorRagType = 'hr' | 'la';

export interface INavigatorRequest {
  request_id: string;
  user_id: string;
  timestamp: string; // ISO 8601 UTC
  session: string;
  chat_history: INavigatorMessage[];
  knowledge_tool?: boolean;
  tools?: string[]; // only "web_search" is honoured server-side
  custom_instructions?: string; // <= 2000 chars
  continue_llm_response?: boolean;
  rag_type?: NavigatorRagType; // <= 32 chars, only meaningful with continue_llm_response
  metadata?: Record<string, unknown>; // echoed to logs only, never reaches the model
}

// ---- Streamed event payloads (text/event-stream, one JSON object per "data:" line) ----

export interface INavigatorContentEvent {
  type: 'content';
  data: { content: string };
  source: string;
}

export interface INavigatorToolStartEvent {
  type: 'tool_start';
  tool_data: {
    tool_name: 'query_3m_knowledge' | 'search_the_web';
    status: 'started';
    message: string;
    metadata?: Record<string, unknown>;
  };
  source: string;
}

export interface INavigatorToolProgressEvent {
  type: 'tool_progress';
  tool_data: {
    tool_name: string;
    status: 'running';
    message: string;
    metadata?: { elapsed_ms?: number };
  };
  source: string;
}

export interface INavigatorToolCompleteEvent {
  type: 'tool_complete';
  tool_data: {
    tool_name: string;
    status: 'completed' | 'error'; // 'error' is NOT fatal - the answer still continues
    message: string;
    results_count?: number;
  };
  source: string;
}

export interface IKnowledgeCitation {
  title?: string;
  url: string;
  snippet?: string;
  confidence_score?: number;
}

export interface IWebCitation {
  url: string;
  domain?: string;
  title?: string;
}

export interface INavigatorCitationsEvent {
  type: 'citations';
  citations_data: { citations: Array<IKnowledgeCitation | IWebCitation> };
  source: string;
}

export interface INavigatorFollowupsEvent {
  type: 'followups';
  followups_data: { followups: string[] };
  source: string;
}

export type NavigatorErrorType = 'api_error' | 'stream_timeout';

export interface INavigatorErrorEvent {
  type: 'error';
  error_data: { error: true; error_type: NavigatorErrorType; message: string };
  status_code?: number;
  source: string;
}

export interface INavigatorDoneEvent {
  type: 'done';
  done: true;
  source: string;
}

export type NavigatorStreamEvent =
  | INavigatorContentEvent
  | INavigatorToolStartEvent
  | INavigatorToolProgressEvent
  | INavigatorToolCompleteEvent
  | INavigatorCitationsEvent
  | INavigatorFollowupsEvent
  | INavigatorErrorEvent
  | INavigatorDoneEvent;

export type NavigatorRoutingDomain = 'hr' | 'la';

export interface INavigatorRoutingHandoff {
  domain: NavigatorRoutingDomain;
}

export interface INavigatorStreamHandlers {
  onContent: (delta: string) => void;
  onToolStart?: (evt: INavigatorToolStartEvent) => void;
  onToolProgress?: (evt: INavigatorToolProgressEvent) => void;
  onToolComplete?: (evt: INavigatorToolCompleteEvent) => void;
  onCitations?: (evt: INavigatorCitationsEvent) => void;
  onFollowups?: (followups: string[]) => void;
  /** Fired instead of onContent when the answer resolves to a bare HR/Legal routing keyword. */
  onRoutingDetected?: (handoff: INavigatorRoutingHandoff) => void;
  onError: (
    message: string,
    errorType: NavigatorErrorType | 'http_error' | 'blocked' | 'client_error',
    retryable: boolean
  ) => void;
  onDone: () => void;
}
