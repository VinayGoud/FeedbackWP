import * as React from 'react';
import { useState, useRef, useCallback, useEffect } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { NavigatorService } from '../Services/NavigatorService';
import {
  INavigatorMessage,
  IKnowledgeCitation,
  IWebCitation,
  INavigatorRoutingHandoff,
} from '../Services/INavigatorTypes';
import styles from './AiBanner.module.scss';

export interface IAiBannerProps {
  context: WebPartContext;
  /** e.g. current site title, used for the existing greeting/suggestion pill logic */
  siteTitle?: string;
  userDisplayName?: string;
}

interface IChatUIMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  isStreaming?: boolean;
  toolStatus?: string;
  citations?: Array<IKnowledgeCitation | IWebCitation>;
  followups?: string[];
  errorText?: string;
}

function newId(): string {
  return typeof crypto.randomUUID === 'function' ? crypto.randomUUID() : `id_${Date.now()}_${Math.random()}`;
}

function isKnowledgeCitation(c: IKnowledgeCitation | IWebCitation): c is IKnowledgeCitation {
  return typeof (c as IKnowledgeCitation).confidence_score !== 'undefined' || typeof (c as IKnowledgeCitation).snippet !== 'undefined';
}

export const AiBanner: React.FC<IAiBannerProps> = (props) => {
  const [isOpen, setIsOpen] = useState(false);
  const [inputValue, setInputValue] = useState('');
  const [messages, setMessages] = useState<IChatUIMessage[]>([]);
  const [routingHandoff, setRoutingHandoff] = useState<INavigatorRoutingHandoff | undefined>(undefined);
  const [isSending, setIsSending] = useState(false);
  const [bannerError, setBannerError] = useState<string | undefined>(undefined);

  const sessionIdRef = useRef<string>(newId());
  const serviceRef = useRef<NavigatorService>(new NavigatorService(props.context));
  const abortRef = useRef<AbortController | undefined>(undefined);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: 'smooth' });
  }, [messages]);

  useEffect(() => {
    // Cancel any in-flight stream if the component unmounts.
    return () => abortRef.current?.abort();
  }, []);

  const toChatHistory = useCallback((uiMessages: IChatUIMessage[]): INavigatorMessage[] => {
    return uiMessages
      .filter((m) => m.content && !m.errorText)
      .map((m) => ({ role: m.role, content: m.content }));
  }, []);

  const sendQuery = useCallback(
    (query: string, opts?: { continueLlmResponse?: boolean; ragType?: 'hr' | 'la' }) => {
      setBannerError(undefined);
      setRoutingHandoff(undefined);

      const userMsg: IChatUIMessage = { id: newId(), role: 'user', content: query };
      const assistantMsg: IChatUIMessage = { id: newId(), role: 'assistant', content: '', isStreaming: true };

      setMessages((prev) => {
        const next = [...prev, userMsg, assistantMsg];
        // Fire the request using the history as of this update.
        const chatHistory = toChatHistory(opts?.continueLlmResponse ? prev.concat(userMsg) : next.filter((m) => m.id !== assistantMsg.id));

        const controller = new AbortController();
        abortRef.current = controller;
        setIsSending(true);

        void serviceRef.current
          .sendMessage(
            {
              session: sessionIdRef.current,
              chatHistory,
              knowledgeTool: true, // BRD default; wire to a knowledge-hub selector if/when that ships
              continueLlmResponse: opts?.continueLlmResponse,
              ragType: opts?.ragType,
            },
            {
              onContent: (delta) => {
                setMessages((cur) =>
                  cur.map((m) => (m.id === assistantMsg.id ? { ...m, content: m.content + delta } : m))
                );
              },
              onToolStart: (evt) => {
                setMessages((cur) =>
                  cur.map((m) => (m.id === assistantMsg.id ? { ...m, toolStatus: evt.tool_data.message } : m))
                );
              },
              onToolProgress: (evt) => {
                setMessages((cur) =>
                  cur.map((m) => (m.id === assistantMsg.id ? { ...m, toolStatus: evt.tool_data.message } : m))
                );
              },
              onToolComplete: () => {
                setMessages((cur) => cur.map((m) => (m.id === assistantMsg.id ? { ...m, toolStatus: undefined } : m)));
              },
              onCitations: (evt) => {
                setMessages((cur) =>
                  cur.map((m) => (m.id === assistantMsg.id ? { ...m, citations: evt.citations_data.citations } : m))
                );
              },
              onFollowups: (followups) => {
                setMessages((cur) => cur.map((m) => (m.id === assistantMsg.id ? { ...m, followups } : m)));
              },
              onRoutingDetected: (handoff) => {
                // This turn resolves to a specialist handoff, not a normal answer -
                // drop the empty streaming bubble and show the handoff card instead.
                setMessages((cur) => cur.filter((m) => m.id !== assistantMsg.id));
                setRoutingHandoff(handoff);
              },
              onError: (message, _errorType, retryable) => {
                setMessages((cur) =>
                  cur.map((m) =>
                    m.id === assistantMsg.id ? { ...m, isStreaming: false, errorText: message } : m
                  )
                );
                setBannerError(retryable ? `${message} You can try again.` : message);
              },
              onDone: () => {
                setMessages((cur) => cur.map((m) => (m.id === assistantMsg.id ? { ...m, isStreaming: false } : m)));
              },
            },
            controller.signal
          )
          .finally(() => setIsSending(false));

        return next;
      });
    },
    [toChatHistory]
  );

  const handleSubmit = useCallback(() => {
    const query = inputValue.trim();
    if (!query || isSending) {
      return;
    }
    if (!isOpen) {
      setIsOpen(true);
    }
    setInputValue('');
    sendQuery(query);
  }, [inputValue, isSending, isOpen, sendQuery]);

  const handleFollowupClick = useCallback(
    (text: string) => {
      if (isSending) {
        return;
      }
      sendQuery(text);
    },
    [isSending, sendQuery]
  );

  const handleContinueInThisChat = useCallback(() => {
    const lastUser = [...messages].reverse().find((m) => m.role === 'user');
    if (!lastUser) {
      return;
    }
    sendQuery(lastUser.content, { continueLlmResponse: true, ragType: routingHandoff?.domain });
  }, [messages, routingHandoff, sendQuery]);

  const handleClose = useCallback(() => {
    abortRef.current?.abort();
    setIsOpen(false);
    setMessages([]);
    setRoutingHandoff(undefined);
    setBannerError(undefined);
    sessionIdRef.current = newId(); // fresh session for the next chat
  }, []);

  const handleContinueInNavigator = useCallback(() => {
    // Link Behavior decision is still open (BRD) - this assumes option 4 style
    // and passes the session id along; confirm the real param name with Navigator/Product.
    const url = `/sites/navigator?session=${encodeURIComponent(sessionIdRef.current)}`; // TODO: replace with real destination once Link Behavior is confirmed
    window.open(url, '_blank', 'noopener');
  }, []);

  return (
    <div className={styles.aiBanner}>
      {/* Existing greeting + pill search bar markup goes here unchanged - keep whatever
          time-based greeting / "Try: [suggestion]" pill logic already exists, just point
          its onSubmit at handleSubmit and its value/onChange at inputValue/setInputValue. */}

      <div className={styles.searchRow}>
        <input
          className={styles.searchInput}
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && handleSubmit()}
          placeholder="Ask anything..."
        />
        <button className={styles.submitButton} onClick={handleSubmit} disabled={isSending}>
          {isSending ? '...' : '→'}
        </button>
      </div>

      {isOpen && (
        <div className={styles.chatPanel}>
          <div className={styles.chatPanelHeader}>
            <span>Navigator</span>
            <button className={styles.closeButton} onClick={handleClose} aria-label="Close">
              ×
            </button>
          </div>

          <div className={styles.messageList} ref={scrollRef}>
            {messages.map((m) => (
              <div key={m.id} className={m.role === 'user' ? styles.userBubble : styles.assistantBubble}>
                <div>{m.content}</div>

                {m.toolStatus && <div className={styles.toolStatus}>{m.toolStatus}</div>}

                {m.errorText && <div className={styles.errorText}>{m.errorText}</div>}

                {m.citations && m.citations.length > 0 && (
                  <ul className={styles.citationsList}>
                    {m.citations
                      .filter((c) => !!c.url)
                      .map((c, i) => (
                        <li key={i}>
                          <a href={c.url} target="_blank" rel="noopener noreferrer">
                            {c.title || (isKnowledgeCitation(c) ? c.url : c.domain) || c.url}
                          </a>
                        </li>
                      ))}
                  </ul>
                )}

                {m.followups && m.followups.length > 0 && (
                  <div className={styles.followupChips}>
                    {m.followups.map((f, i) => (
                      <button key={i} className={styles.followupChip} onClick={() => handleFollowupClick(f)}>
                        {f}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            ))}

            {routingHandoff && (
              <div className={styles.routingHandoffCard}>
                <p>
                  This question is best answered by the {routingHandoff.domain === 'hr' ? 'HR/Workday' : 'Legal & Compliance'}{' '}
                  assistant.
                </p>
                <div className={styles.routingHandoffActions}>
                  <a href={`/sites/${routingHandoff.domain === 'hr' ? 'hr-assistant' : 'legal-assistant'}`} target="_blank" rel="noopener noreferrer">
                    Open specialist assistant
                  </a>
                  <button onClick={handleContinueInThisChat}>Continue in this chat</button>
                </div>
              </div>
            )}
          </div>

          {bannerError && <div className={styles.bannerError}>{bannerError}</div>}

          <div className={styles.chatPanelFooter}>
            <input
              className={styles.chatInput}
              value={inputValue}
              onChange={(e) => setInputValue(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && handleSubmit()}
              placeholder="Ask a follow-up..."
              disabled={isSending}
            />
            <button onClick={handleSubmit} disabled={isSending}>
              Send
            </button>
            <button className={styles.continueInNavigatorButton} onClick={handleContinueInNavigator}>
              Continue in Navigator
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default AiBanner;
