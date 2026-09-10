import { NavigatorStreamEvent } from './INavigatorTypes';

/**
 * Parses the chat API's text/event-stream body.
 * Spec quirks handled here (§4):
 *  - No "event:" field, no [DONE] sentinel; terminal frame is {"type":"done","done":true}
 *  - Each frame is a single "data: {...}" line followed by a blank line
 *  - Unknown "type" values must be ignored (forward compatibility)
 */
export class NavigatorStreamParser {
  private buffer = '';

  /** Feed a raw chunk of decoded text; returns any complete events found. */
  public push(chunk: string): NavigatorStreamEvent[] {
    this.buffer += chunk;
    const events: NavigatorStreamEvent[] = [];

    // Frames are separated by a blank line. Keep the trailing partial frame in the buffer.
    const frames = this.buffer.split('\n\n');
    this.buffer = frames.pop() ?? '';

    for (const frame of frames) {
      for (const line of frame.split('\n')) {
        if (!line.startsWith('data:')) {
          continue; // ignore everything that isn't a data line
        }
        const jsonText = line.slice('data:'.length).trim();
        if (!jsonText) {
          continue;
        }
        try {
          events.push(JSON.parse(jsonText) as NavigatorStreamEvent);
        } catch {
          // Malformed frame - skip rather than throw ("read defensively").
        }
      }
    }
    return events;
  }
}

const ROUTING_KEYWORDS: Array<{ keyword: string; domain: 'hr' | 'la' }> = [
  { keyword: 'SWITCH_TO_HR_RAG', domain: 'hr' },
  { keyword: 'SWITCH_TO_LA_RAG', domain: 'la' },
];
const MAX_KEYWORD_LEN = Math.max(...ROUTING_KEYWORDS.map((k) => k.keyword.length));

export interface IRoutingBufferResult {
  /** Text safe to render immediately (never contains part of a matched keyword). */
  textToEmit: string;
  /** Set when a full routing keyword was matched. */
  matchedDomain?: 'hr' | 'la';
}

/**
 * Holds back content deltas that could still be the start of a routing keyword
 * (e.g. "SWITCH", "_TO", "_HR", "_RAG" arriving as separate tokens), per §6.
 * Create one instance per in-flight assistant turn.
 */
export class RoutingKeywordBuffer {
  private held = '';

  public feed(delta: string): IRoutingBufferResult {
    this.held += delta;

    for (const { keyword, domain } of ROUTING_KEYWORDS) {
      if (this.held === keyword) {
        this.held = '';
        return { textToEmit: '', matchedDomain: domain };
      }
    }

    const isPrefix = ROUTING_KEYWORDS.some(({ keyword }) => keyword.startsWith(this.held));
    if (isPrefix && this.held.length <= MAX_KEYWORD_LEN) {
      return { textToEmit: '' }; // still could become a keyword - keep holding
    }

    // Not a keyword (or grew past the longest keyword) - release everything held.
    const textToEmit = this.held;
    this.held = '';
    return { textToEmit };
  }

  /** Call when the stream ends (done/error) to flush anything still held mid-keyword. */
  public flushEnd(): string {
    const remaining = this.held;
    this.held = '';
    return remaining;
  }
}
