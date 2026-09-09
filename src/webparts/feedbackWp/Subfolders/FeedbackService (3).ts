import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * REST access layer for the Feedback web part. Talks to two SharePoint lists
 * on the current site (resolved dynamically via context.pageContext.web.absoluteUrl,
 * never a hardcoded site URL):
 *   - FeedbackPrompts: Title, Active, WeekNumber
 *   - FeedbackResponses: Title, Comments, Like, PromptId, CommentTranslation (business-filled, unused here)
 */

export interface IFeedbackPrompt {
  Id: number;
  Title: string;
  Active: boolean;
  WeekNumber: number;
}

export interface IFeedbackResponse {
  Title: string;
  Comments: string;
  Like: boolean;
  PromptId: number;
}

interface ISPListItemsResponse<T> {
  value: T[];
}

/**
 * Maps a language key to the matching translation column on FeedbackPrompts.
 * All languages can safely stay uncommented — getAvailableFields() below
 * checks which columns actually exist before building $select, so a column
 * that's missing (not yet created, or deleted) in a given environment is
 * silently skipped instead of causing a 400 for the whole request.
 *
 * CONFIRMED against the real list: es, de. The rest were added along with
 * their matching SharePoint columns — verify column names match exactly if
 * any of these languages don't show up as expected.
 *
 * Keys are lowercase. Most languages key on just the 2-letter language code
 * (e.g. 'fr' matches both fr-FR and fr-CA). Chinese is the one exception —
 * Simplified and Traditional share the same 2-letter prefix ('zh'), so those
 * two entries use the full culture code instead ('zh-cn' / 'zh-tw').
 */
const TITLE_TRANSLATION_COLUMNS: { [languageKey: string]: string } = {
  es: 'Title_Spanish',
  de: 'Title_German',
  'zh-cn': 'Title_Chinese_Simplified',
  'zh-tw': 'Title_Chinese_Traditional',
  nl: 'Title_Dutch',
  fr: 'Title_French',
  it: 'Title_Italian',
  ja: 'Title_Japanese',
  ko: 'Title_Korean',
  pl: 'Title_Polish',
  pt: 'Title_Portuguese',
  th: 'Title_Thai'
};

interface IContextInfoResponse {
  FormDigestValue: string;
}

interface IRawPromptItem {
  Id: number;
  Title: string;
  Active: boolean;
  WeekNumber: number;
  [translationColumn: string]: number | string | boolean | undefined;
}

const PROMPTS_LIST = 'FeedbackPrompts';
const RESPONSES_LIST = 'FeedbackResponses';

export class FeedbackService {
  private readonly siteUrl: string;
  private availableFieldsPromise: Promise<Set<string>> | undefined;

  constructor(private context: WebPartContext) {
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  /**
   * Returns the internal names of every field that actually exists on
   * FeedbackPrompts, fetched once and cached for this service instance.
   * This is what makes it safe to keep every language uncommented above —
   * a column that doesn't exist (never created, or since deleted) is simply
   * absent from this set, so it never makes it into $select and never
   * triggers a 400.
   */
  private async getAvailableFields(): Promise<Set<string>> {
    if (!this.availableFieldsPromise) {
      this.availableFieldsPromise = this.fetchAvailableFields();
    }
    return this.availableFieldsPromise;
  }

  private async fetchAvailableFields(): Promise<Set<string>> {
    try {
      const url = `${this.siteUrl}/_api/web/lists/getByTitle('${PROMPTS_LIST}')/fields?$select=InternalName`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return new Set<string>();
      }

      const data = await response.json() as ISPListItemsResponse<{ InternalName: string }>;
      return new Set(data.value.map(field => field.InternalName));
    } catch {
      // If the fields request itself throws (network error, permissions,
      // etc.) — not just a non-ok response — degrade to "no translation
      // columns available" rather than letting this break the whole prompt
      // load. The prompt should still show in English even if this schema
      // check can't run at all.
      return new Set<string>();
    }
  }

  /**
   * Returns the active prompt. If more than one row has Active = true,
   * the most recently created one wins ($orderby=Created desc, $top=1).
   * Returns undefined if no prompt is currently active.
   */
  public async getActivePrompt(): Promise<IFeedbackPrompt | undefined> {
    const availableFields = await this.getAvailableFields();
    const translationColumns = Object.keys(TITLE_TRANSLATION_COLUMNS)
      .map(languageCode => TITLE_TRANSLATION_COLUMNS[languageCode])
      .filter(columnName => availableFields.has(columnName));

    const url =
      `${this.siteUrl}/_api/web/lists/getByTitle('${PROMPTS_LIST}')/items` +
      `?$select=Id,Title,${translationColumns.join(',')},Active,WeekNumber,Created` +
      `&$filter=Active eq 1` +
      `&$orderby=Created desc` +
      `&$top=1`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`getActivePrompt failed (${response.status}): ${errorText}`);
    }

    const data = await response.json() as ISPListItemsResponse<IRawPromptItem>;
    const items = data.value;

    if (!items || items.length === 0) {
      return undefined;
    }

    const item = items[0];
    return {
      Id: item.Id,
      Title: this.resolveLocalizedTitle(item),
      Active: item.Active,
      WeekNumber: item.WeekNumber
    };
  }

  /**
   * Picks the prompt text matching the current user's SharePoint UI language,
   * via the TITLE_TRANSLATION_COLUMNS map above. Falls back to the default
   * Title column if there's no mapped column for the user's language, or if
   * the matching column is empty (not yet translated).
   */
  private resolveLocalizedTitle(item: IRawPromptItem): string {
    const languageKey = this.getLanguageKey(this.context.pageContext.cultureInfo.currentUICultureName);

    const columnName = TITLE_TRANSLATION_COLUMNS[languageKey];
    if (columnName) {
      const translatedValue = item[columnName];
      if (typeof translatedValue === 'string' && translatedValue.length > 0) {
        return translatedValue;
      }
    }

    return item.Title;
  }

  private getLanguageKey(cultureName: string): string {
    const lower = cultureName.toLowerCase();
    if (lower === 'zh-cn' || lower === 'zh-tw') {
      return lower;
    }
    return lower.split('-')[0];
  }

  /** Writes every submission — no duplicate blocking, resubmission is allowed by design. */
  public async addResponse(payload: IFeedbackResponse): Promise<void> {
    const url = `${this.siteUrl}/_api/web/lists/getByTitle('${RESPONSES_LIST}')/items`;

    const digest = await this.getRequestDigest();

    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
          'X-RequestDigest': digest
        },
        body: JSON.stringify(payload)
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`addResponse failed (${response.status}): ${errorText}`);
    }
  }

  private async getRequestDigest(): Promise<string> {
    const url = `${this.siteUrl}/_api/contextinfo`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    const data = await response.json() as IContextInfoResponse;
    return data.FormDigestValue;
  }
}
