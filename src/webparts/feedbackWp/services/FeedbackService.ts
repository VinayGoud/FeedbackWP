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
 * Add one entry per language here — no other code needs to change when a new
 * language is added, as long as the column exists and is included in the
 * $select below.
 *
 * CONFIRMED against the real list: es, de.
 *
 * Keys are lowercase. Most languages key on just the 2-letter language code
 * (e.g. 'fr' matches both fr-FR and fr-CA). Chinese is the one exception —
 * Simplified and Traditional share the same 2-letter prefix ('zh'), so those
 * two entries use the full culture code instead ('zh-cn' / 'zh-tw').
 */
const TITLE_TRANSLATION_COLUMNS: { [languageKey: string]: string } = {
  es: 'Title_Spanish',
  de: 'Title_German'
  // Add each of these back ONE AT A TIME, only after confirming the column
  // missing columns. Adding all of these before the schema rollout is done
  // everywhere would break the web part in any environment still missing one.
  // 'zh-cn': 'Title_Chinese_Simplified',
  // 'zh-tw': 'Title_Chinese_Traditional',
  // nl: 'Title_Dutch',
  // fr: 'Title_French',
  // it: 'Title_Italian',
  // ja: 'Title_Japanese',
  // ko: 'Title_Korean',
  // pl: 'Title_Polish',
  // pt: 'Title_Portuguese',
  // th: 'Title_Thai'
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

  constructor(private context: WebPartContext) {
    this.siteUrl = context.pageContext.web.absoluteUrl;
  }

  public async getActivePrompt(): Promise<IFeedbackPrompt | undefined> {
    const translationColumns = Object.keys(TITLE_TRANSLATION_COLUMNS)
      .map(languageCode => TITLE_TRANSLATION_COLUMNS[languageCode]);

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