import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

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

// Mapping of language code to actual column name in FeedbackPrompts list.
// Keeping all languages here is fine even if column not created yet in some
// environment - getAvailableFields() below will check what actually exists
// on the list before using it in $select, so no risk of 400 error even if
// column is missing.
//
// Most languages just use 2 letter code (fr, de, es etc). Chinese is
// exception - Simplified and Traditional both start with "zh" so had to
// use full code (zh-cn / zh-tw) for those two, otherwise no way to tell
// them apart.
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

// Default list names - used when a web part instance doesn't have the
// property set (older instances added before this feature existed, or
// someone just wants the standard names and doesn't need to reuse this
// webpart with cloned lists).
const DEFAULT_PROMPTS_LIST = 'FeedbackPrompts';
const DEFAULT_RESPONSES_LIST = 'FeedbackResponses';

export class FeedbackService {
  private readonly siteUrl: string;
  private readonly promptsListName: string;
  private readonly responsesListName: string;
  private availableFieldsPromise: Promise<Set<string>> | undefined;

  constructor(
    private context: WebPartContext,
    promptsListName?: string,
    responsesListName?: string
  ) {
    this.siteUrl = context.pageContext.web.absoluteUrl;
    this.promptsListName = promptsListName || DEFAULT_PROMPTS_LIST;
    this.responsesListName = responsesListName || DEFAULT_RESPONSES_LIST;
  }

  // Fetching field list only once and caching it - no point calling this
  // REST endpoint again and again for same webpart instance.
  private async getAvailableFields(): Promise<Set<string>> {
    if (!this.availableFieldsPromise) {
      this.availableFieldsPromise = this.fetchAvailableFields();
    }
    return this.availableFieldsPromise;
  }

  private async fetchAvailableFields(): Promise<Set<string>> {
    try {
      const url = `${this.siteUrl}/_api/web/lists/getByTitle('${this.promptsListName}')/fields?$select=InternalName`;

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
      // Even if this fields check itself fails (network issue, permission
      // issue, whatever) - not going to let that break the whole prompt
      // loading. Just treat it as no translation columns available and
      // move on, prompt will show in English which is still fine.
      return new Set<string>();
    }
  }

  // Gets the currently active prompt from list. If somehow more than one
  // row has Active = true (should not normally happen but just in case),
  // taking the most recently created one - that is why $orderby by Created
  // desc and $top=1.
  //
  // Returns undefined if nothing is active - calling code should handle
  // that and just not show the webpart in that case.
  public async getActivePrompt(language: string): Promise<IFeedbackPrompt | undefined> {
    const availableFields = await this.getAvailableFields();
    const translationColumns = Object.keys(TITLE_TRANSLATION_COLUMNS)
      .map(languageCode => TITLE_TRANSLATION_COLUMNS[languageCode])
      .filter(columnName => availableFields.has(columnName));

    const url =
      `${this.siteUrl}/_api/web/lists/getByTitle('${this.promptsListName}')/items` +
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
      Title: this.resolveLocalizedTitle(item, language),
      Active: item.Active,
      WeekNumber: item.WeekNumber
    };
  }

  // Picks correct Title column based on language passed in (this comes
  // from webpart property "Choose language" dropdown, page owner sets it
  // manually for each translated page). If matching column is empty or
  // does not exist, just falls back to plain Title column (English).
  private resolveLocalizedTitle(item: IRawPromptItem, language: string): string {
    const languageKey = (language || 'en').toLowerCase();

    const columnName = TITLE_TRANSLATION_COLUMNS[languageKey];
    if (columnName) {
      const translatedValue = item[columnName];
      if (typeof translatedValue === 'string' && translatedValue.length > 0) {
        return translatedValue;
      }
    }

    return item.Title;
  }

  // Writes user's response (vote + comment) to FeedbackResponses list.
  // No check for duplicate submission - as per requirement user is allowed
  // to submit again after page refresh, so not blocking that here.
  public async addResponse(payload: IFeedbackResponse): Promise<void> {
    const url = `${this.siteUrl}/_api/web/lists/getByTitle('${this.responsesListName}')/items`;

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
