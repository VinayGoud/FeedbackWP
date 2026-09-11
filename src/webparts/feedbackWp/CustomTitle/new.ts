import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import * as strings from 'NewsStylingCustomizerApplicationCustomizerStrings';

require('./assets/newsStyles.css');

const LOG_SOURCE: string = 'NewsStylingCustomizerApplicationCustomizer';

export interface INewsStylingCustomizerApplicationCustomizerProperties {}

export default class NewsStylingCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<INewsStylingCustomizerApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Capture click BEFORE SharePoint SPA navigation handler runs
    document.addEventListener(
      'click',
      this._handleDocumentClick,
      true
    );

    // Remove temporary hiding class after navigation completes
    this.context.application.navigatedEvent.add(
      this,
      this._onNavigated
    );

    return Promise.resolve();
  }

  private _handleDocumentClick = (event: MouseEvent): void => {
    const target = event.target as HTMLElement;

    if (!target) {
      return;
    }

    const link = target.closest('a') as HTMLAnchorElement | null;

    if (!link) {
      return;
    }

    // Check whether clicked link is inside the News grid
    const newsGrid = link.closest(
      'div[data-automation-id="gridNewsLayout"]'
    );

    if (!newsGrid) {
      return;
    }

    // Only normal left-click navigation
    if (
      event.button !== 0 ||
      event.ctrlKey ||
      event.metaKey ||
      event.shiftKey ||
      event.altKey
    ) {
      return;
    }

    // Hide stale grid immediately
    document.body.classList.add('news-navigation-in-progress');

    // Safety fallback in case navigation fails
    window.setTimeout(() => {
      document.body.classList.remove('news-navigation-in-progress');
    }, 5000);
  };

  private _onNavigated = (): void => {
    document.body.classList.remove('news-navigation-in-progress');
  };

  @override
  public onDispose(): void {
    document.removeEventListener(
      'click',
      this._handleDocumentClick,
      true
    );

    this.context.application.navigatedEvent.remove(
      this,
      this._onNavigated
    );

    document.body.classList.remove('news-navigation-in-progress');
  }
}
