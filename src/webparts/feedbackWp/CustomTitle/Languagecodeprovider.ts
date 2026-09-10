// All "See All" (and default title) translations live right here in code -
// no separate JSON files to manage, matching the same simple pattern used
// for FeedbackStrings.ts on the Feedback web part.
//
// Draft translations - not reviewed by native speakers yet, same caution
// as the Feedback web part's translations. Get these checked before they
// ship to real users, especially Japanese/Korean/Thai/Chinese.

export interface IGlobalnewsTitleStrings {
  DefaultTitleText: string;
  DefaultSeeAllText: string;
}

const STRINGS: { [languageKey: string]: IGlobalnewsTitleStrings } = {
  'en-us': {
    DefaultTitleText: 'Company News',
    DefaultSeeAllText: 'See All'
  },
  'es-es': {
    DefaultTitleText: 'Noticias de la empresa',
    DefaultSeeAllText: 'Ver todo'
  },
  'de-de': {
    DefaultTitleText: 'Unternehmensnachrichten',
    DefaultSeeAllText: 'Alle anzeigen'
  },
  'fr-fr': {
    DefaultTitleText: "Actualités de l'entreprise",
    DefaultSeeAllText: 'Voir tout'
  },
  'it-it': {
    DefaultTitleText: 'Notizie aziendali',
    DefaultSeeAllText: 'Vedi tutto'
  },
  'nl-nl': {
    DefaultTitleText: 'Bedrijfsnieuws',
    DefaultSeeAllText: 'Alles bekijken'
  },
  'ja-jp': {
    DefaultTitleText: '会社のニュース',
    DefaultSeeAllText: 'すべて表示'
  },
  'ko-kr': {
    DefaultTitleText: '회사 소식',
    DefaultSeeAllText: '모두 보기'
  },
  'pl-pl': {
    DefaultTitleText: 'Wiadomości firmowe',
    DefaultSeeAllText: 'Zobacz wszystko'
  },
  'pt-br': {
    DefaultTitleText: 'Notícias da empresa',
    DefaultSeeAllText: 'Ver tudo'
  },
  'th-th': {
    DefaultTitleText: 'ข่าวสารบริษัท',
    DefaultSeeAllText: 'ดูทั้งหมด'
  },
  'zh-cn': {
    DefaultTitleText: '公司新闻',
    DefaultSeeAllText: '查看全部'
  },
  'zh-tw': {
    DefaultTitleText: '公司新聞',
    DefaultSeeAllText: '查看全部'
  }
};

export class LanguageCodeProvider {

  private static supportedLanguages: string[] = Object.keys(STRINGS);

  public static getCurrentLanguage(languageOverride?: string): string {
    // Manual override wins if the web part property (dropdown) is set -
    // this is what lets a page owner explicitly pick the language instead
    // of relying on the URL always containing the right segment.
    if (languageOverride) {
      const normalizedOverride = languageOverride.toLowerCase();
      if (LanguageCodeProvider.supportedLanguages.indexOf(normalizedOverride) !== -1) {
        return normalizedOverride;
      }
    }

    const pathSegments: string[] = window.location.pathname
      .toLowerCase()
      .split('/')
      .filter((segment: string) => segment.length > 0);

    for (const lang of LanguageCodeProvider.supportedLanguages) {
      const langCode: string = lang.split('-')[0];
      if (pathSegments.indexOf(langCode) !== -1) {
        return lang;
      }
    }

    return 'en-us';
  }

  public static getStrings(languageOverride?: string): IGlobalnewsTitleStrings {
    const lang: string = LanguageCodeProvider.getCurrentLanguage(languageOverride);
    return STRINGS[lang] || STRINGS['en-us'];
  }
}
