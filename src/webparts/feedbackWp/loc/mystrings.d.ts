declare interface IFeedbackWpWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  YesButtonLabel: string;
  NoButtonLabel: string;
  CommentPlaceholder: string;
  SubmitButtonLabel: string;
  SubmittingButtonLabel: string;
  CancelButtonLabel: string;
  ThankYouTitle: string;
  ThankYouSubtitle: string;
}

declare module 'FeedbackWpWebPartStrings' {
  const strings: IFeedbackWpWebPartStrings;
  export = strings;
}
