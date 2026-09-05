import { WebPartContext } from '@microsoft/sp-webpart-base';
 
/** Props passed from FeedbackWpWebPart.ts into the FeedbackWp React component. */
export interface IFeedbackWpProps {
  context: WebPartContext;
  isActive: boolean;
}
 