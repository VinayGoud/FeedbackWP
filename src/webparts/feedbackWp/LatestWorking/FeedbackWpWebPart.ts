import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'FeedbackWpWebPartStrings';
import FeedbackWp from './components/FeedbackWp';
import { IFeedbackWpProps } from './components/IFeedbackWpProps';

export interface IFeedbackWpWebPartProps {
  description: string;
  isActive: boolean;
  language: string;
}

export default class FeedbackWpWebPart extends BaseClientSideWebPart<IFeedbackWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFeedbackWpProps> = React.createElement(
      FeedbackWp,
      {
        context: this.context,
        isActive: this.properties.isActive,
        language: this.properties.language
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Default to active if the "Show Feedback web part" checkbox was never touched.
    if (this.properties.isActive === undefined) {
      this.properties.isActive = true;
    }

    // Same reasoning as above - default to English if language dropdown was never touched. 
    // Page owner needs to manually change this to language that page actually is.
    if (this.properties.language === undefined) {
      this.properties.language = 'en';
    }

    return super.onInit();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneCheckbox('isActive', {
                  text: 'Show Feedback web part'
                }),
                PropertyPaneDropdown('language', {
                  label: 'Choose language',
                  options: [
                    { key: 'en', text: 'English' },
                    { key: 'es', text: 'Spanish' },
                    { key: 'de', text: 'German' },
                    { key: 'fr', text: 'French' },
                    { key: 'it', text: 'Italian' },
                    { key: 'nl', text: 'Dutch' },
                    { key: 'ja', text: 'Japanese' },
                    { key: 'ko', text: 'Korean' },
                    { key: 'pl', text: 'Polish' },
                    { key: 'pt', text: 'Portuguese (Brazil)' },
                    { key: 'th', text: 'Thai' },
                    { key: 'zh-cn', text: 'Chinese (Simplified)' },
                    { key: 'zh-tw', text: 'Chinese (Traditional)' }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}