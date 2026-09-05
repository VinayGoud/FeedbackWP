import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'FeedbackWpWebPartStrings';
import FeedbackWp from './components/FeedbackWp';
import { IFeedbackWpProps } from './components/IFeedbackWpProps';

export interface IFeedbackWpWebPartProps {
  description: string;
  isActive: boolean;
}

export default class FeedbackWpWebPart extends BaseClientSideWebPart<IFeedbackWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFeedbackWpProps> = React.createElement(
      FeedbackWp,
      {
        context: this.context,
        isActive: this.properties.isActive
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Default the toggle to "on" for any instance that hasn't set it yet
    // (e.g. existing web part instances added to a page before this property existed).
    if (this.properties.isActive === undefined) {
      this.properties.isActive = true;
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}