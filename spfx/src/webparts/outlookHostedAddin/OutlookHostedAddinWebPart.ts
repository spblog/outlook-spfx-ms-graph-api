import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'OutlookHostedAddinWebPartStrings';
import { OutlookHostedAddin } from './components/OutlookHostedAddin';
import { IOutlookHostedAddinProps } from './components/IOutlookHostedAddinProps';

export interface IOutlookHostedAddinWebPartProps {
  description: string;
}

export default class OutlookHostedAddinWebPart extends BaseClientSideWebPart<IOutlookHostedAddinWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IOutlookHostedAddinProps> = React.createElement(
      OutlookHostedAddin,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
