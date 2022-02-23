import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HomeWebPartStrings';
import Home from './components/Home';
import { IHomeProps } from './components/IHomeProps';
import { SiteUrl } from './common/Constants';

export interface IHomeWebPartProps {
  description: string;
}

export default class HomeWebPart extends BaseClientSideWebPart<IHomeWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHomeProps> = React.createElement(
      Home,
      {
        description: this.properties.description,
        context: this.context,
        // passing siteUrl here for mutlti tenant.
        // siteUrl: this.context.pageContext.web.absoluteUrl.replace(this.context.pageContext.web.serverRelativeUrl, ""),
        siteUrl: { SiteUrl }
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
