import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClient } from "@microsoft/sp-http";
import * as strings from 'TaqeefDefinitionsWebPartStrings';
import TaqeefDefinitions from './components/TaqeefDefinitions';
import { ITaqeefDefinitionsProps } from './components/ITaqeefDefinitionsProps';

export interface ITaqeefDefinitionsWebPartProps {
  description: string;
}

export default class TaqeefDefinitionsWebPart extends BaseClientSideWebPart<ITaqeefDefinitionsWebPartProps> {
  private graphClient: MSGraphClient;
  public onInit(): Promise<void>{
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
        this.graphClient=client;
        resolve();
      },err => reject(err));
    });
  }

  public render(): void {
    const element: React.ReactElement<ITaqeefDefinitionsProps> = React.createElement(
      TaqeefDefinitions,
      {
        description: this.properties.description,
        absoluteURL : this.context.pageContext.web.absoluteUrl,
        spHttpClient : this.context.spHttpClient,
        graphClient: this.graphClient
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
