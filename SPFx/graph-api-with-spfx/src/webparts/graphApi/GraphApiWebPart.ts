import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'GraphApiWebPartStrings';
import GraphApi from './components/GraphApi';
import { IGraphApiUserProps } from './components/IGraphApiUserProps';

export interface IGraphApiWebPartProps {
  description: string;
}

export default class GraphApiWebPart extends BaseClientSideWebPart<IGraphApiWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private async getMsGraphClient() {
    return await this.context.msGraphClientFactory.getClient("3");
  }

  private async getUser() {
    const msGraphClient = await this.getMsGraphClient();
    const apiResponse = await msGraphClient.api('/me').get();
    return await apiResponse;
  }

  private logInfo(message: any): void {
    console.info(message);
  }
  public render(): void {

    this.getUser().then(async (result: any) => {
      this.logInfo(result);
      const element: React.ReactElement<IGraphApiUserProps> = React.createElement(
        GraphApi,
        {
          displayName: result.displayName,
          email: result.mail,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
        }
      );
      ReactDom.render(element, this.domElement);
    }).catch((err:any)=>{
      this.logInfo(err)
    });

  }

  protected onInit(): Promise<void> {

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });

  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
