import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ChatbotWebPartStrings';
import Chatbot from './components/Chatbot';
import { IChatbotProps } from './components/IChatbotProps';

export interface IChatbotWebPartProps {
  botSchemaName: string;
  botSubtitle: string
  botname: string;
  botimage: string;
  botlogo: string;
}

export default class ChatbotWebPart extends BaseClientSideWebPart<IChatbotWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IChatbotProps> = React.createElement(
      Chatbot,
      {
        botSchemaName: this.properties.botSchemaName,
        botSubtitle: this.properties.botSubtitle,
        botname: this.properties.botname,
        botimage: this.properties.botimage,
        botlogo: this.properties.botlogo,
        userName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // protected onInit(): Promise<void> {
  //   return this._getEnvironmentMessage().then(message => {
  //     this._environmentMessage = message;
  //   });
  // }

  // private _getEnvironmentMessage(): Promise<string> {
  //   if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
  //     return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
  //       .then(context => {
  //         let environmentMessage: string = '';
  //         switch (context.app.host.name) {
  //           case 'Office': // running in Office
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
  //             break;
  //           case 'Outlook': // running in Outlook
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
  //             break;
  //           case 'Teams': // running in Teams
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
  //             break;
  //           default:
  //             throw new Error('Unknown host');
  //         }

  //         return environmentMessage;
  //       });
  //   }

  //   return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  // }

  // protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
  //   if (!currentTheme) {
  //     return;
  //   }

  //   this._isDarkTheme = !!currentTheme.isInverted;
  //   const {
  //     semanticColors
  //   } = currentTheme;

  //   if (semanticColors) {
  //     this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
  //     this.domElement.style.setProperty('--link', semanticColors.link || null);
  //     this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
  //   }

  // }

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
              groupName: "Bot Settings",
              groupFields: [
                PropertyPaneTextField('botSchemaName', {
                  label: "BOT Schema Name"
                }),
                PropertyPaneTextField('botSubtitle', {
                  label: "BOT Subtitle"
                }),
                PropertyPaneTextField('botname', {
                  label: "BOT Name"
                }),
                PropertyPaneTextField('botlogo', {
                  label: "BOT Logo"
                }),
                PropertyPaneTextField('botimage', {
                  label: "BOT Image"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
