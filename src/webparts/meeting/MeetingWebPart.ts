import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'MeetingWebPartStrings';
import Meeting from './components/Meeting';
import { IMeetingProps } from './components/IMeeting';
import { sp } from '@pnp/sp';
import { Logger, LogLevel, ConsoleListener } from "@pnp/logging";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { business } from '../../business';


export interface IMeetingWebPartProps {
  description: string;
}

export default class MeetingWebPart extends BaseClientSideWebPart<IMeetingWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMeetingProps> = React.createElement(
      Meeting,
      {
        description: this.properties.description
      }
    );

    sp.setup({
      spfxContext: this.context
    });

    // we are setting up the sp-pnp-js logging for debugging
    Logger.activeLogLevel = LogLevel.Info;
    Logger.subscribe(new ConsoleListener());

    SPComponentLoader.loadScript(this.context.pageContext.web.absoluteUrl + '/SiteAssets/Scripts/CommitteeConstants.js')
      .then(() => {
        business.setup({
          spfxContext: this.context
        });
      });

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
