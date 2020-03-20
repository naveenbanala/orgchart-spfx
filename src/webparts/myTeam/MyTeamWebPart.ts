import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  IPropertyPaneCheckboxProps
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MyTeamWebPartStrings';
import MyTeam from './components/MyTeam';
import { IMyTeamProps } from './components/IMyTeamProps';

export interface IMyTeamWebPartProps {
  description: string;
  context: any;
  checkboxPeers: boolean;
  checkboxManagers: boolean;
  checkboxDirectReports: Boolean;
}

export default class MyTeamWebPart extends BaseClientSideWebPart<IMyTeamWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMyTeamProps> = React.createElement(
      MyTeam,
      {
        description: this.properties.description,
        context: this.context,
        checkboxPeers: this.properties.checkboxPeers,
        checkboxManagers: this.properties.checkboxManagers,
        checkboxDirectReports: this.properties.checkboxDirectReports
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
            },
            {
              groupName: "Config",
              groupFields: [
                PropertyPaneCheckbox('checkboxPeers', { text: "Peers" }),
                PropertyPaneCheckbox('checkboxManagers', { text: "Managers" }),
                PropertyPaneCheckbox('checkboxDirectReports', { text: "DirectReports" }),
              ]
            }
          ]
        }
      ]
    };
  }
}
