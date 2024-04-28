import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SideMenuWebPartStrings';
import SideMenu from './components/SideMenu';
import { ISideMenuProps } from './components/ISideMenuProps';

export interface ISideMenuWebPartProps {
  description: string;
  context: string;
  // TeamList: string;
  adminGroup: string;
  hrGroup: string;
}

export default class SideMenuWebPart extends BaseClientSideWebPart<ISideMenuWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISideMenuProps> = React.createElement(
      SideMenu,
      {
        context: this.context,
        description: this.properties.description,
        // TeamList: this.properties.TeamList,
        adminGroup: this.properties.adminGroup,
        hrGroup: this.properties.hrGroup
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
                }),
                // PropertyPaneTextField('TeamList', {
                //   label: "Team List",
                //   value: "Managers"
                // }),
                PropertyPaneTextField('adminGroup', {
                  label: 'Admin Group Name',
                  value: 'QT Portal Admin'
                }),  
                PropertyPaneTextField('hrGroup', {
                  label: 'HR Group Name',
                  value: 'HR'//'Human Resource (HR)'
                })  
              ]
            }
          ]
        }
      ]
    };
  }
}
