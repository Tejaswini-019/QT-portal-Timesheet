import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TeamTimesheetWebPartStrings';
import TeamTimesheet from './components/TeamTimesheet';
import { ITeamTimesheetProps } from './components/ITeamTimesheetProps';

export interface ITeamTimesheetWebPartProps {
  description: string;
  context: string;
  // TeamList: string;
  EmployeeList: string;
  ClientList: string;
  ProjectList: string;
  TaskType: string;
  // TaskDescription: any;
  adminGroup: string;
  hrGroup: string;
  Leaves:string;
}

export default class TeamTimesheetWebPart extends BaseClientSideWebPart<ITeamTimesheetWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITeamTimesheetProps> = React.createElement(
      TeamTimesheet,
      {
        context: this.context,
        description: this.properties.description,
        // TeamList: this.properties.TeamList,
        EmployeeList: this.properties.EmployeeList,
        ClientList: this.properties.ClientList,
        ProjectList: this.properties.ProjectList,
        TaskType: this.properties.TaskType,
        Leaves:this.properties.Leaves,
        // TaskDescription: this.properties.TaskDescription,
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
                PropertyPaneTextField('EmployeeList', {
                  label: 'Timesheet List',
                  value: 'Timesheet'
                }),
                // PropertyPaneTextField('TaskDescription', {
                //   label: 'TaskDescription Library Internal Name',
                //   value: 'TaskDescription'
                // }), 
                PropertyPaneTextField('ClientList', {
                  label: 'Client List',
                  value: 'Clients'
                }),
                PropertyPaneTextField('ProjectList', {
                  label: 'Project List',
                  value: 'Projects'
                }),
                PropertyPaneTextField('TaskType', {
                  label: 'Task Type',
                  value: 'TaskType'
                }),
                PropertyPaneTextField('adminGroup', {
                  label: 'Admin Group Name',
                  value: 'QT Portal Admin'
                }),  
                PropertyPaneTextField('hrGroup', {
                  label: 'HR Group Name',
                  value: 'HR'//'Human Resource (HR)'
                }),
                PropertyPaneTextField('Leaves', {
                  label: "Leaves",
                  value: "Leaves"
                }),                          
              ]
            }
          ]
        }
      ]
    };
  }
}
