import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MyTimesheetWebPartStrings';
import MyTimesheet from './components/MyTimesheet';
import { IMyTimesheetProps } from './components/IMyTimesheetProps';
import * as jQuery from 'jquery';
import * as bootstrap from 'bootstrap';

export interface IMyTimesheetWebPartProps {
  description: string;
  context: string;
  EmployeeList: string;
  ClientList: string;
  ProjectList: string;
  TaskType: string;
  // TeamList: string;
  Leaves:string;
  // TaskDescription: any;
}

export default class MyTimesheetWebPart extends BaseClientSideWebPart<IMyTimesheetWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMyTimesheetProps> = React.createElement(
      MyTimesheet,
      {
        context: this.context,
        description: this.properties.description,
        EmployeeList: this.properties.EmployeeList,
        ClientList: this.properties.ClientList,
        ProjectList: this.properties.ProjectList,
        TaskType: this.properties.TaskType,
        // TeamList: this.properties.TeamList,
        Leaves:this.properties.Leaves,
        // TaskDescription: this.properties.TaskDescription,
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
                // PropertyPaneTextField('TeamList', {
                //   label: "Team List",
                //   value: "Managers"
                // }),
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
