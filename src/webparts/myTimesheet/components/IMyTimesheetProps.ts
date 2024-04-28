import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IMyTimesheetProps {
  context:WebPartContext;
  description: string;
  EmployeeList: string;
  ClientList: string;
  ProjectList: string;
  TaskType: string;
  // TeamList: string;
  Leaves:string;
  // TaskDescription: any;
}
