import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ITeamTimesheetProps {
  description: string;
  context: WebPartContext;
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
