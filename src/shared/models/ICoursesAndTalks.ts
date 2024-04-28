import { IUrl } from "../../shared/models/IUrl";

export interface ICoursesAndTalks {
    Title: string;
    CourseDescription:string;
    RegistrationCloseDate:string;
    StartDate:string;
    EndDate:string;
    Location:string;
    CourseImage?: IUrl;
    ContactPeopleId?: any[];
    CentreId: number; 
}