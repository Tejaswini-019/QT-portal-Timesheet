import { IUrl } from "./IUrl";
import { ILookup } from "./ILookup";
import { IUser } from "./IUser";

export interface IStaffProfile {
    ID?: number;
    Id?: number;
    Title: string;
    StaffImage?: IUrl;
    Department?: ILookup;
    DepartmentId?: number;
    Designation?: string;
    Centre?:any;
    CentreId?:number;
    //OfficeLocation?: string;
    EmailAddress?: string;
    //OfficeContact?: string;
    ReportingOfficer?: string;
    Region?: string;
    Mobile?: string;
    Location?: string;
    EmpCode?: string;
    DateOfBirth?: string;
    dateJoined?: string;
    Country?: string;
    ContactNumber?: string;
    Company?: string;
    City?: string;
    BusinessGroup?: string;
    AreaOfExpertise?: string;
    AboutMe?: string;
}