import { ILookup } from "./ILookup";
import { IUser } from "./IUser";

export interface IBPARequest {
    CompanyCode?: any;
    Id?: number;
    Title?: string;
    ApplicationType?: ILookup;
    Subtype?: String;
    Requestor?: IUser;
    SubmittedDateTime?: string;
    Subject?: string;
    ApplicationTypeId?: number;
    FromDate?: string;
    ToDate?: string;
    NoOfApprovers?: number | string;
    Approver1Id?: number;
    Approver2Id?: number;
    Approver3Id?: number;
    Approver4Id?: number;
    Approver5Id?: number;
    Approver1?: IUser;
    Approver2?: IUser;
    Approver3?: IUser;
    Approver4?: IUser;
    Approver5?: IUser;
    Status?: string;
    ReviewersId?: any[];
    Reviewers?: any;
    Description?: string;
    Reason?: string;
    RequestorId?: number;
    Remarks?: string;
    CurrentOwner?: IUser;
    CurrentOwnerId?: number;
    CompanyCodeId?: number;
    Delegatee1Id?: number;
    IsDelegated?: boolean;
}