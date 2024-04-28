import { ILookup } from "./ILookup";
import { IUrl } from "./IUrl";
import { IUser } from "./IUser";

export interface ICommonResourceContent {
    Id?: number;
    ID?: number;
    Title?: string;
    Description?: string;
    Image?: IUrl;
    ResourceId?: Number;
    Resource?: ILookup;
    DisplayOrder?: number;
    AssignedUsersId?: number[];
    AssignedUsers?: IUser[];
}