import { IUrl } from "./IUrl";
import { IUser } from "./IUser";

export interface ICommonResources {
    Id?: number;
    ID?: number;
    Title?: string;
    Description?: string;
    Image?: IUrl;
    DisplayOrder?: number;
    AssignedUsersId?: number[];
    AssignedUsers?: IUser[]
}