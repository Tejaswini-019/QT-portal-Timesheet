import { IUser } from "./IUser";

export interface IOutOfOffice {
    User?: IUser;
    UserId?: number,
    IsEnabled?: boolean;
    Delegatee?: IUser;
    DelegateeId?: number
    StartDate?: String;
    EndDate?: String;
    
}