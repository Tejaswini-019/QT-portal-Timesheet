import { ILookup } from "./ILookup";
import { IUser } from "./IUser";

export default interface IAwards {
    ID?: number;
    Id?: number;
    AwardDate?: string;
    Awardee?: IUser;
    AwardeeId?: number;
    AwardType?: ILookup;
    AwardTypeId?: string;
    Keywords?: string;
}