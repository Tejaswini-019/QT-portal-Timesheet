import { IUser } from "./IUser";

export interface IResourceLibrary{
    Name: string;
    DisplayName: string;
    InternalName: string;
    ItemCount: string;
    ServerRelativeUrl:  string;
    AbsoluteUrl: string;
    SiteAbsoluteUrl: string;
    CurrentUser: IUser;
}