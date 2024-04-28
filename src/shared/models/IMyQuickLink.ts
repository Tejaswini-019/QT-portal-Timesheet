import { IUrl } from "./IUrl";

export interface IMyQuickLink {
    Id?: number;
    ID?: number;
    URL?: IUrl;
    Icon?: IUrl;
    IsActive?: boolean;
    IsOpenNewTab?: boolean;
    DisplayOrder?: number;
    IsEditable?: boolean;
    TargetQuickLinkId?: number;
    UserId?: number;
}