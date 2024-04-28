import { IUrl } from "./IUrl";

export interface IQuickLink {
    Id?: number;
    ID?: number;
    URL?: IUrl;
    Icon?: IUrl;
    IsActive?: boolean;
    IsOpenNewTab?: boolean;
    DisplayOrder?: number;
    IsEditable?: boolean;
}