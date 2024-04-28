import { ILookup } from "./ILookup";
import { IUrl } from "./IUrl";

export interface IFieldValuesAsText {
    Body: string;
    Expires: string;
    PublishDate: string;
}

export interface IAnnouncement {
    Id?: number;
    ID?: number;
    Title: string;
    Body: string;
    FieldValuesAsText?: IFieldValuesAsText;
    Category?: ILookup;
    Expires: string;
    PublishDate: string;
    CategoryId?: string;
    CentreId?: string;
    Image?: IUrl;
    Url?: IUrl;
    Keywords?: string;
}