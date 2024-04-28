import { ILookup } from "./ILookup";
import { IUrl } from "./IUrl";

export interface IFieldValuesAsText {
    Descriptions: string;
    ExpiryDate: string;

}

export interface IMediaLinks{
    Id?: number;
    ID?: number;
    //Title: string;
    Descriptions: string;
    FieldValuesAsText?: IFieldValuesAsText;
    DivisionsId?: ILookup;
    CentreId?: ILookup;
    AlbumStatus:string;
    ExpiryDate: string;
    CategoryId?: string;
    Image?: IUrl;
    URL?: IUrl;
    Keywords?: string;
}