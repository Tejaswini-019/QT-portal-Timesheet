import { IFile } from "./IFile";
import { IUrl } from "./IUrl";

export interface IRotatingBanner {
    Id?: number;
    ID?: number;
    Title?: string;
    Description?: string;
    Keywords?: string;
    PublishDate?: string;
    ExpiryDate?: string;
    AuthorId?: string;
    DisplayOrder?: number;
    Url?: IUrl;
    ImageCreateDate?: string | null;
    FileRef?: string;
    File?: IFile;
    FileLeafRef?: string; 
    Copyright?: string;
    Comments?: string;
    PreviewImageUrl?: IUrl;
}