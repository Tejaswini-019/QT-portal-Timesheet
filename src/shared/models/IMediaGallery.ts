import { IUrl } from "./IUrl";

export default interface IMediaGallery {
    Title : string;
    Descriptions : string;
    Image? : IUrl;
    IsAlbum? : boolean;
    ExpiryDate : string;
    AlbumStatus : string;
    DivisionsId? : string;
    CentreId: number;
}