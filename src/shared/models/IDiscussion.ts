import { IFile } from "./IFile";
import { IUser } from "./IUser";

export interface IDiscussion {
    ID?: number;
    Id?: number;
    Title?: string;
    Body?: string;
    HasAttachment?: boolean;
    InvitePeople?: IUser[];
    InvitePeopleId?: number[];
    LikedPeopleId?: number[];
    LikedPeople?: IUser[];
    ParentDiscussionId?: number;
    ParentReplyId?: number | null;
    TotalComment?: number;
    Archive?: boolean;
    Author?: IUser;
    AuthorId?: number;
    Created?: string;
    Attachments?: IFile[];
}