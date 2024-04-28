import { IUser } from "./IUser";

export interface IPollSurveyTracker {
    ID?: number;
    Id?: number;
    Title?: string;
    Expirydate?: string;
    Participantsname?: IUser[];
    ParticipantsList?: IUser[];
    RespondedUsers?: IUser[];
    ParticipantsnameId?: number[];
    ParticipantsListId?: number[];
    RespondedUsersId?: number[];
    SendMail?: boolean;
    RootFolderUrl?: string;
    SurveyGUID?: string;
    ViewID?: string;
    Created?: string;
    Modified?: string;
    Effectivedate?: string; //polls additional column
}