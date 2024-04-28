import { ILookup } from "./ILookup";
import { IUrl } from "./IUrl";
import { IUser } from "./IUser";

interface IFieldValuesAsText {
    EventDate: string;
    EndDate: string;
    RegistrationCloseDate?: string;
}

export interface ICalendar {
    Id?: number;
    ID?: number;
    Title?: string;
    EventDescription?: any;
    Location?: string;
    EventDate?: string;
    EndDate?: string;
    MonthName?: string;
    MonthShortName?: string;
    EventDay?: string;
    Color?: string;
    OwnerInitial?: string;
    OwnerPhoto?: string;
    OwnerEmail?: string;
    OwnerName?: string;
    fAllDayEvent?: boolean;
    ParticipantsPickerId?: number[];
    ParticipantsPicker?: any[];
    Category?: string;
    Duration?: number;
    RecurrenceData?: string;
    fRecurrence?: string | boolean;
    EventType?: number;
    UID?: string;
    RecurrenceID?: Date;
    MasterSeriesItemID?: string;
    RegistrationCloseDate?: string;
    SendEmail?: boolean;
    Scope?: string;
    ContactPersonId?: number[];
    ContactPerson?: IUser[];
    DivisionsId?: number;
    Division?: ILookup;
    EventCategoryId?: number;
    EventCategory?: ILookup;
    Url?: IUrl;
    Image?: IUrl;
    FieldValuesAsText?: IFieldValuesAsText;
    CentreId: number;
    Centre?: ILookup;
}