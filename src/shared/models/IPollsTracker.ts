import { IUser } from "./IUser";

export interface IPollsTracker {
    Title?: string;
    Effectivedate?: string;
    Expirydate?: string;
    Email?: boolean | number;
    TargetAudienceId?: number[];
    TargetAudience?: IUser[];
}