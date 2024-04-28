import { IHttpClientError } from "./IHttpClientError";

export interface IHttpClientResponse {
    ok: boolean;
    status: number;
    statusText: string;
    type: string;
    url: string;
    data?: any;
    error: IHttpClientError;
}