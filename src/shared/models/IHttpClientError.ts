import { IHttpClientMessage } from "./IHttpClientMessage";

export interface IHttpClientError {
    code: string;
    message: IHttpClientMessage;
}