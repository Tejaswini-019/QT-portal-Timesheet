import { IValidationField } from "./IValidationField";

export interface IValidationResult {
    IsValid: boolean;
    ValidFields: IValidationField[];
    InValidFields: IValidationField[];
}