import { IBaseInterface } from "../../../../interfaces/IBaseInterface";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Enums } from "../../../../globals/Enums";
export interface IDIGFormProps extends IBaseInterface {
    AppContext: WebPartContext;
    ItemID: string;
    IsLoading:boolean;
    hasEditItemPermission:boolean;
    ProjectDetails:any;
    CurrentUserRoles: Enums.UserRoles[];
    DisableSaveButton:boolean;
    DisableSubmitButton:boolean;
    DisableRevertButton:boolean;
    OnlyEnableForReviewer:boolean;
    OnlyVisibleForReviewer:boolean;
}
