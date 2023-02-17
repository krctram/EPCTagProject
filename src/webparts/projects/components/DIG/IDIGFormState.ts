import { IBaseInterface } from "../../../../interfaces/IBaseInterface";
import { TAG_ProjectDetails } from "../../../../domain/models/TAG_ProjectDetails";
import { Enums } from "../../../../globals/Enums";


export interface IDIGFormState extends IBaseInterface{ 
    ProjectDetails: TAG_ProjectDetails;
  IsCreateMode: boolean;
  CurrentUserRoles: Enums.UserRoles[];
  DisableSaveButton: boolean;
  DisableSubmitButton: boolean;
  DisableRevertButton: boolean;

  OnlyVisibleForReviewer: boolean;
  OnlyEnableForReviewer: boolean;
}