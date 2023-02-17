import { TAG_ProjectDetails } from "../../../domain/models/TAG_ProjectDetails";
import { Enums } from "../../../globals/Enums";
import { IBaseInterface } from "../../../interfaces/IBaseInterface";

export interface IProjectsState extends IBaseInterface {
  ProjectDetails: TAG_ProjectDetails;
  IsCreateMode: boolean;
  CurrentUserRoles: Enums.UserRoles[];
  DisableSaveButton: boolean;
  DisableSubmitButton: boolean;
  DisableRevertButton: boolean;

  OnlyVisibleForReviewer: boolean;
  OnlyEnableForReviewer: boolean;
  
}
