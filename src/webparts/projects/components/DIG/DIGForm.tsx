import * as React from "react";
import styles from "../Projects.module.scss";
import ListItemService from "./../../../../services/ListItemService";
import UserService from "../../../../services/UserService";
import WebService from "../../../../services/WebService";
import { TAG_ProjectDetails } from "../../../../domain/models/TAG_ProjectDetails";
import { Enums } from "../../../../globals/Enums";
import { Config } from "../../../../globals/Config";
import { User } from "../../../../domain/models/types/User";
import { Label } from "@fluentui/react/node_modules/office-ui-fabric-react";

import {
  Dropdown,
  IDropdownOption,
  IStackTokens,
  PrimaryButton,
  Stack,
  TextField,
} from "@fluentui/react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Parser } from "html-to-react";
import MapResult from "../../../../domain/mappers/MapResult";
import { IDIGFormProps } from "./IDIGFormProps";
import { IDIGFormState } from "./IDIGFormState";

export default class DIGForm extends React.Component<
  IDIGFormProps,
  IDIGFormState
> {
  private listItemService: ListItemService;
  private userService: UserService;
  private webService: WebService;
  private hasEditItemPermission: boolean = true;

  constructor(props: any) {
    super(props);
    this.state = {
      IsCreateMode:
        this.props.ItemID == undefined ||
        this.props.ItemID == null ||
        this.props.ItemID == "0"
          ? true
          : false,
      IsLoading: this.props.IsLoading,
      AppContext: this.props.AppContext,
      DisableSaveButton: this.props.DisableSaveButton,
      DisableRevertButton: this.props.DisableRevertButton,
      DisableSubmitButton: this.props.DisableSubmitButton,
      CurrentUserRoles: this.props.CurrentUserRoles,
      ProjectDetails: this.props.ProjectDetails,
      OnlyVisibleForReviewer: this.props.OnlyVisibleForReviewer,
      OnlyEnableForReviewer: this.props.OnlyEnableForReviewer,
    };
    this.onSaveAsDraft = this.onSaveAsDraft.bind(this);
    this.onSubmit = this.onSubmit.bind(this);
    this.onCancel = this.onCancel.bind(this);
    this.onReplaceMe = this.onReplaceMe.bind(this);
    this.onRevert = this.onRevert.bind(this);
    this.onFormFieldValueChange = this.onFormFieldValueChange.bind(this);
    this.onChangeDropdownValues = this.onChangeDropdownValues.bind(this);
    this.onChangeTextField = this.onChangeTextField.bind(this);
    this.onChangePersonField = this.onChangePersonField.bind(this);
  }
  // Things to be performed when the component is being mounted
  public async componentDidMount() {
    this.userService = new UserService(this.props.AppContext);
    this.webService = new WebService(this.props.AppContext);
    const userRoles: Enums.UserRoles[] = await this.GetCurrentUserRoles();

    // CASE: CREATE ITEM
    if (this.state.IsCreateMode) {
    }
    // CASE: EDIT ITEM
    else {
      this.listItemService = new ListItemService(
        this.props.AppContext,
        Config.ListNames.Projects
      );
      // const projectDetails: TAG_ProjectDetails = await this.listItemService.getItemUsingCAML(parseInt(this.props.ItemID), [], undefined, Enums.ItemResultType.TAG_ProjectDetails);
      //   this.hasEditItemPermission = await this.listItemService.CheckCurrentUserCanEditItem(parseInt(this.props.ItemID));
      //   const allowSave: boolean = this.validateSave(this.state.ProjectDetails);
      //   const allowSubmit: boolean = this.validateSubmit(this.state.ProjectDetails);
      //   const allowRevert: boolean = this.validateRevert(this.state.ProjectDetails);
      //   const enableForReviewer: boolean = projectDetails.StatusOfReview == Config.Strings.Status_AwaitingReviewer;
      //   const visibleForReviewer: boolean = projectDetails.StatusOfReview == Config.Strings.Status_AwaitingReviewer;

      this.setState({
        IsLoading: false,
        //CurrentUserRoles: userRoles,
        // ProjectDetails: projectDetails,
        // DisableSaveButton: !allowSave,
        // DisableSubmitButton: !allowSubmit,
        // DisableRevertButton: !allowRevert,
        // OnlyEnableForReviewer: enableForReviewer,
        // OnlyVisibleForReviewer: visibleForReviewer
      });
    }
  }

  //#region  On chnage Evenet
  // Updating Field values in Reviews List
  private async onFormSave(saveType: Enums.SaveType) {
    const projectDetails = this.state.ProjectDetails;
    let data = {};
    const columns = Config.ProjectsListColumns;

    data[columns["Fiscal Year"]] = projectDetails.FiscalYear;
    data[columns["Service Line"]] = projectDetails.ServiceLine;
    data[columns["Reviewer Name ID"]] = projectDetails.Reviewer.Id;
    data[columns["Lead MD Name ID"]] = projectDetails.LeadMD.Id;

    data[columns["Hours Worked"]] = projectDetails.HoursWorked;
    data[columns.Complexity] = projectDetails.Complexity;

    // Service Excellence
    data[columns.Q1] = projectDetails.Q1;
    data[columns.Q2] = projectDetails.Q2;

    // Fundamental Expertise
    data[columns.Q3] = projectDetails.Q3;
    data[columns.Q4] = projectDetails.Q4;
    data[columns.Q5] = projectDetails.Q5;
    data[columns.Q6] = projectDetails.Q6;

    // Practice Operations & Leadership
    data[columns.Q7] = projectDetails.Q7;
    data[columns.Q8] = projectDetails.Q8;
    data[columns.Q9] = projectDetails.Q9;

    // Personal Effectiveness
    data[columns.Q10] = projectDetails.Q10;
    data[columns.Q11] = projectDetails.Q11;
    data[columns.Q12] = projectDetails.Q12;
    data[columns.Q13] = projectDetails.Q13;

    // Reviewee Comments Questions
    data[columns["Strong Performance"]] = projectDetails.StrongPerformance;
    data[columns["Development Areas"]] = projectDetails.DevelopmentAreas;
    data[columns["Needed Skills"]] = projectDetails.NeededSkills;

    // Other Comments
    data[columns["Lead MD Comments"]] = projectDetails.LeadMDComments;
    data[columns["Reviewer Reversion Comments"]] =
      projectDetails.ReviewerReversionComments;
    data[columns["Lead MD Reversion Comments"]] =
      projectDetails.LeadMDReversionComments;
    data[columns["Acknowledgement Comments"]] =
      projectDetails.AcknowledgementComments;

    // Review Status
    data[columns["Status of Review"]] = projectDetails.StatusOfReview;
    data[columns.Submitted] = projectDetails.Submitted;

    // Only for Replace Me
    if (
      projectDetails.SubstituteUser.Id &&
      saveType == Enums.SaveType.ReplaceMe
    ) {
      data[columns["SubstituteUser Id"]] = projectDetails.SubstituteUser.Id;
    }

    this.listItemService = new ListItemService(
      this.props.AppContext,
      Config.ListNames.Projects
    );
    if (this.state.IsCreateMode) {
      await this.listItemService.createItem(data);
    } else {
      await this.listItemService.updateItem(parseInt(this.props.ItemID), data);
    }

    // Redirecting user to main listing page once saving is done
    if (saveType != Enums.SaveType.SaveAsDraft) {
      this.gotoListPage();
    } else {
      alert("Changes are saved successfully.");
    }
  }

  // On change of dropdown values

  private onChangeDropdownValues(fieldName: string, newValue: string): void {
    let curretState = this.state.ProjectDetails;
    curretState[fieldName] = newValue;
    this.onFormFieldValueChange(curretState);
  }
  // On change of textbox values
  private onChangeTextField(fieldName: string, newValue: string): void {
    let curretState = this.state.ProjectDetails;
    curretState[fieldName] = newValue;
    this.onFormFieldValueChange(curretState);
  }

  // On click event of 'Save as Draft' button
  private async onSaveAsDraft() {
    let requiredProjectDetails = this.state.ProjectDetails;
    requiredProjectDetails.Submitted = this.getSubmitStatusValue(
      this.state.ProjectDetails.StatusOfReview,
      Enums.SaveType.SaveAsDraft
    );
    this.setState(
      {
        ProjectDetails: requiredProjectDetails,
      },
      async () => {
        await this.onFormSave(Enums.SaveType.SaveAsDraft);
      }
    );
  }

  // On click event of 'Submit' buttons
  // - Start Review
  // - Submit to Reviewer for Approval
  // - Submit to Lead MD for Approval
  // - Submit to Reviewee for Acknowledgement
  // - Submit Final Review
  private async onSubmit() {
    let requiredProjectDetails = this.state.ProjectDetails;
    requiredProjectDetails.Submitted = this.getSubmitStatusValue(
      this.state.ProjectDetails.StatusOfReview,
      Enums.SaveType.Submit
    );
    this.setState(
      {
        ProjectDetails: requiredProjectDetails,
      },
      async () => {
        await this.onFormSave(Enums.SaveType.Submit);
      }
    );
  }

  // On click event of 'Revert' buttons
  // - Revert to Reviewee
  // - Revert to Reviewer
  private async onRevert() {
    const requiredProjectDetails = this.state.ProjectDetails;
    requiredProjectDetails.Submitted = this.getSubmitStatusValue(
      this.state.ProjectDetails.StatusOfReview,
      Enums.SaveType.Revert
    );
    this.setState(
      {
        ProjectDetails: requiredProjectDetails,
      },
      async () => {
        await this.onFormSave(Enums.SaveType.Revert);
      }
    );
  }

  // On click event of 'Replace Me' button
  private async onReplaceMe() {
    const requiredProjectDetails = this.state.ProjectDetails;
    requiredProjectDetails.Submitted = this.getSubmitStatusValue(
      this.state.ProjectDetails.StatusOfReview,
      Enums.SaveType.ReplaceMe
    );
    this.setState(
      {
        ProjectDetails: requiredProjectDetails,
      },
      async () => {
        await this.onFormSave(Enums.SaveType.ReplaceMe);
      }
    );
  }

  // On click event of 'Cancel' button
  private async onCancel() {
    this.gotoListPage();
  }
  // Updating Project Details updated in child components
  private onFormFieldValueChange(updateDetails: TAG_ProjectDetails) {
    let allowSave: boolean = this.validateSave(updateDetails);
    let allowSubmit: boolean = this.validateSubmit(updateDetails);
    let allowRevert: boolean = this.validateRevert(updateDetails);
    this.setState({
      ProjectDetails: updateDetails,
      DisableSaveButton: !allowSave,
      DisableSubmitButton: !allowSubmit,
      DisableRevertButton: !allowRevert,
    });
  }
  // On change of person field
  private async onChangePersonField(
    fieldName: string,
    items: any[]
  ): Promise<void> {
    let curretState = this.state.ProjectDetails;
    if (items != null && items.length > 0) {
      curretState[fieldName] = await MapResult.map(
        items[0],
        Enums.MapperType.PnPControlResult,
        Enums.ItemResultType.User
      );
    } else {
      curretState[fieldName] = new User();
    }
    this.onFormFieldValueChange(curretState);
  }
  //#endregion
  //#region "Utility Methods"

  // Deciding Submitted field value
  private getSubmitStatusValue(
    currentStatusOfReview: string,
    buttonType: Enums.SaveType
  ): number {
    let result: number;

    //Case: Replace Me
    if (buttonType == Enums.SaveType.ReplaceMe) {
      result = 8;
    }

    // Case: Save as Draft
    if (buttonType == Enums.SaveType.SaveAsDraft) {
      result = 1;
    }

    //Case: Submit
    if (buttonType == Enums.SaveType.Submit) {
      switch (currentStatusOfReview) {
        case Config.Strings.Status_NotStarted:
          result = 99; // Review Started
          break;
        case Config.Strings.Status_AwaitingReviewee:
          result = 2; // Reviewee Approved/Responded
          break;
        case Config.Strings.Status_AwaitingReviewer:
          result = 4; // Reviewer Approved
          break;
        case Config.Strings.Status_AwaitingLeadMD:
          result = 6; // Lead MD Approved
          break;
        case Config.Strings.Status_AwaitingAcknowledgement:
          result = 7; // Acknowledged by Reviewee
          break;
      }
    }

    //Case: Revert
    if (buttonType == Enums.SaveType.Revert) {
      switch (currentStatusOfReview) {
        case Config.Strings.Status_AwaitingReviewer:
          result = 3;
          break;
        case Config.Strings.Status_AwaitingLeadMD:
          result = 5;
          break;
      }
    }
    return result;
  }

  // Redirect user to 'Employee Summary' Listing page
  private gotoListPage() {
    let returnURL =
      this.props.AppContext.pageContext.web.absoluteUrl +
      Config.Links.ProjectsListAllItems;
    window.location.href = returnURL;
    return false;
  }

  // Generating empty object for Review Details
  private async generateEmptyDetails(): Promise<TAG_ProjectDetails> {
    let details: TAG_ProjectDetails = {
      ID: null,
      AcknowledgementComments: "",
      ClientName: "",
      Complexity: "",
      DateOriginated: null,
      DateReviewCompleted: null,
      DevelopmentAreas: "",
      FiscalYear: "",
      HomeOffice: "",
      HoursWorked: null,
      JobTitle: "",
      LastHoursBilled: null,
      LeadMD: new User(),
      LeadMDComments: "",
      LeadMDReversionComments: "",
      Mentor: new User(),
      NeededSkills: "",
      PermReset: null,
      ProjectCode: "",
      ProjectEndDate: null,
      ProjectManager: new User(),
      ProjectName: "",
      ProjectStartDate: null,
      ProjectStatus: "",
      Reviewee: new User(),
      Reviewer: new User(),
      ReviewerReversionComments: "",
      ServiceLine: "",
      SignoffHistory: "",
      StatusOfReview: "",
      StrongPerformance: "",
      Submitted: 0,
      SubstituteUser: new User(),
      Q1: "",
      Q10: "",
      Q10Text: "",
      Q11: "",
      Q11Text: "",
      Q12: "",
      Q12Text: "",
      Q13: "",
      Q13Text: "",
      Q14: "",
      Q14Text: "",
      Q15: "",
      Q15Text: "",
      Q16: "",
      Q16Text: "",
      Q17: "",
      Q17Text: "",
      Q18: "",
      Q18Text: "",
      Q1Text: "",
      Q2: "",
      Q2Text: "",
      Q3: "",
      Q3Text: "",
      Q4: "",
      Q4Text: "",
      Q5: "",
      Q5Text: "",
      Q6: "",
      Q6Text: "",
      Q7: "",
      Q7Text: "",
      Q8: "",
      Q8Text: "",
      Q9: "",
      Q9Text: "",
      DateOriginatedFormatted: "",
      DateReviewCompletedFormatted: "",
      LastHoursBilledFormatted: "",
      ProjectEndDateFormatted: "",
      ProjectStartDateFormatted: "",
      ModifiedBy: new User(),
      ModifiedOn: null,
      ModifiedOnFormatted: "",
    };

    return details;
  }

  // Validations reqired for 'Save' button
  private validateSave(updatedProjectDetails: TAG_ProjectDetails): boolean {
    let valid: boolean = true;
    // If use has no edit rights
    if (!this.hasEditItemPermission) {
      valid = false;
    }

    // If Previous updation is in progress
    if (
      updatedProjectDetails.Submitted != 1 &&
      updatedProjectDetails.Submitted != 0 &&
      updatedProjectDetails.Submitted != null
    ) {
      valid = false;
    }
    return valid;
  }

  // Validations reqired for 'Submit' button
  // - Start Review
  // - Submit to Reviewer for Approval
  // - Submit to Lead MD for Approval
  // - Submit to Reviewee for Acknowledgement
  // - Submit Final Review
  private validateSubmit(updatedProjectDetails: TAG_ProjectDetails): boolean {
    let valid: boolean = true;
    const details = updatedProjectDetails;

    // If use has no edit rights
    if (!this.hasEditItemPermission) {
      valid = false;
    }

    // If Previous updation is in progress
    if (
      updatedProjectDetails.Submitted != 1 &&
      updatedProjectDetails.Submitted != 0 &&
      updatedProjectDetails.Submitted != null
    ) {
      valid = false;
    }

    // Validations required when Status is "Not Started"
    if (details.StatusOfReview == Config.Strings.Status_NotStarted) {
      if (
        details.FiscalYear == "" ||
        details.Reviewer.Id == null ||
        details.LeadMD.Id == null ||
        details.ServiceLine == ""
      ) {
        valid = false;
      }
    }

    // Validations required when status is "Awaiting Reviewee"
    if (details.StatusOfReview == Config.Strings.Status_AwaitingReviewee) {
      if (details.Reviewer.Id == null || details.LeadMD.Id == null) {
        valid = false;
      }
    }

    // Validations required when status is "Awaiting Reviewer"
    if (details.StatusOfReview == Config.Strings.Status_AwaitingReviewer) {
      if (
        details.Reviewer.Id == null ||
        details.LeadMD.Id == null ||
        details.Q1 == "" ||
        details.Q2 == "" ||
        details.Q3 == "" ||
        details.Q4 == "" ||
        details.Q5 == "" ||
        details.Q6 == "" ||
        details.Q7 == "" ||
        details.Q8 == "" ||
        details.Q9 == "" ||
        details.Q10 == "" ||
        details.Q11 == "" ||
        details.Q12 == "" ||
        details.Complexity == ""
      ) {
        valid = false;
      }
    }

    // Validations required when status is "Awaiting Lead MD"
    if (details.StatusOfReview == Config.Strings.Status_AwaitingLeadMD) {
      if (details.Reviewer.Id == null || details.LeadMD.Id == null) {
        valid = false;
      }
    }

    return valid;
  }

  // Validations reqired for 'Revert' button
  // - Revert to Reviewer
  // - Revert to Reviewee
  private validateRevert(updatedProjectDetails: TAG_ProjectDetails): boolean {
    let valid: boolean = true;
    const details = updatedProjectDetails;

    // If use has no edit rights
    if (!this.hasEditItemPermission) {
      valid = false;
    }

    // If Previous updation is in progress
    if (
      updatedProjectDetails.Submitted != 1 &&
      updatedProjectDetails.Submitted != 0 &&
      updatedProjectDetails.Submitted != null
    ) {
      valid = false;
    }

    // Validations required when status is "Awaiting Reviewer"
    if (details.StatusOfReview == Config.Strings.Status_AwaitingReviewer) {
      // No Validations as of now
    }

    // Validations required when status is "Awaiting Lead MD"
    if (details.StatusOfReview == Config.Strings.Status_AwaitingLeadMD) {
      if (details.Reviewer.Id == null || details.LeadMD.Id == null) {
        valid = false;
      }
    }

    return valid;
  }

  // Deciding the roles associated with current user
  private async GetCurrentUserRoles(): Promise<Enums.UserRoles[]> {
    let result: Enums.UserRoles[] = [];
    // Checking user is site collection admin  or member of 'DI Admin Group'
    const isSiteCollectionAdmin: boolean =
      await this.userService.CheckCurrentUserIsAdmin();
    const ownerGroupName: string =
      await this.webService.GetAssociatedOwnerGroupName();
    const isMemberOfOwnersGroup: boolean =
      await this.userService.CheckCurrentUserInSPGroup(ownerGroupName);
    if (isSiteCollectionAdmin || isMemberOfOwnersGroup) {
      result.push(Enums.UserRoles.SuperAdmin);
    }
    return result;
  }

  //#endregion

  public render(): React.ReactElement<IDIGFormProps> {
    const questionChoices: IDropdownOption[] = [
      { key: "N/A", text: "N/A" },
      { key: "1", text: "1" },
      { key: "2", text: "2" },
      { key: "3", text: "3" },
      { key: "4", text: "4" },
    ];

    const question7Choices: IDropdownOption[] = [
      { key: "N/A", text: "N/A" },
      { key: "Demonstrates", text: "Demonstrates" },
      { key: "Does not demonstrate", text: "Does not demonstrate" },
    ];
    const stackTokens: IStackTokens = { childrenGap: 20 };
    return this.props.IsLoading == true ? (
      <React.Fragment></React.Fragment>
    ) : (
      <div className={styles.sectionContainer}>
        <div>
          <div className={styles.sectionContainer}>
            <div className={styles.sectionHeader}>
              <div className={styles.colHeader100}>
                <span className={styles.subTitle}>Business Development</span>
              </div>
            </div>
            <div className={styles.sectionContent}>
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q1Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    disabled={!this.state.OnlyEnableForReviewer}
                    selectedKey={this.state.ProjectDetails.Q1}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q1", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q2Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q2}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q2", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q3Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q3}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q3", selectedOption.text);
                    }}
                  />
                </div>
              </div>

              <div className={styles.Spacer}>&nbsp;</div>
            </div>
          </div>

          <div className={styles.sectionContainer}>
            <div className={styles.sectionHeader}>
              <div className={styles.colHeader100}>
                <span className={styles.subTitle}>
                  Practice Operations and Leadership
                </span>
              </div>
            </div>
            <div className={styles.sectionContent}>
              {/* Q4 */}
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q4Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q4}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q4", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
              {/* Q5 */}
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q5Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q5}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q5", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
              {/* Q6 */}
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q6Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q6}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q6", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
              {/* Q7 */}
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q7Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q7}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q7", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
              {/* Q8 */}
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q8Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q8}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q8", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
              {/* Q9 */}
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q9Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q9}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q9", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
            </div>
          </div>

          <div className={styles.sectionContainer}>
            <div className={styles.sectionHeader}>
              <div className={styles.colHeader100}>
                <span className={styles.subTitle}>Hard Skill Assessment</span>
              </div>
            </div>
            <div className={styles.sectionContent}>
              {/* Q10 */}
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q10Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={question7Choices}
                    selectedKey={this.state.ProjectDetails.Q10}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q10", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
              {/* Q111 */}
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q11Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q11}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q11", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
              {/* Q12 */}
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q12Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q12}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q12", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
              {/* Q13 */}
              <div className={styles.row}>
                <div className={styles.col80left}>
                  <div>{Parser().parse(this.state.ProjectDetails.Q13Text)}</div>
                </div>
                <div className={styles.col20left}>
                  <Dropdown
                    placeholder="Select"
                    options={questionChoices}
                    selectedKey={this.state.ProjectDetails.Q13}
                    disabled={!this.state.OnlyEnableForReviewer}
                    onChange={(e, selectedOption) => {
                      this.onChangeDropdownValues("Q13", selectedOption.text);
                    }}
                  />
                </div>
              </div>
              <div className={styles.Spacer}>&nbsp;</div>
            </div>
          </div>

          {/* <div className={styles.sectionContainer}>
                            <div className={styles.sectionHeader}>
                                <div className={styles.colHeader100}>
                                    <span className={styles.subTitle}>PERSONAL EFFECTIVENESS</span>
                                </div>
                            </div>
                            <div className={styles.sectionContent}>
                                <div className={styles.row}>
                                    <div className={styles.col80left}>
                                        <div>
                                            {Parser().parse(this.state.ProjectDetails.Q10Text)}
                                        </div>
                                    </div>
                                    <div className={styles.col20left}>
                                        <Dropdown
                                            placeholder="Select"
                                            options={questionChoices}
                                            selectedKey={this.state.ProjectDetails.Q10}
                                            disabled={!this.state.OnlyEnableForReviewer}
                                            onChange={(e, selectedOption) => {
                                                this.onChangeDropdownValues("Q10", selectedOption.text);
                                            }} />
                                    </div>
                                </div>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <div className={styles.row}>
                                    <div className={styles.col80left}>
                                        <div>
                                            {Parser().parse(this.state.ProjectDetails.Q11Text)}
                                        </div>
                                    </div>
                                    <div className={styles.col20left}>
                                        <Dropdown
                                            placeholder="Select"
                                            options={questionChoices}
                                            selectedKey={this.state.ProjectDetails.Q11}
                                            disabled={!this.state.OnlyEnableForReviewer}
                                            onChange={(e, selectedOption) => {
                                                this.onChangeDropdownValues("Q11", selectedOption.text);
                                            }} />
                                    </div>
                                </div>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <div className={styles.row}>
                                    <div className={styles.col80left}>
                                        <div>
                                            {Parser().parse(this.state.ProjectDetails.Q12Text)}
                                        </div>
                                    </div>
                                    <div className={styles.col20left}>
                                        <Dropdown
                                            placeholder="Select"
                                            options={question7Choices}
                                            selectedKey={this.state.ProjectDetails.Q12}
                                            disabled={!this.state.OnlyEnableForReviewer}
                                            onChange={(e, selectedOption) => {
                                                this.onChangeDropdownValues("Q12", selectedOption.text);
                                            }} />
                                    </div>
                                </div>
                                <div className={styles.Spacer}>&nbsp;</div>
                            </div>
                        </div> */}

          <div className={styles.sectionContainer}>
            <div className={styles.sectionHeader}>
              <div className={styles.colHeader100}></div>
            </div>
            <div className={styles.sectionContent}>
              <div className={styles.row}>
                <div className={styles.col100}>
                  <Label>
                    Briefly comment on the Reviewee's top areas of strong
                    performance on this project. Your comments should support,
                    at minimum, any 4 ratings above.{" "}
                    <i>(Commentary required)</i>
                  </Label>
                  <TextField
                    resizable={false}
                    multiline={true}
                    value={this.state.ProjectDetails.StrongPerformance}
                    disabled={
                      this.state.ProjectDetails.StatusOfReview ==
                        Config.Strings.Status_AwaitingLeadMD ||
                      this.state.ProjectDetails.StatusOfReview ==
                        Config.Strings.Status_AwaitingAcknowledgement ||
                      this.state.ProjectDetails.StatusOfReview ==
                        Config.Strings.Status_Acknowledged
                    }
                    onChange={(e, newValue) => {
                      this.onChangeTextField("StrongPerformance", newValue);
                    }}
                  ></TextField>
                </div>
              </div>
              <div className={styles.row}>
                <div className={styles.col100}>
                  <Label>
                    Briefly comment on the Reviewee's top areas for development.
                    Your comments should support, at minimum, any 1 rating
                    above. <i>(Commentary required)</i>
                  </Label>
                  <TextField
                    resizable={false}
                    multiline={true}
                    value={this.state.ProjectDetails.DevelopmentAreas}
                    disabled={
                      this.state.ProjectDetails.StatusOfReview ==
                        Config.Strings.Status_AwaitingLeadMD ||
                      this.state.ProjectDetails.StatusOfReview ==
                        Config.Strings.Status_AwaitingAcknowledgement ||
                      this.state.ProjectDetails.StatusOfReview ==
                        Config.Strings.Status_Acknowledged
                    }
                    onChange={(e, newValue) => {
                      this.onChangeTextField("DevelopmentAreas", newValue);
                    }}
                  ></TextField>
                </div>
              </div>
              <div className={styles.row}>
                <div className={styles.col100}>
                  <Label>
                    Briefly comment on what skills are necessary for the
                    reviewee to continue to develop in order to progress in
                    their career. <i>(Commentary required)</i>
                  </Label>
                  <TextField
                    resizable={false}
                    multiline={true}
                    disabled={
                      this.state.ProjectDetails.StatusOfReview ==
                        Config.Strings.Status_AwaitingLeadMD ||
                      this.state.ProjectDetails.StatusOfReview ==
                        Config.Strings.Status_AwaitingAcknowledgement ||
                      this.state.ProjectDetails.StatusOfReview ==
                        Config.Strings.Status_Acknowledged
                    }
                    value={this.state.ProjectDetails.NeededSkills}
                    onChange={(e, newValue) => {
                      this.onChangeTextField("NeededSkills", newValue);
                    }}
                  ></TextField>
                </div>
              </div>
              <div className={styles.row}>
                <div className={styles.col100}>
                  <Label>
                    Additional comments from Lead MD <i>(optional)</i>
                  </Label>
                  <TextField
                    resizable={false}
                    multiline={true}
                    disabled={
                      this.state.ProjectDetails.StatusOfReview !=
                      Config.Strings.Status_AwaitingLeadMD
                    }
                    value={this.state.ProjectDetails.LeadMDComments}
                    onChange={(e, newValue) => {
                      this.onChangeTextField("LeadMDComments", newValue);
                    }}
                  ></TextField>
                </div>
              </div>
            </div>
          </div>

          {
            // SECTION: Reviewee Approval - Visible only while Awaiting Reviewee Approvall
            this.state.ProjectDetails.StatusOfReview !=
              Config.Strings.Status_AwaitingReviewer &&
              this.state.ProjectDetails.StatusOfReview !=
                Config.Strings.Status_AwaitingLeadMD &&
              this.state.ProjectDetails.StatusOfReview !=
                Config.Strings.Status_AwaitingAcknowledgement &&
              this.state.ProjectDetails.StatusOfReview !=
                Config.Strings.Status_Acknowledged && (
                <div className={styles.sectionContainer}>
                  <div className={styles.sectionContent}>
                    <div className={styles.row}>
                      <div className={styles.col100}>
                        <b>REVIEWEE:</b> When your comments are complete, click
                        the Submit button below. To identify a different
                        Reviewer or Lead MD to perform this review, change the
                        corresponding field(s) at the top of this form before
                        submitting. (Not ready yet? You can <b>Save Draft</b> to
                        preserve your inputs prior to submitting to the
                        Reviewer.)
                      </div>
                    </div>
                    <div className={styles.row}>
                      <div className={styles.col100}>
                        <Stack
                          horizontal
                          tokens={stackTokens}
                          className={styles.stackCenter}
                        >
                          <PrimaryButton
                            text="SAVE DRAFT"
                            onClick={this.onSaveAsDraft}
                            disabled={this.state.DisableSaveButton}
                          ></PrimaryButton>
                          <PrimaryButton
                            text="SUBMIT TO REVIEWER FOR APPROVAL"
                            onClick={this.onSubmit}
                            disabled={this.state.DisableSubmitButton}
                          ></PrimaryButton>
                        </Stack>
                      </div>
                    </div>
                  </div>
                </div>
              )
          }

          {
            // SECTION: Reviewer Approval - Visible only while Awaiting Reviewer
            this.state.OnlyVisibleForReviewer && (
              <div className={styles.sectionContainer}>
                <div className={styles.sectionContent}>
                  <div className={styles.row}>
                    <div className={styles.col100}>
                      <b>REVIEWER:</b> When you are ready to advance the review
                      to the Lead MD, click the Submit button below. Click Save
                      Draft to save and return later. To revert back to the
                      Reviewee, complete the gray section below.
                      <br />
                      <b>
                        To substitute a different Reviewer in this review,
                      </b>{" "}
                      enter the new name at the top of the form and click{" "}
                      <b>Replace Me.</b> Your current inputs will be saved, and
                      the review will be assigned to the new person.
                      <br />
                      <b>To identify a new Lead MD,</b> change the Lead MD name
                      at the top of this form and click either Save Draft or
                      Submit.
                    </div>
                  </div>
                  <div className={styles.row}>
                    <div className={styles.col100}>
                      <Stack
                        horizontal
                        tokens={stackTokens}
                        className={styles.stackCenter}
                      >
                        <PrimaryButton
                          text="SAVE DRAFT"
                          onClick={this.onSaveAsDraft}
                          disabled={this.state.DisableSaveButton}
                        ></PrimaryButton>
                        <PrimaryButton
                          text="SUBMIT TO LEAD MD FOR APPROVAL"
                          onClick={this.onSubmit}
                          disabled={this.state.DisableSubmitButton}
                        ></PrimaryButton>
                      </Stack>
                    </div>
                  </div>
                  <div className={styles.Spacer}>&nbsp;</div>
                  <div className={styles.row}>
                    <div className={styles.col75left}>
                      <label>Optional Reversion Comment (visible)</label>
                      <TextField
                        resizable={false}
                        multiline={false}
                        value={
                          this.state.ProjectDetails.ReviewerReversionComments
                        }
                        onChange={(e, newValue) => {
                          this.onChangeTextField(
                            "ReviewerReversionComments",
                            newValue
                          );
                        }}
                      ></TextField>
                    </div>
                    <div className={styles.col25Right}>
                      <div className={styles.Spacer}>&nbsp;</div>
                      <PrimaryButton
                        text="REVERT TO REVIEWEE"
                        onClick={this.onRevert}
                        disabled={this.state.DisableRevertButton}
                      ></PrimaryButton>
                    </div>
                  </div>
                </div>
              </div>
            )
          }

          {
            // SECTION: Lead MD Approval - Visible only while Awaiting Lead MD Approval
            this.state.ProjectDetails.StatusOfReview !=
              Config.Strings.Status_AwaitingReviewee &&
              this.state.ProjectDetails.StatusOfReview !=
                Config.Strings.Status_AwaitingReviewer &&
              this.state.ProjectDetails.StatusOfReview !=
                Config.Strings.Status_AwaitingAcknowledgement &&
              this.state.ProjectDetails.StatusOfReview !=
                Config.Strings.Status_Acknowledged && (
                <div className={styles.sectionContainer}>
                  <div className={styles.sectionContent}>
                    <div className={styles.row}>
                      <div className={styles.col100}>
                        <b>LEAD MD:</b> Review the form. Add any optional
                        comments in the text area above. When you are satisfied,
                        click the Submit button below. Alternately, you could
                        choose to revert to the Reviewer for more changes.
                        Complete the gray section below.
                        <br />
                        <b>
                          To substitute a different Lead MD in this review,
                        </b>{" "}
                        enter the new name at the top of the form and click{" "}
                        <b>Replace Me</b>. Your current inputs will be saved,
                        and the review will be assigned to the new person.
                      </div>
                    </div>
                    <div className={styles.row}>
                      <div className={styles.col100}>
                        <Stack
                          horizontal
                          tokens={stackTokens}
                          className={styles.stackCenter}
                        >
                          <PrimaryButton
                            text="SUBMIT TO REVIEWEE FOR ACKNOWLEDGEMENT"
                            onClick={this.onSubmit}
                            disabled={this.state.DisableSubmitButton}
                          ></PrimaryButton>
                        </Stack>
                      </div>
                    </div>
                    <div className={styles.Spacer}>&nbsp;</div>
                    <div className={styles.row}>
                      <div className={styles.col75left}>
                        <label>Optional Reversion Comment (visible)</label>
                        <TextField
                          resizable={false}
                          multiline={false}
                          value={
                            this.state.ProjectDetails.LeadMDReversionComments
                          }
                          onChange={(e, newValue) => {
                            this.onChangeTextField(
                              "LeadMDReversionComments",
                              newValue
                            );
                          }}
                        ></TextField>
                      </div>
                      <div className={styles.col25Right}>
                        <div className={styles.Spacer}>&nbsp;</div>
                        <PrimaryButton
                          text="REVERT TO REVIEWER"
                          onClick={this.onRevert}
                          disabled={this.state.DisableRevertButton}
                        ></PrimaryButton>
                      </div>
                    </div>
                  </div>
                </div>
              )
          }
          {
            // SECTION: Reviewee Acknowledgement Comments - Visible only after Lead MD approval
            this.state.ProjectDetails.StatusOfReview !=
              Config.Strings.Status_AwaitingReviewee &&
              this.state.ProjectDetails.StatusOfReview !=
                Config.Strings.Status_AwaitingReviewer &&
              this.state.ProjectDetails.StatusOfReview !=
                Config.Strings.Status_AwaitingLeadMD && (
                <div className={styles.sectionContainer}>
                  <div className={styles.sectionContent}>
                    <div className={styles.row}>
                      <div className={styles.col100}>
                        <b>REVIEWEE ACKNOWLEDGEMENT COMMENTS</b> (Comments are
                        optional and visible.)
                      </div>
                    </div>
                    <div className={styles.row}>
                      <div className={styles.col100}>
                        <TextField
                          resizable={false}
                          multiline={false}
                          disabled={
                            this.state.ProjectDetails.StatusOfReview ==
                            Config.Strings.Status_Acknowledged
                          }
                          value={
                            this.state.ProjectDetails.AcknowledgementComments
                          }
                          onChange={(e, newValue) => {
                            this.onChangeTextField(
                              "AcknowledgementComments",
                              newValue
                            );
                          }}
                        ></TextField>
                      </div>
                    </div>
                    {this.state.ProjectDetails.StatusOfReview !=
                      Config.Strings.Status_Acknowledged && (
                      <div className={styles.row}>
                        <div className={styles.col100}>
                          <Stack
                            horizontal
                            tokens={stackTokens}
                            className={styles.stackCenter}
                          >
                            <PrimaryButton
                              text="SAVE DRAFT"
                              onClick={this.onSaveAsDraft}
                              disabled={this.state.DisableSaveButton}
                            ></PrimaryButton>
                            <PrimaryButton
                              text="SUBMIT FINAL REVIEW"
                              onClick={this.onSubmit}
                              disabled={this.state.DisableSubmitButton}
                            ></PrimaryButton>
                          </Stack>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )
          }

          <div className={styles.sectionContainer}>
            <div className={styles.sectionContent}>
              <div className={styles.row}>
                <div className={styles.col100}>
                  <label>Signoff History</label>
                </div>
              </div>
              <div className={styles.row}>
                <div className={styles.col100}>
                  <TextField
                    resizable={false}
                    multiline={true}
                    readOnly={true}
                    value={this.state.ProjectDetails.SignoffHistory}
                  ></TextField>
                </div>
              </div>
            </div>
          </div>

          <div className={styles.row}>
            <div className={styles.col100right}>
              <div className={styles.Spacer}>&nbsp;</div>
              <PrimaryButton
                text="Close"
                onClick={this.onCancel}
              ></PrimaryButton>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
