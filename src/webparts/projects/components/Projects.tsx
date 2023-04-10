import * as React from "react";
import styles from "./Projects.module.scss";
import { IProjectsProps } from "./IProjectsProps";
import { IProjectsState } from "./IProjectsState";
import ListItemService from "../../../services/ListItemService";
import UserService from "../../../services/UserService";
import WebService from "../../../services/WebService";
import { TAG_ProjectDetails } from "../../../domain/models/TAG_ProjectDetails";
import { TAG_QuestionText } from "../../../domain/models/TAG_QuestionText";
import { Enums } from "../../../globals/Enums";
import { Config } from "../../../globals/Config";
import { User } from "../../../domain/models/types/User";
import { Label } from "@fluentui/react/node_modules/office-ui-fabric-react";
import PerformanceRatingScale from "./PerformanceRatingScale/PerformanceRatingScale";
import {
  Dropdown,
  IDropdownOption,
  IStackTokens,
  PrimaryButton,
  Stack,
  TextField,
  Toggle,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Parser } from "html-to-react";
import MapResult from "../../../domain/mappers/MapResult";
import DIGForm from "./DIG/DIGForm";

import { sp } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";

let QuestionArr = [
  1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18,
];

export default class Projects extends React.Component<
  IProjectsProps,
  IProjectsState
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
      CurrentUserRoles: [],
      IsLoading: true,
      AppContext: this.props.AppContext,
      DisableSaveButton: true,
      DisableRevertButton: true,
      DisableSubmitButton: true,
      ProjectDetails: new TAG_ProjectDetails(),
      OnlyVisibleForReviewer: false,
      OnlyEnableForReviewer: false,
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
      const emptyProjectDetails = await this.generateEmptyDetails();
      this.setState({
        // IsLoading: false,
        CurrentUserRoles: userRoles,
        ProjectDetails: emptyProjectDetails,
      });
    }
    // CASE: EDIT ITEM
    else {
      this.listItemService = new ListItemService(
        this.props.AppContext,
        Config.ListNames.Projects
      );
      const projectDetails: TAG_ProjectDetails =
        await this.listItemService.getItemUsingCAML(
          parseInt(this.props.ItemID),
          [],
          undefined,
          Enums.ItemResultType.TAG_ProjectDetails
        );
      this.hasEditItemPermission =
        await this.listItemService.CheckCurrentUserCanEditItem(
          parseInt(this.props.ItemID)
        );
      const allowSave: boolean = this.validateSave(projectDetails);
      const allowSubmit: boolean = this.validateSubmit(projectDetails);
      const allowRevert: boolean = this.validateRevert(projectDetails);
      const enableForReviewer: boolean =
        projectDetails.StatusOfReview == Config.Strings.Status_AwaitingReviewer;
      const visibleForReviewer: boolean =
        projectDetails.StatusOfReview == Config.Strings.Status_AwaitingReviewer;

      this.ProjectDetails(projectDetails);
      this.setState({
        IsLoading: false,
        CurrentUserRoles: userRoles,
        // ProjectDetails: projectDetails,
        DisableSaveButton: !allowSave,
        DisableSubmitButton: !allowSubmit,
        DisableRevertButton: !allowRevert,
        OnlyEnableForReviewer: enableForReviewer,
        OnlyVisibleForReviewer: visibleForReviewer,
      });
    }
  }

  //#region "Control Change Events"

  private async ProjectDetails(data: any) {
    let tempProject = data;

    /* Deva changes start */
    let TempJobTitle: string = "";
    if (data.JobTitle == "Manager I") {
      TempJobTitle = "Manager";
    } else if (data.JobTitle == "Director II") {
      TempJobTitle = "Director";
    } else {
      TempJobTitle = data.JobTitle;
    }
    let camlFilterConditions = `<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>${TempJobTitle}</Value></Eq></Where>`;
    /* Deva changes end */

    this.listItemService = new ListItemService(
      this.props.AppContext,
      Config.ListNames.QuestionText
    );
    let questionDetails = await this.listItemService.getItemsUsingCAML(
      [],
      undefined,
      camlFilterConditions,
      100,
      Enums.ItemResultType.TAG_QuestionText
    );
    questionDetails = questionDetails.filter((arr) => {
      return arr.ServiceLine == data.ServiceLine;
    });
    if (questionDetails.length > 0) {
      QuestionArr.map((i) => {
        tempProject["Q" + i + "Category"] =
          questionDetails[0]["Q" + i + "Category"];

        if (questionDetails[0]["Q" + i + "IsRating"] == "Yes") {
          tempProject["Q" + i + "IsRating"] = true;
        } else {
          tempProject["Q" + i + "IsRating"] = false;
        }

        // if (
        //   Number(tempProject["Q" + i]) > 0 ||
        //   tempProject["Q" + i] == "N/A" ||
        //   tempProject["Q" + i] == ""
        // ) {
        //   tempProject["Q" + i + "IsRating"] = true;
        // } else {
        //   tempProject["Q" + i + "IsRating"] = false;
        // }
      });

      this.setState({
        IsLoading: false,
        ProjectDetails: tempProject,
      });
    } else {
      this.setState({
        IsLoading: false,
        ProjectDetails: tempProject,
      });
    }
  }

  private onRenderHTML(type) {
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

    return (
      <>
        {QuestionArr.map((i) => {
          return (
            this.state.ProjectDetails["Q" + i + "Category"] == type &&
            this.state.ProjectDetails["Q" + i + "Text"].replace(
              /<[^>]*>/g,
              ""
            ) != "Text" &&
            this.state.ProjectDetails["Q" + i + "Text"].replace(
              /<[^>]*>/g,
              ""
            ) != "" && (
              <>
                <div className={styles.row}>
                  <div className={styles.col80left}>
                    <div>
                      {Parser().parse(
                        this.state.ProjectDetails["Q" + i + "Text"]
                      )}
                    </div>
                  </div>
                  <div className={styles.col20left}>
                    {/* <div>
                      <Toggle
                        label={
                          <div
                            style={{
                              position: "absolute",
                              left: "7px",
                              top: "4px",
                              width: "200px",
                            }}
                          >
                            Rating
                          </div>
                        }
                        disabled={!this.state.OnlyEnableForReviewer}
                        inlineLabel
                        checked={
                          this.state.ProjectDetails["Q" + i + "IsRating"]
                        }
                        style={{ transform: "translateX(100px)" }}
                        onChange={() => {
                          this.onChangeToggleValues(
                            "Q" + i + "IsRating",
                            !this.state.ProjectDetails["Q" + i + "IsRating"]
                          );
                        }}
                      />
                    </div> */}
                    <div className={styles.Spacer}>&nbsp;</div>
                    <Dropdown
                      placeholder="Select"
                      options={
                        this.state.ProjectDetails["Q" + i + "IsRating"]
                          ? questionChoices
                          : question7Choices
                      }
                      disabled={!this.state.OnlyEnableForReviewer}
                      selectedKey={this.state.ProjectDetails["Q" + i]}
                      onChange={(e, selectedOption) => {
                        this.onChangeDropdownValues(
                          "Q" + i,
                          selectedOption.text
                        );
                      }}
                    />
                  </div>
                </div>
                <div className={styles.Spacer}>&nbsp;</div>
              </>
            )
          );
        })}
      </>
    );
  }

  // On change of toggle values
  private onChangeToggleValues(fieldName: string, newValue: boolean): void {
    let curretState = this.state.ProjectDetails;
    curretState[fieldName] = newValue;
    this.onFormFieldValueChange(curretState);
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

    //! Technorucs
    data[columns.Q13] = projectDetails.Q13;
    data[columns.Q14] = projectDetails.Q14;
    data[columns.Q15] = projectDetails.Q15;
    data[columns.Q16] = projectDetails.Q16;
    data[columns.Q17] = projectDetails.Q17;
    data[columns.Q18] = projectDetails.Q18;

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

    // ContentType - TAG Employee
    data["ContentTypeId"] =
      "0x0100632AB58492E4E44FB1DA1F844F6346AB00EC81959E6ACFF8489A8CBBAEC10CF7C2";

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
      this.gotoListPage();
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
    let currentStatusOfReview = this.state.ProjectDetails.StatusOfReview;
    let result =
      this.props.AppContext.pageContext.web.absoluteUrl +
      Config.Links.RevieweeLink;
    switch (currentStatusOfReview) {
      case Config.Strings.Status_AwaitingReviewee:
        result =
          this.props.AppContext.pageContext.web.absoluteUrl +
          Config.Links.ReviewerLink; // Reviewee Approved/Responded
        break;
      case Config.Strings.Status_AwaitingReviewer:
        result =
          this.props.AppContext.pageContext.web.absoluteUrl +
          Config.Links.LeadMDLink; // Reviewer Approved
        break;
      case Config.Strings.Status_AwaitingLeadMD:
        result =
          this.props.AppContext.pageContext.web.absoluteUrl +
          // Config.Links.ProjectsListAllItems; // Lead MD Approved
          Config.Links.HomePageLink;
        break;
      case Config.Strings.Status_AwaitingAcknowledgement:
        result =
          this.props.AppContext.pageContext.web.absoluteUrl +
          Config.Links.RevieweeLink; // Acknowledged by Reviewee
        break;
    }

    // let returnURL = this.props.AppContext.pageContext.web.absoluteUrl + Config.Links.ProjectsListAllItems;
    window.location.href = result;
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
      Q1Category: "",
      Q2Category: "",
      Q3Category: "",
      Q4Category: "",
      Q5Category: "",
      Q6Category: "",
      Q7Category: "",
      Q8Category: "",
      Q9Category: "",
      Q10Category: "",
      Q11Category: "",
      Q12Category: "",
      Q13Category: "",
      Q14Category: "",
      Q15Category: "",
      Q16Category: "",
      Q17Category: "",
      Q18Category: "",

      Q1IsRating: false,
      Q2IsRating: false,
      Q3IsRating: false,
      Q4IsRating: false,
      Q5IsRating: false,
      Q6IsRating: false,
      Q7IsRating: false,
      Q8IsRating: false,
      Q9IsRating: false,
      Q10IsRating: false,
      Q11IsRating: false,
      Q12IsRating: false,
      Q13IsRating: false,
      Q14IsRating: false,
      Q15IsRating: false,
      Q16IsRating: false,
      Q17IsRating: false,
      Q18IsRating: false,
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
      if (
        details.Reviewer.Id == null ||
        details.LeadMD.Id == null ||
        details.StrongPerformance == "" ||
        details.DevelopmentAreas == "" ||
        details.NeededSkills == ""
      ) {
        valid = false;
      }
    }

    // Validations required when status is "Awaiting Reviewer"
    if (details.StatusOfReview == Config.Strings.Status_AwaitingReviewer) {
      if (
        details.Reviewer.Id == null ||
        details.LeadMD.Id == null ||
        (details.Q1 == "" &&
          (details.Q1Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q1Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q2 == "" &&
          (details.Q2Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q2Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q3 == "" &&
          (details.Q3Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q3Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q4 == "" &&
          (details.Q4Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q4Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q5 == "" &&
          (details.Q5Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q5Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q6 == "" &&
          (details.Q6Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q6Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q7 == "" &&
          (details.Q7Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q7Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q8 == "" &&
          (details.Q8Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q8Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q9 == "" &&
          (details.Q9Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q9Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q10 == "" &&
          (details.Q10Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q10Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q11 == "" &&
          (details.Q11Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q11Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q12 == "" &&
          (details.Q12Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q12Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q13 == "" &&
          (details.Q13Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q13Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q14 == "" &&
          (details.Q14Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q14Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q15 == "" &&
          (details.Q15Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q15Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q16 == "" &&
          (details.Q16Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q16Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q17 == "" &&
          (details.Q17Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q17Text.replace(/<[^>]*>/g, "") == "")) ||
        (details.Q18 == "" &&
          (details.Q18Text.replace(/<[^>]*>/g, "") == "Text" ||
            details.Q18Text.replace(/<[^>]*>/g, "") == "")) ||
        // details.Q1 == "" ||
        // details.Q2 == "" ||
        // details.Q3 == "" ||
        // details.Q4 == "" ||
        // details.Q5 == "" ||
        // details.Q6 == "" ||
        // details.Q7 == "" ||
        // details.Q8 == "" ||
        // details.Q9 == "" ||
        // details.Q10 == "" ||
        // details.Q11 == "" ||
        // details.Q12 == "" ||
        // details.Q13 == "" ||
        // details.Q14 == "" ||
        // details.Q15 == "" ||
        // details.Q16 == "" ||
        // details.Q17 == "" ||
        // details.Q18 == "" ||
        details.Complexity == "" ||
        details.StrongPerformance == "" ||
        details.DevelopmentAreas == "" ||
        details.NeededSkills == ""
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

  public render(): React.ReactElement<IProjectsProps> {
    const complexityOptions: IDropdownOption[] = [
      { key: "Easy", text: "Easy" },
      { key: "Moderate", text: "Moderate" },
      { key: "Difficult", text: "Difficult" },
    ];

    const fiscalYearOptions: IDropdownOption[] = [
      {
        key: new Date().getFullYear().toString(),
        text: new Date().getFullYear().toString(),
      },
    ];

    const serviceLineOptions: IDropdownOption[] = [
      { key: "", text: "" },
      { key: "Financial Due Diligence", text: "Financial Due Diligence" },
      {
        key: "Global Transaction Analytics",
        text: "Global Transaction Analytics",
      },
      {
        key: "Capital Markets & Accounting Advisory",
        text: "Capital Markets & Accounting Advisory",
      },
      { key: "Data Intelligence Gateway", text: "Data Intelligence Gateway" },
      {
        key: "ESG",
        text: "ESG",
      },
    ];

    // const questionChoices: IDropdownOption[] = [
    //   { key: "N/A", text: "N/A" },
    //   { key: "1", text: "1" },
    //   { key: "2", text: "2" },
    //   { key: "3", text: "3" },
    //   { key: "4", text: "4" },
    // ];

    // const question7Choices: IDropdownOption[] = [
    //   { key: "N/A", text: "N/A" },
    //   { key: "Demonstrates", text: "Demonstrates" },
    //   { key: "Does not demonstrate", text: "Does not demonstrate" },
    // ];

    const stackTokens: IStackTokens = { childrenGap: 20 };

    return (
      <div className={styles.projects}>
        <div className={styles.container}>
          {this.state.IsLoading == false ? (
            <>
              <div className={styles.logoImg} title="logo"></div>
              <div className={styles.sectionContainer}>
                <div className={styles.sectionHeader}>
                  <div className={styles.colHeader100}></div>
                </div>
                {
                  // View when the Review is not started or declined or soft deleted
                  (this.state.ProjectDetails.StatusOfReview ==
                    Config.Strings.Status_NotStarted ||
                    this.state.ProjectDetails.StatusOfReview ==
                      Config.Strings.Status_Declined ||
                    this.state.ProjectDetails.StatusOfReview ==
                      Config.Strings.Status_SoftDeleted ||
                    this.state.ProjectDetails.StatusOfReview ==
                      Config.Strings.Status_Split ||
                    this.state.ProjectDetails.StatusOfReview ==
                      Config.Strings.Status_Combined) && (
                    <div className={styles.sectionContent}>
                      <div className={styles.row}>
                        <div className={styles.col50left}>
                          <div className={styles.lblLeftTitle}>Reviewee : </div>{" "}
                          {this.state.ProjectDetails.Reviewee.Title}
                        </div>
                        <div
                          className={styles.col50left}
                          style={{
                            display: "flex",
                            alignItems: "center",
                          }}
                        >
                          <div className={styles.lblRightTitle}>
                            Fiscal Year :{" "}
                          </div>
                          <div
                            style={{
                              width: "40%",
                            }}
                          >
                            <Dropdown
                              placeholder="Select Fiscal Year"
                              options={fiscalYearOptions}
                              selectedKey={this.state.ProjectDetails.FiscalYear}
                              onChange={(e, selectedOption) => {
                                this.onChangeDropdownValues(
                                  "FiscalYear",
                                  selectedOption.text
                                );
                              }}
                            />
                          </div>
                        </div>
                      </div>
                      <div className={styles.row}>
                        <div className={styles.col50left}>
                          <div className={styles.lblLeftTitle}>
                            Project Name :
                          </div>{" "}
                          {this.state.ProjectDetails.ProjectName}
                        </div>
                        <div
                          className={styles.col50left}
                          style={{
                            display: "flex",
                            alignItems: "center",
                          }}
                        >
                          <div className={styles.lblRightTitle}>
                            Service Line
                            <span className={styles.inRed}> * </span>:
                          </div>
                          <div
                            style={{
                              width: "40%",
                            }}
                          >
                            <Dropdown
                              placeholder="Select Service line"
                              options={serviceLineOptions}
                              selectedKey={
                                this.state.ProjectDetails.ServiceLine
                              }
                              onChange={(e, selectedOption) => {
                                this.onChangeDropdownValues(
                                  "ServiceLine",
                                  selectedOption.text
                                );
                              }}
                            />
                          </div>
                        </div>
                      </div>
                      <div className={styles.row}>
                        <div className={styles.col50left}>
                          <div className={styles.lblLeftTitle}>
                            Project Code :
                          </div>{" "}
                          {this.state.ProjectDetails.ProjectCode}
                        </div>
                        <div className={styles.col50left}>
                          <div className={styles.lblRightTitle}>
                            Performance Period :
                          </div>{" "}
                          {this.state.ProjectDetails.ProjectStartDateFormatted}{" "}
                          - {this.state.ProjectDetails.ProjectEndDateFormatted}
                        </div>
                      </div>
                      <div className={styles.SpacerSmall}>&nbsp;</div>
                      <div className={styles.row}>
                        <div className={styles.col50left}>
                          <div className={styles.lblLeftTitle}>Client :</div>{" "}
                          {this.state.ProjectDetails.ClientName}
                        </div>
                        <div className={styles.col50left}>
                          <div className={styles.lblRightTitle}>
                            Hours Worked :
                          </div>{" "}
                          {this.state.ProjectDetails.HoursWorked}
                        </div>
                      </div>
                      <div className={styles.SpacerSmall}>&nbsp;</div>
                      <div className={styles.row}>
                        <div className={styles.col50left}>
                          <div className={styles.lblLeftTitle}>
                            Home Office :
                          </div>{" "}
                          {this.state.ProjectDetails.HomeOffice}
                        </div>
                        <div className={styles.col50left}>
                          <div className={styles.lblRightTitle}>Job Role :</div>{" "}
                          {this.state.ProjectDetails.JobTitle}
                        </div>
                      </div>
                      <div className={styles.Spacer}>&nbsp;</div>
                      <div className={styles.row}>
                        <div className={styles.col100}>
                          <b>REVIEWEE: </b> To initiate a review, indicate the
                          Reviewer and confirm the Lead MD below. Choose the
                          Fiscal Year and Service Line at the top and then click{" "}
                          <b>Start Review</b>.
                          <br />
                          <ul>
                            <li>
                              <b>Combined reviews:</b> If this review is to be
                              combined with other reviews,{" "}
                              <b className={styles.inRed}>do not</b> start it
                              here. Click "Combine Reviews" on the left-hand
                              side.
                            </li>
                          </ul>
                        </div>
                      </div>
                      <div className={styles.row}>
                        <div className={styles.col100}>
                          <div className={styles.highlightedInstruction}>
                            {this.state.ProjectDetails.StatusOfReview ==
                              Config.Strings.Status_Split && (
                              <b>
                                This review was split into at least one
                                additional review.
                              </b>
                            )}
                            {this.state.ProjectDetails.StatusOfReview ==
                              Config.Strings.Status_Combined && (
                              <b>
                                This review is now a part of a Combined Review.
                              </b>
                            )}
                            {this.state.ProjectDetails.StatusOfReview ==
                              Config.Strings.Status_Declined && (
                              <b>
                                This review was declined by{" "}
                                {this.state.ProjectDetails.ModifiedBy.Title} on{" "}
                                {this.state.ProjectDetails.ModifiedOnFormatted}
                              </b>
                            )}
                          </div>
                        </div>
                      </div>

                      {this.state.ProjectDetails.StatusOfReview !=
                        Config.Strings.Status_Declined &&
                      this.state.ProjectDetails.StatusOfReview !=
                        Config.Strings.Status_Split &&
                      this.state.ProjectDetails.StatusOfReview !=
                        Config.Strings.Status_Combined ? (
                        <div className={styles.row}>
                          <div className={styles.col25left}>
                            <label>
                              <b>Reviewer </b>
                              <span className={styles.inRed}>*</span>
                            </label>
                            <PeoplePicker
                              context={this.props.AppContext as any}
                              personSelectionLimit={1}
                              groupName={""} // Leave this blank in case you want to filter from all users
                              showtooltip={true}
                              required={true}
                              ensureUser={true}
                              showHiddenInUI={false}
                              principalTypes={[PrincipalType.User]}
                              defaultSelectedUsers={[
                                this.state.ProjectDetails.Reviewer.Email,
                              ]}
                              onChange={(selected) => {
                                this.onChangePersonField("Reviewer", selected);
                              }}
                              resolveDelay={1000}
                            />
                          </div>
                          <div className={styles.col25left}>
                            <label>
                              <b>Lead MD </b>
                              <span className={styles.inRed}>*</span>
                            </label>
                            <PeoplePicker
                              context={this.props.AppContext as any}
                              personSelectionLimit={1}
                              groupName={""} // Leave this blank in case you want to filter from all users
                              showtooltip={true}
                              required={true}
                              ensureUser={true}
                              showHiddenInUI={false}
                              principalTypes={[PrincipalType.User]}
                              defaultSelectedUsers={[
                                this.state.ProjectDetails.LeadMD.Email,
                              ]}
                              onChange={(selected) => {
                                this.onChangePersonField("LeadMD", selected);
                              }}
                              resolveDelay={1000}
                            />
                          </div>
                          <div className={styles.col50left}>
                            <br />
                            <Stack
                              horizontal
                              tokens={stackTokens}
                              className={styles.stackCenter}
                            >
                              <PrimaryButton
                                text="START REVIEW"
                                onClick={this.onSubmit}
                                disabled={this.state.DisableSubmitButton}
                              ></PrimaryButton>
                              <PrimaryButton
                                text="Cancel"
                                onClick={this.onCancel}
                              ></PrimaryButton>
                            </Stack>
                          </div>
                        </div>
                      ) : (
                        <div className={styles.row}>
                          <div className={styles.col100right}>
                            <PrimaryButton
                              text="Cancel"
                              onClick={this.onCancel}
                            ></PrimaryButton>
                          </div>
                        </div>
                      )}
                    </div>
                  )
                }
                {
                  // View when the Review is started and acknowledged
                  (this.state.ProjectDetails.StatusOfReview ==
                    Config.Strings.Status_Acknowledged ||
                    this.state.ProjectDetails.StatusOfReview ==
                      Config.Strings.Status_AwaitingAcknowledgement ||
                    this.state.ProjectDetails.StatusOfReview ==
                      Config.Strings.Status_AwaitingLeadMD ||
                    this.state.ProjectDetails.StatusOfReview ==
                      Config.Strings.Status_AwaitingReviewer ||
                    this.state.ProjectDetails.StatusOfReview ==
                      Config.Strings.Status_AwaitingReviewee) && (
                    <div className={styles.sectionContent}>
                      <div className={styles.row}>
                        <div className={styles.col50left}>
                          <div className={styles.lblLeftTitle}>Reviewee : </div>{" "}
                          {this.state.ProjectDetails.Reviewee.Title}
                        </div>
                        <div className={styles.col50left}>
                          <div className={styles.lblRightTitle}>
                            Performance Period :
                          </div>{" "}
                          {this.state.ProjectDetails.ProjectStartDateFormatted}{" "}
                          - {this.state.ProjectDetails.ProjectEndDateFormatted}
                        </div>
                      </div>
                      <div className={styles.row}>
                        <div className={styles.col50left}>
                          <div className={styles.lblLeftTitle}>
                            Home Office :
                          </div>{" "}
                          {this.state.ProjectDetails.HomeOffice}
                        </div>
                        <div className={styles.col50left}>
                          <div className={styles.lblRightTitle}>Job Role :</div>{" "}
                          {this.state.ProjectDetails.JobTitle}
                        </div>
                      </div>
                      <div className={styles.row}>
                        <div className={styles.col50left}>
                          <div className={styles.lblLeftTitle}>
                            Project Name :
                          </div>{" "}
                          {this.state.ProjectDetails.ProjectName}
                        </div>
                        <div className={styles.col50left}>
                          <div className={styles.lblRightTitle}>
                            Date Review Originated :
                          </div>{" "}
                          {this.state.ProjectDetails.DateOriginatedFormatted}
                        </div>
                      </div>
                      <div className={styles.row}>
                        <div className={styles.col50left}>
                          <div className={styles.lblLeftTitle}>Client : </div>{" "}
                          {this.state.ProjectDetails.ClientName}
                        </div>
                        <div className={styles.col50left}>
                          <div className={styles.lblRightTitle}>
                            Date Review Completed :
                          </div>{" "}
                          {
                            this.state.ProjectDetails
                              .DateReviewCompletedFormatted
                          }
                        </div>
                      </div>
                      <div className={styles.row}>
                        <div className={styles.col50left}>
                          <div className={styles.lblLeftTitle}>
                            Fiscal Year :{" "}
                          </div>{" "}
                          {this.state.ProjectDetails.FiscalYear}
                        </div>
                        <div className={styles.col50left}>
                          <div className={styles.lblRightTitle}>
                            Service Line :
                          </div>{" "}
                          {this.state.ProjectDetails.ServiceLine}
                        </div>
                      </div>
                      <div className={styles.row}>
                        <div className={styles.col100}>
                          <div className={styles.highlightedInstruction}>
                            <b>
                              REVIEWEES: Please go directly to the Commentary
                              section at the bottom.
                            </b>
                            <br />
                            The Reviewer will determine the complexity & ratings
                            by competency.
                          </div>
                        </div>
                      </div>
                      <div className={styles.Spacer}>&nbsp;</div>
                      <div className={styles.row}>
                        <div className={styles.col25left}>
                          <b>Hours worked :</b>
                          {/* </div>
                        <div className={styles.col25left}> */}
                          <TextField
                            value={
                              this.state.ProjectDetails.HoursWorked != null
                                ? this.state.ProjectDetails.HoursWorked.toString()
                                : ""
                            }
                            disabled={!this.state.OnlyEnableForReviewer}
                            onChange={(e, newValue) => {
                              this.onChangeTextField("HoursWorked", newValue);
                            }}
                          />
                        </div>
                        <div className={styles.col25left}>
                          <b>Complexity :</b>
                          <span className={styles.inRed}>*</span>
                          {/* </div>
                        <div className={styles.col25left}> */}
                          <Dropdown
                            placeholder="Select Complexity"
                            options={complexityOptions}
                            disabled={!this.state.OnlyEnableForReviewer}
                            selectedKey={this.state.ProjectDetails.Complexity}
                            onChange={(e, selectedOption) => {
                              this.onChangeDropdownValues(
                                "Complexity",
                                selectedOption.text
                              );
                            }}
                          />
                        </div>
                      </div>
                      {/* <div className={styles.row}>
                        <div className={styles.col10left}>
                          <b>Complexity :</b>{" "}
                          <span className={styles.inRed}>*</span>
                        </div>
                        <div className={styles.col25left}>
                          <Dropdown
                            placeholder="Select Complexity"
                            options={complexityOptions}
                            disabled={!this.state.OnlyEnableForReviewer}
                            selectedKey={this.state.ProjectDetails.Complexity}
                            onChange={(e, selectedOption) => {
                              this.onChangeDropdownValues(
                                "Complexity",
                                selectedOption.text
                              );
                            }}
                          />
                        </div>
                      </div> */}

                      <React.Fragment>
                        <div className={styles.Spacer}>&nbsp;</div>
                        <div className={styles.row}>
                          {(this.state.ProjectDetails.StatusOfReview ==
                            Config.Strings.Status_AwaitingReviewer ||
                            this.state.ProjectDetails.StatusOfReview ==
                              Config.Strings.Status_AwaitingReviewee ||
                            this.state.ProjectDetails.StatusOfReview ==
                              Config.Strings.Status_AwaitingLeadMD ||
                            this.state.ProjectDetails.StatusOfReview ==
                              Config.Strings.Status_Acknowledged) && (
                            <React.Fragment>
                              <div className={styles.col25left}>
                                <label>
                                  <b>Reviewer :</b>
                                </label>
                                <PeoplePicker
                                  context={this.props.AppContext as any}
                                  personSelectionLimit={1}
                                  groupName={""} // Leave this blank in case you want to filter from all users
                                  showtooltip={true}
                                  required={true}
                                  ensureUser={true}
                                  disabled={true}
                                  showHiddenInUI={false}
                                  principalTypes={[PrincipalType.User]}
                                  defaultSelectedUsers={[
                                    this.state.ProjectDetails.Reviewer.Email,
                                  ]}
                                  resolveDelay={1000}
                                />
                              </div>
                              <div className={styles.col25left}>
                                <label>
                                  <b>Lead MD :</b>
                                </label>
                                <PeoplePicker
                                  context={this.props.AppContext as any}
                                  personSelectionLimit={1}
                                  groupName={""} // Leave this blank in case you want to filter from all users
                                  showtooltip={true}
                                  required={true}
                                  ensureUser={true}
                                  disabled={true}
                                  showHiddenInUI={false}
                                  principalTypes={[PrincipalType.User]}
                                  defaultSelectedUsers={[
                                    this.state.ProjectDetails.LeadMD.Email,
                                  ]}
                                  resolveDelay={1000}
                                />
                              </div>
                            </React.Fragment>
                          )}
                          <div
                            className={
                              this.state.ProjectDetails.StatusOfReview ==
                                Config.Strings.Status_AwaitingReviewer ||
                              this.state.ProjectDetails.StatusOfReview ==
                                Config.Strings.Status_AwaitingReviewee ||
                              this.state.ProjectDetails.StatusOfReview ==
                                Config.Strings.Status_AwaitingLeadMD ||
                              this.state.ProjectDetails.StatusOfReview ==
                                Config.Strings.Status_Acknowledged
                                ? styles.col25Right
                                : styles.col75right
                            }
                          ></div>
                          <div className={styles.col25left}>
                            <label>
                              <b>Review Status :</b>
                            </label>
                            <br />
                            <TextField
                              disabled={true}
                              value={this.state.ProjectDetails.StatusOfReview}
                            />
                          </div>
                        </div>
                      </React.Fragment>
                      {(this.state.ProjectDetails.StatusOfReview ==
                        Config.Strings.Status_AwaitingReviewer ||
                        this.state.ProjectDetails.StatusOfReview ==
                          Config.Strings.Status_AwaitingLeadMD) && (
                        <React.Fragment>
                          <div className={styles.Spacer}>&nbsp;</div>
                          <div className={styles.row}>
                            <div className={styles.col35left}>
                              <div className={styles.row}>
                                <div className={styles.col50left}>
                                  <PeoplePicker
                                    context={this.props.AppContext as any}
                                    personSelectionLimit={1}
                                    groupName={""} // Leave this blank in case you want to filter from all users
                                    showtooltip={true}
                                    required={true}
                                    ensureUser={true}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    defaultSelectedUsers={[
                                      this.state.ProjectDetails.SubstituteUser
                                        .Email,
                                    ]}
                                    onChange={(selected) => {
                                      this.onChangePersonField(
                                        "SubstituteUser",
                                        selected
                                      );
                                    }}
                                    resolveDelay={1000}
                                  />
                                </div>
                                <div className={styles.col50left}>
                                  <PrimaryButton
                                    text="REPLACE ME"
                                    onClick={this.onReplaceMe}
                                    disabled={
                                      !this.hasEditItemPermission ||
                                      this.state.ProjectDetails.SubstituteUser
                                        .Id == null
                                    }
                                  ></PrimaryButton>
                                </div>
                              </div>
                            </div>
                            <div className={styles.col50left}>
                              <b>Should you be reviewing this person?</b> If
                              not, enter your replacement in the box at left and
                              click <b>Replace Me</b>. The review's current
                              status will be saved, and your replacement will
                              pick up where you left off.
                            </div>
                          </div>
                        </React.Fragment>
                      )}
                      <div className={styles.Spacer}>&nbsp;</div>

                      <PerformanceRatingScale
                        AppContext={this.props.AppContext}
                        IsLoading={false}
                      ></PerformanceRatingScale>
                      {!(
                        this.state.ProjectDetails.ServiceLine ==
                        "Data Intelligence Gateway"
                      ) && (
                        <div>
                          <div className={styles.sectionContainer}>
                            <div className={styles.sectionHeader}>
                              <div className={styles.colHeader100}>
                                <span className={styles.subTitle}>
                                  SERVICE EXCELLENCE
                                </span>
                              </div>
                            </div>
                            <div className={styles.sectionContent}>
                              {this.onRenderHTML("SERVICE EXCELLENCE")}

                              {/* <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q1Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={questionChoices}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  selectedKey={this.state.ProjectDetails.Q1}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q1",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div>
                            <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q2Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={questionChoices}
                                  selectedKey={this.state.ProjectDetails.Q2}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q2",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div> */}
                            </div>
                          </div>

                          <div className={styles.sectionContainer}>
                            <div className={styles.sectionHeader}>
                              <div className={styles.colHeader100}>
                                <span className={styles.subTitle}>
                                  FOUNDATIONAL EXPERTISE
                                </span>
                              </div>
                            </div>
                            <div className={styles.sectionContent}>
                              {this.onRenderHTML("FOUNDATIONAL EXPERTISE")}
                              {/* <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q3Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={questionChoices}
                                  selectedKey={this.state.ProjectDetails.Q3}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q3",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div>
                            <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q4Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={questionChoices}
                                  selectedKey={this.state.ProjectDetails.Q4}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q4",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div>
                            <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q5Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={questionChoices}
                                  selectedKey={this.state.ProjectDetails.Q5}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q5",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div>
                            <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q6Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={questionChoices}
                                  selectedKey={this.state.ProjectDetails.Q6}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q6",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div> */}
                            </div>
                          </div>

                          <div className={styles.sectionContainer}>
                            <div className={styles.sectionHeader}>
                              <div className={styles.colHeader100}>
                                <span className={styles.subTitle}>
                                  PRACTICE OPERATIONS & LEADERSHIP
                                </span>
                              </div>
                            </div>
                            <div className={styles.sectionContent}>
                              {this.onRenderHTML(
                                "PRACTICE OPERATIONS & LEADERSHIP"
                              )}
                              {/* <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q7Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={question7Choices}
                                  selectedKey={this.state.ProjectDetails.Q7}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q7",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div>
                            <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q8Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={questionChoices}
                                  selectedKey={this.state.ProjectDetails.Q8}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q8",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div>
                            <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q9Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={questionChoices}
                                  selectedKey={this.state.ProjectDetails.Q9}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q9",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div> */}
                            </div>
                          </div>

                          <div className={styles.sectionContainer}>
                            <div className={styles.sectionHeader}>
                              <div className={styles.colHeader100}>
                                <span className={styles.subTitle}>
                                  PERSONAL EFFECTIVENESS
                                </span>
                              </div>
                            </div>
                            <div className={styles.sectionContent}>
                              {this.onRenderHTML("PERSONAL EFFECTIVENESS")}
                              {/* <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q10Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={questionChoices}
                                  selectedKey={this.state.ProjectDetails.Q10}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q10",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div>
                            <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q11Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={questionChoices}
                                  selectedKey={this.state.ProjectDetails.Q11}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q11",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div>
                            <div className={styles.row}>
                              <div className={styles.col80left}>
                                <div>
                                  {Parser().parse(
                                    this.state.ProjectDetails.Q12Text
                                  )}
                                </div>
                              </div>
                              <div className={styles.col20left}>
                                <div className={styles.Spacer}>&nbsp;</div>
                                <Dropdown
                                  placeholder="Select"
                                  options={question7Choices}
                                  selectedKey={this.state.ProjectDetails.Q12}
                                  disabled={!this.state.OnlyEnableForReviewer}
                                  onChange={(e, selectedOption) => {
                                    this.onChangeDropdownValues(
                                      "Q12",
                                      selectedOption.text
                                    );
                                  }}
                                />
                              </div>
                            </div>
                            <div className={styles.Spacer}>&nbsp;</div> */}
                            </div>
                          </div>

                          <div className={styles.sectionContainer}>
                            <div className={styles.sectionHeader}>
                              <div className={styles.colHeader100}></div>
                            </div>
                            <div className={styles.sectionContent}>
                              <div className={styles.row}>
                                <div className={styles.col100}>
                                  <Label>
                                    Briefly comment on the Reviewee's top areas
                                    of strong performance on this project and
                                    following the RDTA principles. Your comments
                                    should support, at minimum, any 4 ratings
                                    above. <i>(Commentary required)</i>
                                  </Label>
                                  <TextField
                                    resizable={false}
                                    multiline={true}
                                    value={
                                      this.state.ProjectDetails
                                        .StrongPerformance
                                    }
                                    disabled={
                                      this.state.ProjectDetails
                                        .StatusOfReview ==
                                        Config.Strings.Status_AwaitingLeadMD ||
                                      this.state.ProjectDetails
                                        .StatusOfReview ==
                                        Config.Strings
                                          .Status_AwaitingAcknowledgement ||
                                      this.state.ProjectDetails
                                        .StatusOfReview ==
                                        Config.Strings.Status_Acknowledged
                                    }
                                    onChange={(e, newValue) => {
                                      this.onChangeTextField(
                                        "StrongPerformance",
                                        newValue
                                      );
                                    }}
                                  ></TextField>
                                </div>
                              </div>
                              <div className={styles.row}>
                                <div className={styles.col100}>
                                  <Label>
                                    Briefly comment on the Reviewee's top areas
                                    for development with the RDTA principles.
                                    Your comments should support, at minimum,
                                    any 1 rating above.{" "}
                                    <i>(Commentary required)</i>
                                  </Label>
                                  <TextField
                                    resizable={false}
                                    multiline={true}
                                    value={
                                      this.state.ProjectDetails.DevelopmentAreas
                                    }
                                    disabled={
                                      this.state.ProjectDetails
                                        .StatusOfReview ==
                                        Config.Strings.Status_AwaitingLeadMD ||
                                      this.state.ProjectDetails
                                        .StatusOfReview ==
                                        Config.Strings
                                          .Status_AwaitingAcknowledgement ||
                                      this.state.ProjectDetails
                                        .StatusOfReview ==
                                        Config.Strings.Status_Acknowledged
                                    }
                                    onChange={(e, newValue) => {
                                      this.onChangeTextField(
                                        "DevelopmentAreas",
                                        newValue
                                      );
                                    }}
                                  ></TextField>
                                </div>
                              </div>
                              <div className={styles.row}>
                                <div className={styles.col100}>
                                  <Label>
                                    Briefly comment on what RDTA skills are
                                    necessary for the reviewee to continue to
                                    develop in order to progress in their
                                    career. <i>(Commentary required)</i>
                                  </Label>
                                  <TextField
                                    resizable={false}
                                    multiline={true}
                                    disabled={
                                      this.state.ProjectDetails
                                        .StatusOfReview ==
                                        Config.Strings.Status_AwaitingLeadMD ||
                                      this.state.ProjectDetails
                                        .StatusOfReview ==
                                        Config.Strings
                                          .Status_AwaitingAcknowledgement ||
                                      this.state.ProjectDetails
                                        .StatusOfReview ==
                                        Config.Strings.Status_Acknowledged
                                    }
                                    value={
                                      this.state.ProjectDetails.NeededSkills
                                    }
                                    onChange={(e, newValue) => {
                                      this.onChangeTextField(
                                        "NeededSkills",
                                        newValue
                                      );
                                    }}
                                  ></TextField>
                                </div>
                              </div>
                              <div className={styles.row}>
                                <div className={styles.col100}>
                                  <Label>
                                    Additional comments from Lead MD{" "}
                                    <i>(optional)</i>
                                  </Label>
                                  <TextField
                                    resizable={false}
                                    multiline={true}
                                    disabled={
                                      this.state.ProjectDetails
                                        .StatusOfReview !=
                                      Config.Strings.Status_AwaitingLeadMD
                                    }
                                    value={
                                      this.state.ProjectDetails.LeadMDComments
                                    }
                                    onChange={(e, newValue) => {
                                      this.onChangeTextField(
                                        "LeadMDComments",
                                        newValue
                                      );
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
                                        <b>REVIEWEE:</b> When your comments are
                                        complete, click the Submit button below.
                                        To identify a different Reviewer or Lead
                                        MD to perform this review, change the
                                        corresponding field(s) at the top of
                                        this form before submitting. (Not ready
                                        yet? You can <b>Save Draft</b> to
                                        preserve your inputs prior to submitting
                                        to the Reviewer.)
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
                                            disabled={
                                              this.state.DisableSaveButton
                                            }
                                          ></PrimaryButton>
                                          <PrimaryButton
                                            text="SUBMIT TO REVIEWER FOR APPROVAL"
                                            className={
                                              !this.state.DisableSubmitButton
                                                ? styles.btnApprovedForReviewerGreen
                                                : ""
                                            }
                                            onClick={this.onSubmit}
                                            disabled={
                                              this.state.DisableSubmitButton
                                            }
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
                                      <b>REVIEWER:</b> When you are ready to
                                      advance the review to the Lead MD, click
                                      the Submit button below. Click Save Draft
                                      to save and return later. To revert back
                                      to the Reviewee, complete the gray section
                                      below.
                                      <br />
                                      <b>
                                        To substitute a different Reviewer in
                                        this review,
                                      </b>{" "}
                                      enter the new name at the top of the form
                                      and click <b>Replace Me.</b> Your current
                                      inputs will be saved, and the review will
                                      be assigned to the new person.
                                      <br />
                                      <b>To identify a new Lead MD,</b> change
                                      the Lead MD name at the top of this form
                                      and click either Save Draft or Submit.
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
                                          disabled={
                                            this.state.DisableSaveButton
                                          }
                                        ></PrimaryButton>
                                        <PrimaryButton
                                          text="SUBMIT TO LEAD MD FOR APPROVAL"
                                          className={
                                            !this.state.DisableSubmitButton
                                              ? styles.btnApprovedForReviewerGreen
                                              : ""
                                          }
                                          onClick={this.onSubmit}
                                          disabled={
                                            this.state.DisableSubmitButton
                                          }
                                        ></PrimaryButton>
                                      </Stack>
                                    </div>
                                  </div>
                                  <div className={styles.Spacer}>&nbsp;</div>
                                  <div className={styles.row}>
                                    <div className={styles.col75left}>
                                      <label>
                                        Optional Reversion Comment (visible)
                                      </label>
                                      <TextField
                                        resizable={false}
                                        multiline={false}
                                        value={
                                          this.state.ProjectDetails
                                            .ReviewerReversionComments
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
                                      <div className={styles.Spacer}>
                                        &nbsp;
                                      </div>
                                      <PrimaryButton
                                        text="REVERT TO REVIEWEE"
                                        onClick={this.onRevert}
                                        disabled={
                                          this.state.DisableRevertButton
                                        }
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
                                        <b>LEAD MD:</b> Review the form. Add any
                                        optional comments in the text area
                                        above. When you are satisfied, click the
                                        Submit button below. Alternately, you
                                        could choose to revert to the Reviewer
                                        for more changes. Complete the gray
                                        section below.
                                        <br />
                                        <b>
                                          To substitute a different Lead MD in
                                          this review,
                                        </b>{" "}
                                        enter the new name at the top of the
                                        form and click <b>Replace Me</b>. Your
                                        current inputs will be saved, and the
                                        review will be assigned to the new
                                        person.
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
                                            className={
                                              !this.state.DisableSubmitButton
                                                ? styles.btnApprovedForReviewerGreen
                                                : ""
                                            }
                                            onClick={this.onSubmit}
                                            disabled={
                                              this.state.DisableSubmitButton
                                            }
                                          ></PrimaryButton>
                                        </Stack>
                                      </div>
                                    </div>
                                    <div className={styles.Spacer}>&nbsp;</div>
                                    <div className={styles.row}>
                                      <div className={styles.col75left}>
                                        <label>
                                          Optional Reversion Comment (visible)
                                        </label>
                                        <TextField
                                          resizable={false}
                                          multiline={false}
                                          value={
                                            this.state.ProjectDetails
                                              .LeadMDReversionComments
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
                                        <div className={styles.Spacer}>
                                          &nbsp;
                                        </div>
                                        <PrimaryButton
                                          text="REVERT TO REVIEWER"
                                          onClick={this.onRevert}
                                          disabled={
                                            this.state.DisableRevertButton
                                          }
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
                                        <b>REVIEWEE ACKNOWLEDGEMENT COMMENTS</b>{" "}
                                        (Comments are optional and visible.)
                                      </div>
                                    </div>
                                    <div className={styles.row}>
                                      <div className={styles.col100}>
                                        <TextField
                                          resizable={false}
                                          multiline={false}
                                          disabled={
                                            this.state.ProjectDetails
                                              .StatusOfReview ==
                                            Config.Strings.Status_Acknowledged
                                          }
                                          value={
                                            this.state.ProjectDetails
                                              .AcknowledgementComments
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
                                              disabled={
                                                this.state.DisableSaveButton
                                              }
                                            ></PrimaryButton>
                                            <PrimaryButton
                                              text="SUBMIT FINAL REVIEW"
                                              onClick={this.onSubmit}
                                              disabled={
                                                this.state.DisableSubmitButton
                                              }
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
                                    value={
                                      this.state.ProjectDetails.SignoffHistory
                                    }
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
                      )}
                      {(this.state.ProjectDetails.ServiceLine ==
                        "Data Intelligence Gateway" ||
                        this.state.ProjectDetails.ServiceLine ==
                          "data intelligence gateway") && (
                        <div>
                          <DIGForm
                            ItemID={this.props.ItemID}
                            AppContext={this.props.AppContext}
                            hasEditItemPermission={this.hasEditItemPermission}
                            IsLoading={this.state.IsLoading}
                            CurrentUserRoles={this.state.CurrentUserRoles}
                            ProjectDetails={this.state.ProjectDetails}
                            DisableSaveButton={this.state.DisableSaveButton}
                            DisableSubmitButton={this.state.DisableSubmitButton}
                            DisableRevertButton={this.state.DisableRevertButton}
                            OnlyEnableForReviewer={
                              this.state.OnlyEnableForReviewer
                            }
                            OnlyVisibleForReviewer={
                              this.state.OnlyVisibleForReviewer
                            }
                          ></DIGForm>
                        </div>
                      )}
                    </div>
                  )
                }
              </div>
            </>
          ) : (
            <Spinner size={SpinnerSize.large} />
          )}
        </div>
      </div>
    );
  }
}
