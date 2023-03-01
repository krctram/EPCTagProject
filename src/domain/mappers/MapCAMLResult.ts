import { ContextService } from "../../services/ContextService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as moment from "moment";
import { User } from "../models/types/User";
import { Enums } from "../../globals/Enums";
import { Config } from "../../globals/Config";
import { TAG_ProjectDetails } from "../models/TAG_ProjectDetails";
import { TAG_QuestionText } from "../models/TAG_QuestionText";

export default class MapCAMLResult extends ContextService {
  constructor(AppContext: WebPartContext, Lcid: number) {
    super(AppContext);
  }

  // Mapping results based on provided type
  public static map(items: any, type: Enums.ItemResultType): any[] {
    let allResults: any[] = [];
    items.forEach((item) => {
      let result: any;
      switch (type) {
        case Enums.ItemResultType.TAG_ProjectDetails:
          result = this.mapProjectDetails(item);
          break;
        case Enums.ItemResultType.TAG_QuestionText:
          result = this.mapQuestionText(item);
          break;
      }
      allResults.push(result);
    });
    return allResults;
  }

  //#region "Solution Related Mappers"

  private static mapQuestionText(item: any) {
    const Columns = Config.QuestionListColumns;
    let result = new TAG_QuestionText();

    result.Title = item[Columns.Title];
    result.ServiceLine = item[Columns.ServiceLine];
    result.Q1 = item[Columns.Q1.replace("OData_", "")];
    result.Q2 = item[Columns.Q2.replace("OData_", "")];
    result.Q3 = item[Columns.Q3.replace("OData_", "")];
    result.Q4 = item[Columns.Q4.replace("OData_", "")];
    result.Q5 = item[Columns.Q5.replace("OData_", "")];
    result.Q6 = item[Columns.Q6.replace("OData_", "")];
    result.Q7 = item[Columns.Q7.replace("OData_", "")];
    result.Q8 = item[Columns.Q8.replace("OData_", "")];
    result.Q9 = item[Columns.Q9.replace("OData_", "")];
    result.Q10 = item[Columns.Q10.replace("OData_", "")];
    result.Q11 = item[Columns.Q11.replace("OData_", "")];
    result.Q12 = item[Columns.Q12.replace("OData_", "")];
    result.Q13 = item[Columns.Q13.replace("OData_", "")];
    result.Q14 = item[Columns.Q4.replace("OData_", "")];
    result.Q15 = item[Columns.Q5.replace("OData_", "")];
    result.Q16 = item[Columns.Q6.replace("OData_", "")];
    result.Q17 = item[Columns.Q7.replace("OData_", "")];
    result.Q18 = item[Columns.Q8.replace("OData_", "")];

    result.Q1Category = item[Columns.Q1Category];
    result.Q2Category = item[Columns.Q2Category];
    result.Q3Category = item[Columns.Q3Category];
    result.Q4Category = item[Columns.Q4Category];
    result.Q5Category = item[Columns.Q5Category];
    result.Q6Category = item[Columns.Q6Category];
    result.Q7Category = item[Columns.Q7Category];
    result.Q8Category = item[Columns.Q8Category];
    result.Q9Category = item[Columns.Q9Category];
    result.Q10Category = item[Columns.Q10Category];
    result.Q11Category = item[Columns.Q11Category];
    result.Q12Category = item[Columns.Q12Category];
    result.Q13Category = item[Columns.Q13Category];
    result.Q14Category = item[Columns.Q14Category];
    result.Q15Category = item[Columns.Q15Category];
    result.Q16Category = item[Columns.Q16Category];
    result.Q17Category = item[Columns.Q17Category];
    result.Q18Category = item[Columns.Q18Category];

    result.Q1IsRating = item[Columns.Q1IsRating];
    result.Q2IsRating = item[Columns.Q2IsRating];
    result.Q3IsRating = item[Columns.Q3IsRating];
    result.Q4IsRating = item[Columns.Q4IsRating];
    result.Q5IsRating = item[Columns.Q5IsRating];
    result.Q6IsRating = item[Columns.Q6IsRating];
    result.Q7IsRating = item[Columns.Q7IsRating];
    result.Q8IsRating = item[Columns.Q8IsRating];
    result.Q9IsRating = item[Columns.Q9IsRating];
    result.Q10IsRating = item[Columns.Q10IsRating];
    result.Q11IsRating = item[Columns.Q11IsRating];
    result.Q12IsRating = item[Columns.Q12IsRating];
    result.Q13IsRating = item[Columns.Q13IsRating];
    result.Q14IsRating = item[Columns.Q14IsRating];
    result.Q15IsRating = item[Columns.Q15IsRating];
    result.Q16IsRating = item[Columns.Q16IsRating];
    result.Q17IsRating = item[Columns.Q17IsRating];
    result.Q18IsRating = item[Columns.Q18IsRating];

    return result;
  }

  private static mapProjectDetails(item: any) {
    const Columns = Config.ProjectsListColumns;
    let result = new TAG_ProjectDetails();
    result.AcknowledgementComments = item[Columns["Acknowledgement Comments"]];
    result.ClientName = item[Columns["Client Name"]];
    result.Complexity = item[Columns.Complexity];
    result.DateOriginated = this.mapDate(item[Columns["Date Originated"]]);
    result.DateOriginatedFormatted = this.mapDateWithFormat(
      item[Columns["Date Originated"]]
    );
    result.DateReviewCompleted = this.mapDate(
      item[Columns["Date Review Completed"]]
    );
    result.DateReviewCompletedFormatted = this.mapDateWithFormat(
      item[Columns["Date Review Completed"]]
    );
    result.DevelopmentAreas = item[Columns["Development Areas"]];
    result.FiscalYear = item[Columns["Fiscal Year"]];
    result.HomeOffice = item[Columns["Home Office"]];
    result.HoursWorked = item[Columns["Hours Worked"]];
    result.ID = item[Config.BaseColumns.Id];
    result.JobTitle = item[Columns.JobTitle];
    result.LastHoursBilled = this.mapDate(item[Columns["Last Hours Billed"]]);
    result.LastHoursBilledFormatted = this.mapDateWithFormat(
      item[Columns["Last Hours Billed"]]
    );
    result.LeadMD = this.mapUser(item[Columns["Lead MD Name"]]);
    result.LeadMDComments = item[Columns["Lead MD Comments"]];
    result.LeadMDReversionComments =
      item[Columns["Lead MD Reversion Comments"]];
    result.Mentor = this.mapUser(item[Columns["Mentor Name"]]);
    result.NeededSkills = item[Columns["Needed Skills"]];
    result.PermReset = item[Columns["Perm Reset"]];
    result.ProjectCode = item[Columns["Project Code"]];
    result.ProjectEndDate = this.mapDate(item[Columns["Project End Date"]]);
    result.ProjectEndDateFormatted = this.mapDateWithFormat(
      item[Columns["Project End Date"]]
    );
    result.ProjectManager = this.mapUser(item[Columns["Project Manager"]]);
    result.ProjectName = item[Columns["Project Name"]];
    result.ProjectStartDate = this.mapDate(item[Columns["Project Start Date"]]);
    result.ProjectStartDateFormatted = this.mapDateWithFormat(
      item[Columns["Project Start Date"]]
    );
    result.ProjectStatus = item[Columns["Project Status"]];
    result.Reviewee = this.mapUser(item[Columns["Reviewee Name"]]);
    result.Reviewer = this.mapUser(item[Columns["Reviewer Name"]]);
    result.ReviewerReversionComments =
      item[Columns["Reviewer Reversion Comments"]];
    result.ServiceLine = item[Columns["Service Line"]];

    let SignoffHistory = item[Columns["Signoff History"]]
      ? item[Columns["Signoff History"]].split(";")
      : "";
    let html = "";
    for (var i = 0; i < Object.keys(SignoffHistory).length; i++) {
      if (SignoffHistory[Object.keys(SignoffHistory)[i]] != " ") {
        html += SignoffHistory[Object.keys(SignoffHistory)[i]].trim() + "\n";
      }
    }
    result.SignoffHistory = html;
    // result.SignoffHistory = item[Columns["Signoff History"]];
    result.StatusOfReview = item[Columns["Status of Review"]];
    result.StrongPerformance = item[Columns["Strong Performance"]];
    result.Submitted = item[Columns.Submitted];
    result.SubstituteUser = this.mapUser(item[Columns.SubstituteUser]);

    result.Q1 = item[Columns.Q1.replace("OData_", "")];
    result.Q2 = item[Columns.Q2.replace("OData_", "")];
    result.Q3 = item[Columns.Q3.replace("OData_", "")];
    result.Q4 = item[Columns.Q4.replace("OData_", "")];
    result.Q5 = item[Columns.Q5.replace("OData_", "")];
    result.Q6 = item[Columns.Q6.replace("OData_", "")];
    result.Q7 = item[Columns.Q7.replace("OData_", "")];
    result.Q8 = item[Columns.Q8.replace("OData_", "")];
    result.Q9 = item[Columns.Q9.replace("OData_", "")];
    result.Q10 = item[Columns.Q10.replace("OData_", "")];
    result.Q11 = item[Columns.Q11.replace("OData_", "")];
    result.Q12 = item[Columns.Q12.replace("OData_", "")];
    result.Q13 = item[Columns.Q13.replace("OData_", "")];
    result.Q14 = item[Columns.Q14.replace("OData_", "")];
    result.Q15 = item[Columns.Q15.replace("OData_", "")];
    result.Q16 = item[Columns.Q16.replace("OData_", "")];
    result.Q17 = item[Columns.Q17.replace("OData_", "")];
    result.Q18 = item[Columns.Q18.replace("OData_", "")];

    result.Q1Text = item[Columns.Q1Text];
    result.Q2Text = item[Columns.Q2Text];
    result.Q3Text = item[Columns.Q3Text];
    result.Q4Text = item[Columns.Q4Text];
    result.Q5Text = item[Columns.Q5Text];
    result.Q6Text = item[Columns.Q6Text];
    result.Q7Text = item[Columns.Q7Text];
    result.Q8Text = item[Columns.Q8Text];
    result.Q9Text = item[Columns.Q9Text];
    result.Q10Text = item[Columns.Q10Text];
    result.Q11Text = item[Columns.Q11Text];
    result.Q12Text = item[Columns.Q12Text];
    result.Q13Text = item[Columns.Q13Text];
    result.Q14Text = item[Columns.Q14Text];
    result.Q15Text = item[Columns.Q15Text];
    result.Q16Text = item[Columns.Q16Text];
    result.Q17Text = item[Columns.Q17Text];
    result.Q18Text = item[Columns.Q18Text];

    result.Q1Category = "";
    result.Q2Category = "";
    result.Q3Category = "";
    result.Q4Category = "";
    result.Q5Category = "";
    result.Q6Category = "";
    result.Q7Category = "";
    result.Q8Category = "";
    result.Q9Category = "";
    result.Q10Category = "";
    result.Q11Category = "";
    result.Q12Category = "";
    result.Q13Category = "";
    result.Q14Category = "";
    result.Q15Category = "";
    result.Q16Category = "";
    result.Q17Category = "";
    result.Q18Category = "";

    result.Q1IsRating = false;
    result.Q2IsRating = false;
    result.Q3IsRating = false;
    result.Q4IsRating = false;
    result.Q5IsRating = false;
    result.Q6IsRating = false;
    result.Q7IsRating = false;
    result.Q8IsRating = false;
    result.Q9IsRating = false;
    result.Q10IsRating = false;
    result.Q11IsRating = false;
    result.Q12IsRating = false;
    result.Q13IsRating = false;
    result.Q14IsRating = false;
    result.Q15IsRating = false;
    result.Q16IsRating = false;
    result.Q17IsRating = false;
    result.Q18IsRating = false;

    result.ModifiedBy = this.mapUser(item[Config.BaseColumns.ModifedBy]);
    result.ModifiedOnFormatted = this.mapDateWithFormat(
      item[Config.BaseColumns.ModifiedOn]
    );
    return result;
  }
  ////#endregion

  //#region "Common Mappers"

  // Mapping multiple user
  private static mapUsers(userEntries: any): User[] {
    let result: User[] = [];
    if (userEntries instanceof Array) {
      userEntries.forEach((user) => {
        result.push(this.mapUser(user));
      });
    } else {
      result.push(this.mapUser(userEntries));
    }

    return result;
  }

  // Mapping single user
  private static mapUser(user: any): User {
    // This in required, as in CAML it returns array even if it is single user
    if (user instanceof Array && user.length > 0) {
      user = user[0];
    }
    // Case : when it is null
    if (!user) {
      return new User();
    }
    let result: User = new User();
    result.Email = user["email"];
    result.Id = user["id"];
    result.LoginName = user["sip"];

    if (result.LoginName.indexOf("i:0#") < 0) {
      result.LoginName = "i:0#.f|membership|" + result.Email;
    }

    result.Title = user["title"];
    return result;
  }

  // Mapping boolean value
  private static mapBoolean(itemValue: any): boolean {
    if (itemValue) {
      let result: boolean;
      result = itemValue == "Yes" || itemValue.value == "1" ? true : false;
      return result;
    }
    return undefined;
  }

  // Mapping date field
  private static mapDate(dateField: any): Date {
    if (dateField) {
      return new Date(dateField);
    }
    return undefined;
  }

  // Mapping date field and return formatted date string
  private static mapDateWithFormat(dateField: any): string {
    if (dateField) {
      return moment(dateField).format("M/DD/YYYY");
    }
    return "";
  }

  //#endregion
}
