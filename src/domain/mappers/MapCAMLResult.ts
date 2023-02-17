import { ContextService } from "../../services/ContextService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as moment from 'moment';
import { User } from "../models/types/User";
import { Enums } from "../../globals/Enums";
import { Config } from "../../globals/Config";
import { TAG_ProjectDetails } from "../models/TAG_ProjectDetails";

export default class MapCAMLResult extends ContextService {

    constructor(AppContext: WebPartContext, Lcid: number) {
        super(AppContext);
    }

    // Mapping results based on provided type
    public static map(items: any, type: Enums.ItemResultType): any[] {
        let allResults: any[] = [];
        items.forEach(item => {
            let result: any;
            switch (type) {
                case Enums.ItemResultType.TAG_ProjectDetails: result = this.mapProjectDetails(item);
                    break;
            }
            allResults.push(result);
        });
        return allResults;
    }

    //#region "Solution Related Mappers"

    private static mapProjectDetails(item: any) {
        const Columns = Config.ProjectsListColumns;
        let result = new TAG_ProjectDetails();
        result.AcknowledgementComments = item[Columns["Acknowledgement Comments"]];
        result.ClientName = item[Columns["Client Name"]];
        result.Complexity = item[Columns.Complexity];
        result.DateOriginated = this.mapDate(item[Columns["Date Originated"]]);
        result.DateOriginatedFormatted = this.mapDateWithFormat(item[Columns["Date Originated"]]);
        result.DateReviewCompleted = this.mapDate(item[Columns["Date Review Completed"]]);
        result.DateReviewCompletedFormatted = this.mapDateWithFormat(item[Columns["Date Review Completed"]]); 
        result.DevelopmentAreas = item[Columns["Development Areas"]];
        result.FiscalYear = item[Columns["Fiscal Year"]];
        result.HomeOffice = item[Columns["Home Office"]];
        result.HoursWorked = item[Columns["Hours Worked"]];
        result.ID = item[Config.BaseColumns.Id];
        result.JobTitle = item[Columns.JobTitle];
        result.LastHoursBilled = this.mapDate(item[Columns["Last Hours Billed"]]);
        result.LastHoursBilledFormatted = this.mapDateWithFormat(item[Columns["Last Hours Billed"]]);
        result.LeadMD = this.mapUser(item[Columns["Lead MD Name"]]);
        result.LeadMDComments = item[Columns["Lead MD Comments"]];
        result.LeadMDReversionComments = item[Columns["Lead MD Reversion Comments"]];
        result.Mentor = this.mapUser(item[Columns["Mentor Name"]]);
        result.NeededSkills = item[Columns["Needed Skills"]];
        result.PermReset = item[Columns["Perm Reset"]];
        result.ProjectCode = item[Columns["Project Code"]];
        result.ProjectEndDate = this.mapDate(item[Columns["Project End Date"]]);
        result.ProjectEndDateFormatted = this.mapDateWithFormat(item[Columns["Project End Date"]]);
        result.ProjectManager = this.mapUser(item[Columns["Project Manager"]]);
        result.ProjectName = item[Columns["Project Name"]];
        result.ProjectStartDate = this.mapDate(item[Columns["Project Start Date"]]);
        result.ProjectStartDateFormatted = this.mapDateWithFormat(item[Columns["Project Start Date"]]);
        result.ProjectStatus = item[Columns["Project Status"]];
        result.Reviewee = this.mapUser(item[Columns["Reviewee Name"]]);
        result.Reviewer = this.mapUser(item[Columns["Reviewer Name"]]);
        result.ReviewerReversionComments = item[Columns["Reviewer Reversion Comments"]];
        result.ServiceLine = item[Columns["Service Line"]];
        result.SignoffHistory = item[Columns["Signoff History"]];
        result.StatusOfReview = item[Columns["Status of Review"]];
        result.StrongPerformance = item[Columns["Strong Performance"]];
        result.Submitted = item[Columns.Submitted];
        result.SubstituteUser = this.mapUser(item[Columns.SubstituteUser]);

        result.Q1 = item[Columns.Q1.replace("OData_","")];
        result.Q2 = item[Columns.Q2.replace("OData_","")];
        result.Q3 = item[Columns.Q3.replace("OData_","")];
        result.Q4 = item[Columns.Q4.replace("OData_","")];
        result.Q5 = item[Columns.Q5.replace("OData_","")];
        result.Q6 = item[Columns.Q6.replace("OData_","")];
        result.Q7 = item[Columns.Q7.replace("OData_","")];
        result.Q8 = item[Columns.Q8.replace("OData_","")];
        result.Q9 = item[Columns.Q9.replace("OData_","")];
        result.Q10 = item[Columns.Q10.replace("OData_","")];
        result.Q11 = item[Columns.Q11.replace("OData_","")];
        result.Q12 = item[Columns.Q12.replace("OData_","")];
        result.Q13 = item[Columns.Q13.replace("OData_","")];
        
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
        
        result.ModifiedBy = this.mapUser(item[Config.BaseColumns.ModifedBy]);
        result.ModifiedOnFormatted = this.mapDateWithFormat(item[Config.BaseColumns.ModifiedOn]);
        return result;
    }

    ////#endregion

    //#region "Common Mappers"

    // Mapping multiple user
    private static mapUsers(userEntries: any): User[] {
        let result: User[] = [];
        if (userEntries instanceof Array) {
            userEntries.forEach(user => {
                result.push(this.mapUser(user));
            });
        }
        else {
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
            result = (itemValue == "Yes" || itemValue.value == "1") ? true : false;
            return result;
        }
        return undefined;
    }

    // Mapping date field
    private static mapDate(dateField: any): Date {
        if (dateField) {
            return (new Date(dateField));
        }
        return undefined;
    }

    // Mapping date field and return formatted date string
    private static mapDateWithFormat(dateField: any): string {
        if (dateField) {
            return (moment(dateField).format('M/DD/YYYY'));
        }
        return "";
    }

    //#endregion
}



