// This namespace contains solution specific Configuration items
export namespace Config {

    export const Search_RowLimit = 250;
    export const List_ThresholdLimit = 5000;

    // Date Formats
    export const DateFormatMoment = "";

    export const ListNames = {
        Projects: "Projects",
        Mentor: "Mentor"
    };

    export const ListCAMLFields = {
    };

    // List sharepoint generated columns with internal name
    export const BaseColumns = {
        Id: "Id",
        Title: "Title",
        ModifedBy: "Editor",
        ModifiedOn: "Modified"
    };

    // Key Value pair of Feedbacks list column title and internal names
    export const ProjectsListColumns = {
      "Acknowledgement Comments" : "AcknowledgementComments",
      "Client Name" : "Customer_x0020_Name",
      "Complexity" : "Complexity",
      "Date Originated" : "Date_x0020_Originated",
      "Date Review Completed" : "Date_x0020_Review_x0020_Complete",
      "Development Areas" : "Development_x0020_Areas",
      "Fiscal Year" : "Fiscal_x0020_Year",
      "Home Office" : "Home_x0020_Office",
      "Hours Worked" : "Hours_x0020_Worked",
      "JobTitle" : "JobTitle1",
      "Last Hours Billed" : "Last_x0020_Hours_x0020_Billed",
      "Lead MD Comments" : "Lead_x0020_MD_x0020_Comments",
      "Lead MD Name" : "Lead_x0020_MD_x0020_Name",
      "Lead MD Name ID" : "Lead_x0020_MD_x0020_NameId",
      "Lead MD Reversion Comments" : "Lead_x0020_MD_x0020_Reversion_x0",
      "Mentor Name" : "Mentor_x0020_Name",
      "Mentor Name ID" : "Mentor_x0020_NameId",
      "Needed Skills" : "Needed_x0020_Skills",
      "Perm Reset" : "Perm_x0020_Reset",
      "Project Code" : "Project_x0020_Code",
      "Project End Date" : "Project_x0020_End_x0020_Date",
      "Project Manager" : "Project_x0020_Manager",
      "Project Manager ID" : "Project_x0020_ManagerId",
      "Project Name" : "Title",
      "Project Start Date" : "Project_x0020_Start_x0020_Date",
      "Project Status" : "Project_x0020_Status",
      "Reviewee Name" : "Reviewee_x0020_Name",
      "Reviewee Name ID" : "Reviewee_x0020_NameId",
      "Reviewer Name" : "Reviewer_x0020_Name",
      "Reviewer Name ID" : "Reviewer_x0020_NameId", 
      "Reviewer Reversion Comments" : "Reviewer_x0020_Reversion_x0020_C",
      "Service Line" : "Service_x0020_Line",
      "Signoff History" : "Signoff_x0020_History",
      "Status of Review" : "Status_x0020_of_x0020_Review",
      "Strong Performance" : "Strong_x0020_Performance",
      "Submitted" : "Submitted",
      "SubstituteUser" : "SubstituteUser",
      "SubstituteUser Id" : "SubstituteUserId",
      "Q1" : "OData__x0051_1",
      "Q2" : "OData__x0051_2",
      "Q3" : "OData__x0051_3",
      "Q4" : "OData__x0051_4",
      "Q5" : "OData__x0051_5",
      "Q6" : "OData__x0051_6",
      "Q7" : "OData__x0051_7",
      "Q8" : "OData__x0051_8",
      "Q9" : "OData__x0051_9",
      "Q10" : "OData__x0051_10",
      "Q11" : "OData__x0051_11",
      "Q12" : "OData__x0051_12",
      "Q13" : "OData__x0051_13",
      "Q1Text" : "Q1Text",
      "Q2Text" : "Q2Text",
      "Q3Text" : "Q3Text",
      "Q4Text" : "Q4Text",
      "Q5Text" : "Q5Text",
      "Q6Text" : "Q6Text",
      "Q7Text" : "Q7Text",
      "Q8Text" : "Q8Text",
      "Q9Text" : "Q9Text",
      "Q10Text" : "Q10Text",
      "Q11Text" : "Q11Text",
      "Q12Text" : "Q12Text",
      "Q13Text" : "Q13Text"
    };

    // Key Value pair of Mentor list column title and internal names
    export const MentorListColumns = {
        "RevieweeName" : "Reviewee_x0020_Name",
        "RevieweeName Id" : "Reviewee_x0020_NameId",
        "Mentor": "Mentor_x0020_Name",
        "Mentor Id": "Mentor_x0020_NameId"
    };

    export const Links = {
        ProjectsListAllItems: "/Lists/Projects/AllItems.aspx",
        HomePageLink: "/",
        RevieweeLink: "/SitePages/Reviewee.aspx",
        ReviewerLink: "/SitePages/Reviewer.aspx",
        LeadMDLink: "/SitePages/LeadMD.aspx",
    };

    export const Strings = {
        NotApplicable: "NA",
        Status_NotStarted: "",
        Status_AwaitingReviewee: "Awaiting Reviewee",
        Status_AwaitingReviewer: "Awaiting Reviewer",
        Status_AwaitingLeadMD: "Awaiting Lead MD",
        Status_AwaitingAcknowledgement: "Awaiting Acknowledgement",
        Status_Acknowledged : "Acknowledged",
        Status_Declined: "Declined",
        Status_SoftDeleted: "Soft Deleted",
        Status_Split: "Split",
        Status_Combined: "Combined"
    };

}