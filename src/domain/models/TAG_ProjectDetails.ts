
import { User } from "./types/User";

export class TAG_ProjectDetails {
    public ID: number;
    public AcknowledgementComments: string;
    public ClientName: string;
    public Complexity: string;
    public DateOriginated: Date;
    public DateOriginatedFormatted: string;
    public DateReviewCompleted: Date;
    public DateReviewCompletedFormatted: string;
    public DevelopmentAreas : string;
    public FiscalYear: string;
    public HomeOffice: string;
    public HoursWorked: number;
    public JobTitle: string;
    public LastHoursBilled: Date;
    public LastHoursBilledFormatted: string;
    public LeadMDComments: string;
    public LeadMD: User;
    public LeadMDReversionComments: string;
    public Mentor: User;
    public NeededSkills: string;
    public PermReset: string; // Choice in Back-end
    public ProjectCode: string;
    public ProjectEndDate: Date;
    public ProjectEndDateFormatted: string;
    public ProjectManager: User;
    public ProjectName: string;
    public ProjectStartDate: Date;
    public ProjectStartDateFormatted: string;
    public ProjectStatus: string;
    public Reviewee: User;
    public Reviewer: User;
    public ReviewerReversionComments: string;
    public ServiceLine: string;
    public SignoffHistory: string;
    public StatusOfReview: string;
    public StrongPerformance: string;
    public Submitted: number;
    public SubstituteUser: User;
    public Q1: string;
    public Q1Text: string;
    public Q2: string;
    public Q2Text: string;
    public Q3: string;
    public Q3Text: string;
    public Q4: string;
    public Q4Text: string;
    public Q5: string;
    public Q5Text: string;
    public Q6: string;
    public Q6Text: string;
    public Q7: string;
    public Q7Text: string;
    public Q8: string;
    public Q8Text: string;
    public Q9: string;
    public Q9Text: string;
    public Q10: string;
    public Q10Text: string;
    public Q11: string;
    public Q11Text: string;
    public Q12: string;
    public Q12Text: string;
    public Q13: string;
    public Q13Text: string;
    public ModifiedBy: User;
    public ModifiedOn: Date;
    public ModifiedOnFormatted: string;
}