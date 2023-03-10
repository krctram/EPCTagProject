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
  public DevelopmentAreas: string;
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

  public Q14: string;
  public Q14Text: string;
  public Q15: string;
  public Q15Text: string;
  public Q16: string;
  public Q16Text: string;
  public Q17: string;
  public Q17Text: string;
  public Q18: string;
  public Q18Text: string;

  public Q1Category?: string;
  public Q2Category?: string;
  public Q3Category?: string;
  public Q4Category?: string;
  public Q5Category?: string;
  public Q6Category?: string;
  public Q7Category?: string;
  public Q8Category?: string;
  public Q9Category?: string;
  public Q10Category?: string;
  public Q11Category?: string;
  public Q12Category?: string;
  public Q13Category?: string;
  public Q14Category?: string;
  public Q15Category?: string;
  public Q16Category?: string;
  public Q17Category?: string;
  public Q18Category?: string;

  public Q1IsRating?: boolean;
  public Q2IsRating?: boolean;
  public Q3IsRating?: boolean;
  public Q4IsRating?: boolean;
  public Q5IsRating?: boolean;
  public Q6IsRating?: boolean;
  public Q7IsRating?: boolean;
  public Q8IsRating?: boolean;
  public Q9IsRating?: boolean;
  public Q10IsRating?: boolean;
  public Q11IsRating?: boolean;
  public Q12IsRating?: boolean;
  public Q13IsRating?: boolean;
  public Q14IsRating?: boolean;
  public Q15IsRating?: boolean;
  public Q16IsRating?: boolean;
  public Q17IsRating?: boolean;
  public Q18IsRating?: boolean;
}
