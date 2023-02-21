export namespace Enums {
  export enum FieldTypes {
    TaxonomyMulti = "TaxonomyFieldTypeMulti",
    TaxonomySingle = "TaxonomyFieldType",
    PersonMulti = "UserMulti",
    PersonSingle = "User",
    Link = "URL",
    Lookup = "",
    LookupMulti = "",
  }

  export enum MapperType {
    PNPResult,
    PnPControlResult,
    CAMLResult,
    SearchResult,
    None,
  }

  export enum ItemResultType {
    //Common Result Types
    None,
    User,
    UserProfile,
    Users,
    Document,
    Item,
    Task,

    //Solution Specific Result Types
    TAG_ProjectDetails,
    TAG_QuestionText,
  }

  export enum DataPayloadTypes {
    PnPCreateUpdate,
    PnPValidateUpdate,
  }

  export enum UserRoles {
    Reviewee,
    Reviewer,
    SuperAdmin,
  }

  export enum SaveType {
    Submit,
    SaveAsDraft,
    Revert,
    StartReview,
    ReplaceMe,
  }
}
