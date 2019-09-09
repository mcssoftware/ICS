export interface IUser {
    Id: number;
    Title: string;
    EMail: string;
    JobTitle: string;
    Department: string;
}

export interface IFileVersion {
    CheckInComment: string;
    Created: string;
    ID: number;
    IsCurrentVersion: boolean;
    Url: string;
    VersionLabel: string;
}

export interface IFile {
    CheckInComment: string;
    CheckOutType: number;
    ETag: string;
    Exists: boolean;
    IrmEnabled: boolean;
    Length: string;
    Level: number;
    LinkingUrl: string;
    MajorVersion: number;
    MinorVersion: number;
    Name: string;
    ServerRelativeUrl: string;
    Title: string;
    UIVersion: number;
    UIVersionLabel: string;
    UniqueId: string;
}
export interface IContentTypeId {
    StringValue: string;
}
export interface IContentType {
    Id: IContentTypeId;
    StringId: string;
    Name: string;
    DisplayFormUrl: string;
    DocumentTemplate: string;
    DocumentTemplateUrl: string;
}
export interface IListItem {
    Id?: number;
    Title: string;
    AuthorId?: number;
    ContentType?: IContentType;
    ContentTypeId?: string;
    Created?: string | Date;
    Modified?: string | Date;
    Editor?: IUser;
    "odata.etag"?: string;
    "odata.type"?: string;
}
export interface IItemVersion extends IListItem {
    VersionId: number;
    VersionLabel: string;
}
export interface IDocumentItem extends IListItem {
    File?: IFile;
}
export interface IMultipleLookupField {
    __metadata: any;
    results: number[];
}
export interface ISpEvent extends IListItem {
    EventDate: string;
    MeetingStartTime: string;
    EndDate: string;
    Location: string;
    Description: string;
    fAllDayEvent?: boolean;
    fRecurrence?: boolean;
    ParticipantsPickerId?: any;
    Category: any;
    WorkAddress: string;
    WorkCity: string;
    WorkState: string;
    OtherLocationInfo: string;
    CommitteeStaff: string;
    ConferenceNumber: string;
    ApprovedStatus: string;
    EventDocumentsLookupId?: IMultipleLookupField;
    JointEventCommitteeId?: string;
    CommitteeLookupId?: number;
    CommitteeEventLookupId?: number;
    HasLiveStream: boolean;
    IsBudgetHearing: boolean;
}
export interface ISpAgendaTopic extends IListItem {
    AgendaTitle: string;
    AgendaNumber: number;
    AgendaDate: Date | string;
    EventLookupId: number;
    PresentersLookupId?: IMultipleLookupField;
    ParentTopicId: number;
    AgendaDocumentsLookupId?: IMultipleLookupField;
    Presenters?: Array<ISpPresenter>;
    AllowPublicComments: boolean;
}

export interface ISpPresenter extends IListItem {
    PresenterName: string;
    OrganizationName: string;
    SortNumber: number;
}

export interface ISpEventMaterial extends IDocumentItem {
    //AgendaLookupId: number;
    //SubTopicId?: number;
    //PresenterLookupId?: number;
    AgencyName: string;
    lsoDocumentType: string;
    IncludeWithAgenda: boolean;
    SortNumber: number;
}

export interface IURL {
    Description: string;
    Url: string;
}

export interface ISpCommitteeLink extends IListItem {
    URL: IURL;
    CommitteeId: string;
    CommitteeName: string;
    DisplayOrder: string;
    CommitteeDesktopUrl: string;
}

export enum OperationType {
    Add = 1,
    Edit,
    Delete
}