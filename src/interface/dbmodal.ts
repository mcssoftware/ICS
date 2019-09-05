import { ISpEventMaterial } from "./spmodal";

export interface IDbCommittee {
    FullName: string;
    ShortName?: string;
    Year: number;
    Code: string;
    DisplayOrder?: number;
    CommitteeMembers?: Array<IDbMembers>;
    Staff?: Array<IDbStaff>;
}
export interface IDbMembers {
    shortDesc: string;
    description?: any;
    legislativeYear: number;
    committeeId: string;
    firstName: string;
    lastName: string;
    isChairman: boolean;
    isViceChair: boolean;
    chamber: string;
    notes: string;
    id: number;
    title: string;
    legislatureName: string;
    legislatureDisplayName: string;
    party: string;
    legislatorID: number;
    county: string;
    chamberValue: string;
    email: string;
    addressLine1: string;
    addressLine2?: any;
    city: string;
    stateProvince: string;
    zipPostalCode: string;
    legislatureLoginId: string;
    canSponsor: boolean;
}
export interface IDbStaff {
    eMail: string;
    jobTitle: string;
    name: string;
}

export interface IDbMeeting {
    Office365ID: number;
    CommitteeYear: number;
    MeetingType: string;
    Purpose: string;
    StartDate: Date;
    EndTime: Date;
    ConferenceNumber: string;
    Address1: string;
    Address2: string;
    Address3: string;
    City: string;
    State: string;
    CommitteeStaff: string;
    IsPublic: boolean;
    LastUpdated?: Date;
    CommitteeList: IDbCommittee[];
    MeetingAgendas?: Array<IDbAgenda>;
    MeetingDocuments?: Array<IDbDocument>;
    HasLiveStream: boolean;
    IsBudgetHearing: boolean;
    Chairmen: string;
}

export interface IDbAgenda {
    Office365ID: number;
    Title: string;
    SortNumber: number;
    AgendaDate?: Date;
    MeetingId: number;
    ParentAgenda365Id?: number;
    AllowPublicComments: boolean;
    SubAgendaItems?: Array<IDbAgenda>;
    MeetingDocuments?: Array<IDbDocument>;
    MeetingPresenters?: Array<IDbPresenter>;
}

export interface IDbPresenter {
    Office365ID: number;
    Title: string;
    Name: string;
    Organization?: string;
    SortNumber: number;
}

export interface IDbDocument {
    Office365ID: number;
    UniqueId: string;
    FileName: string;
    Title: string;
    DocumentUrl: string;
    VersionId: string;
    DocumentType: string;
    AgencyName: string;
    AgendaTitle?: string;
    IncludeWithAgenda: boolean;
    SortNumber: number;
    Item: ISpEventMaterial;
}