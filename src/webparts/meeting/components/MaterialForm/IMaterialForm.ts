import { ISpEventMaterial } from "../../../../interface/spmodal";
import { IComponentAgenda } from "../../../../business/transformAgenda";

export enum DocumentUploadType {
    InterimDocument = 1,
    LSOBill,
    SessionDocuments
}

export interface IMaterialFormProp {
    document: ISpEventMaterial;
    agenda: IComponentAgenda[];
    requireAgendaSelection?: boolean;
}

export interface IWorkingDocument {
    title: string;
    agency: string;
    lsonumber: string;
    billVersion: string;
    sessionDocumentId: number;
    uploadFile?: FileList;
    includeWithAgenda: boolean;
    selectedAgency: any;
}

export interface IMaterialFormState {
    agenda: IComponentAgenda;
    selectedSubTopic: IComponentAgenda;
    documentUploadType: DocumentUploadType;
    workingDocument: IWorkingDocument;
    documentId: number;
}