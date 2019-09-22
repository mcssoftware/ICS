import { ISpEventMaterial, OperationType } from "../../../../interface/spmodal";
import { IComponentAgenda } from "../../../../business/transformAgenda";

export enum DocumentUploadType {
    InterimDocument = 1,
    LSOBill,
    SessionDocuments
}

export interface IMaterialFormProp {
    meetingId: number;
    document: ISpEventMaterial;
    agenda: IComponentAgenda[];
    requireAgendaSelection?: boolean;
    sortNumber: number;
    onChange: (document: ISpEventMaterial, agenda: IComponentAgenda, type: OperationType) => void;
}

export interface IWorkingDocument {
    sortNumber: string;
    title: string;
    lsonumber: string;
    billVersion: string;
    sessionDocumentId: number;
    uploadFile?: FileList;
    includeWithAgenda: boolean;
    selectedAgency: any;
    selectedBill: any;
    selectedBillVersion: any;
    lsoDocumentType: string;
}

export interface IMaterialFormState {
    agenda: IComponentAgenda;
    selectedSubTopic: IComponentAgenda;
    documentUploadType: DocumentUploadType;
    workingDocument: IWorkingDocument;
    documentId: number;
    waitingMessage?: string;
    loadingBillVersion: boolean;
}