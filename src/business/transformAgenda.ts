import { ISpAgendaTopic, ISpPresenter, ISpEventMaterial } from "../interface/spmodal";
import { McsUtil } from "../utility/helper";
import { sortBy, cloneDeep } from '@microsoft/sp-lodash-subset';
import { business } from ".";


export interface IComponentAgenda extends ISpAgendaTopic {
    SubTopics?: IComponentAgenda[];
    Presenters?: ISpPresenter[];
    Documents?: ISpEventMaterial[];
}

let allAgenda: IComponentAgenda[] = [];
let eventDocuments: ISpEventMaterial[] = [];

export const tranformAgenda = (agendaList: ISpAgendaTopic[], documents: ISpEventMaterial[], presenters: ISpPresenter[]): IComponentAgenda[] => {
    allAgenda = [];
    if (McsUtil.isArray(agendaList)) {
        // make duplicate agenda list
        const tempTopics: IComponentAgenda[] = cloneDeep(agendaList);
        const allDocuments: ISpEventMaterial[] = cloneDeep(documents);

        //assign documents and presenters to agenda items
        tempTopics.forEach(a => {
            a.SubTopics = [];
            a.Presenters = [];
            a.Documents = [];
            const fldname = business.get_AgendaDocumentLookupField();
            if (McsUtil.isArray(a[fldname]) ) {
                (a[fldname] as number[]).forEach((id) => {
                    const index = McsUtil.binarySearch(allDocuments, id, 'Id');
                    if (index >= 0) {
                        a.Documents.push(allDocuments[index]);
                        allDocuments.splice(index, 1);
                    }
                });
            }
            if (McsUtil.isArray(a.PresentersLookupId) ) {
                (a.PresentersLookupId as number[]).forEach((id) => {
                    const index = McsUtil.binarySearch(presenters, id, 'Id');
                    if (index >= 0) {
                        a.Presenters.push(presenters[index]);
                    }
                });
            }
            if (a.Documents.length > 0) {
                a.Documents = sortBy(a.Documents, d => d.SortNumber);
            }
            if (a.Presenters.length > 0) {
                a.Presenters = sortBy(a.Presenters, d => d.SortNumber);
            }
        });

        // all remaining documents are event documents
        eventDocuments = allDocuments;

        // find all subtopics
        const allSubTopics: IComponentAgenda[] = tempTopics.filter(a => McsUtil.isNumeric(a.ParentTopicId) && a.ParentTopicId > 0);
        // find all main topics
        const allTopcis: IComponentAgenda[] = tempTopics.filter(a => !McsUtil.isNumeric(a.ParentTopicId));

        //assign all subtopics to its main topic
        allSubTopics.forEach(a => {
            const index = McsUtil.binarySearch(allTopcis, a.ParentTopicId, 'Id');
            if (index >= 0) {
                allTopcis[index].SubTopics.push(a);
            }
        });
        // sort all subtopics by agenda number
        allTopcis.forEach(a => {
            if (a.SubTopics.length > 0) {
                a.SubTopics = sortBy(a.SubTopics, (b) => b.AgendaNumber);
            }
        });
        // sort all main topics by agenda number
        allAgenda = sortBy(allTopcis, a => a.AgendaNumber);
    } 
    return allAgenda;
};

export const get_tranformAgenda = (): IComponentAgenda[] => {
    return allAgenda;
};

export const get_eventDocuments = (): ISpEventMaterial[] => {
    return eventDocuments;
};