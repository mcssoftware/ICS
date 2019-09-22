import { IDbMeeting, IDbAgenda, IDbDocument, IDbPresenter } from "../interface/dbmodal";
import { ISpEvent, ISpAgendaTopic, ISpEventMaterial, ISpPresenter } from "../interface/spmodal";
import { McsUtil } from "../utility/helper";
import { business } from ".";

const transformEvent = (model: ISpEvent): IDbMeeting => {

    var s1 = McsUtil.convertUtcDateToLocalDate(new Date(model.EventDate));
    var s2 = model.MeetingStartTime;
    const startDate = new Date(s1.toLocaleDateString() + ", " + s2.replace(/[^\d\s:APM]/gi, ""));
    const endDate = McsUtil.convertUtcDateToLocalDate(new Date(model.EndDate));
    const meetingData: IDbMeeting = {
        Office365ID: model.Id,
        CommitteeYear: startDate.getFullYear(),
        MeetingType: model.Category,
        Purpose: model.Description,
        StartDate: startDate,
        EndTime: endDate,
        ConferenceNumber: model.ConferenceNumber,
        Address1: model.Location,
        Address2: model.OtherLocationInfo,
        Address3: model.WorkAddress,
        City: model.WorkCity,
        State: model.WorkState,
        CommitteeStaff: model.CommitteeStaff,
        CommitteeList: null,
        IsPublic: false,
        HasLiveStream: model.HasLiveStream,
        IsBudgetHearing: model.IsBudgetHearing || false,
        Chairmen: ""
    };
    return meetingData;
};

const transformAgenda = (agenda: ISpAgendaTopic, meetingId: number): IDbAgenda => {
    var subagenda: IDbAgenda = {
        Office365ID: agenda.Id,
        Title: agenda.AgendaTitle,
        SortNumber: McsUtil.isDefined(agenda.AgendaNumber) ? parseInt(agenda.AgendaNumber.toString()) : 99,
        MeetingId: meetingId,
        AllowPublicComments: agenda.AllowPublicComments || false,
        MeetingDocuments: [],
        MeetingPresenters: [],
        SubAgendaItems: []
    };
    if (McsUtil.isDefined(agenda.AgendaDate)) {
        subagenda.AgendaDate = new Date(agenda.AgendaDate as any);
    }
    if (agenda.ParentTopicId > 0) {
        subagenda.ParentAgenda365Id = agenda.ParentTopicId;
    }
    return subagenda;
};

const transformDocument = (material: ISpEventMaterial, agenda: IDbAgenda, webAbsoluteUrl: string): IDbDocument => {
    var model: IDbDocument = {
        Office365ID: material.Id,
        FileName: material.File.Name,
        Title: material.Title,
        DocumentUrl: material.File.ServerRelativeUrl,
        VersionId: material.File.UIVersion.toString(),
        DocumentType: material.lsoDocumentType,
        AgencyName: material.AgencyName,
        AgendaTitle: McsUtil.isDefined(agenda) ? agenda.Title : null,
        IncludeWithAgenda: material.IncludeWithAgenda || false,
        Item: material,
        UniqueId: material.File.UniqueId,
        SortNumber: material.SortNumber || 1
    };
    if (material.File.MajorVersion > 0) {
        model.DocumentUrl = McsUtil.getRelativeUrl(McsUtil.combinePaths(webAbsoluteUrl, "/_vti_history/",
            (material.File.MajorVersion * 512).toString(),
            material.File.ServerRelativeUrl));
    }
    return model;
};

const transformPresenter = (presenter: ISpPresenter): IDbPresenter => {
    var model: IDbPresenter = {
        Office365ID: presenter.Id,
        Title: presenter.Title,
        Name: presenter.PresenterName,
        Organization: presenter.OrganizationName,
        SortNumber: presenter.SortNumber || 1
    };
    return model;
};

const tranformPresenterList = (ids: number[], presenterList: ISpPresenter[]): IDbPresenter[] => {
    if (McsUtil.isArray(ids) && ids.length > 0) {
        return ids.map(a => {
            const index = McsUtil.binarySearch(presenterList, a, 'Id');
            if (index >= 0) {
                return transformPresenter(presenterList[index]);
            } else {
                return null;
            }
        }).filter(a => McsUtil.isDefined(a));
    }
    return [];
};


/**
 * Transform Sharepoint list items to DB format for publishing
 *
 * @param {string} webAbsoluteUrl current web url
 * @param {ISpEvent} event
 * @param {ISpAgendaTopic[]} agenda list of agenda sorted by ID
 * @param {ISpEventMaterial[]} material list of documents sorted by ID
 * @param {ISpPresenter[]} presenter list of presenter sorted by ID
 */
const transform = (webAbsoluteUrl: string, event: ISpEvent, agendaList: ISpAgendaTopic[], materialList: ISpEventMaterial[], presenterList: ISpPresenter[]): IDbMeeting => {
    const modal = transformEvent(event);
    const allmaterials: ISpEventMaterial[] = [...materialList];
    let lastModifiedDate = new Date(event.Modified as string);
    // get only TOPIC
    const mainTopics: IDbAgenda[] = agendaList.filter(a => !McsUtil.isNumeric(a.ParentTopicId))
        .map((a) => {
            const agenda = transformAgenda(a, event.Id);
            //update last mofied if agenda was modified
            var agendaModified = new Date(a.Modified as string);
            if (agendaModified > lastModifiedDate) {
                lastModifiedDate = agendaModified;
            }
            agenda.MeetingPresenters = tranformPresenterList(McsUtil.isArray(a.PresentersLookupId as number[]) ? a.PresentersLookupId as number[] : [], presenterList);
            const fldname = business.get_AgendaDocumentLookupField();
            if (McsUtil.isArray(a[fldname])) {
                agenda.MeetingDocuments = (a[fldname] as number[]).map((d) => {
                    const docIndex = McsUtil.binarySearch(allmaterials, d, 'Id');
                    if (docIndex < 0) {
                        return null;
                    } else {
                        // remove document form list if processed
                        const doc = allmaterials.splice(docIndex, 1);
                        return transformDocument(doc[0], agenda, webAbsoluteUrl);
                    }
                }).filter(d => McsUtil.isDefined(d));
            }
            return agenda;
        });

    // handle all subtopic
    agendaList.filter(a => McsUtil.isNumeric(a.ParentTopicId) && a.ParentTopicId > 0)
        .forEach(a => {
            // find parent item
            const index = McsUtil.binarySearch(mainTopics, a.ParentTopicId, 'Id');
            if (index >= 0) {
                const subTopic = transformAgenda(a, event.Id);
                //update last mofied if subtop was modified
                var agendaModified = new Date(a.Modified as string);
                if (agendaModified > lastModifiedDate) {
                    lastModifiedDate = agendaModified;
                }
                subTopic.MeetingPresenters = tranformPresenterList(McsUtil.isArray(a.PresentersLookupId as number[]) ? a.PresentersLookupId as number[] : [], presenterList);
                const fldname = business.get_AgendaDocumentLookupField();
                if (McsUtil.isArray(a[fldname])) {
                    subTopic.MeetingDocuments = (a[fldname] as number[]).map((d) => {
                        const docIndex = McsUtil.binarySearch(allmaterials, d, 'Id');
                        if (docIndex < 0) {
                            return null;
                        } else {
                            // remove document form list if processed
                            const doc = allmaterials.splice(docIndex, 1);
                            return transformDocument(doc[0], subTopic, webAbsoluteUrl);
                        }
                    }).filter(d => McsUtil.isDefined(d));
                }
                mainTopics[index].SubAgendaItems.push(subTopic);
            }
        });
    return modal;
};

export default transform;