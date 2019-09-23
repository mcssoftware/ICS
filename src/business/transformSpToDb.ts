import { IDbMeeting, IDbAgenda, IDbDocument, IDbPresenter } from "../interface/dbmodal";
import { ISpEvent, ISpAgendaTopic, ISpEventMaterial, ISpPresenter } from "../interface/spmodal";
import { McsUtil } from "../utility/helper";
import { tranformAgenda, IComponentAgenda } from "./transformAgenda";
import { sortBy } from "@microsoft/sp-lodash-subset";

const transformEventDbFormat = (model: ISpEvent): IDbMeeting => {

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

const transformDocumentDbFormat = (material: ISpEventMaterial, agenda: IDbAgenda, webAbsoluteUrl: string): IDbDocument => {
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

const transformPresenterDbFormat = (presenter: ISpPresenter): IDbPresenter => {
    var model: IDbPresenter = {
        Office365ID: presenter.Id,
        Title: presenter.Title,
        Name: presenter.PresenterName,
        Organization: presenter.OrganizationName,
        SortNumber: presenter.SortNumber || 1
    };
    return model;
};

const tranformPresenterListDbFormat = (ids: number[], presenterList: ISpPresenter[]): IDbPresenter[] => {
    if (McsUtil.isArray(ids) && ids.length > 0) {
        return ids.map(a => {
            const index = McsUtil.binarySearch(presenterList, a, 'Id');
            if (index >= 0) {
                return transformPresenterDbFormat(presenterList[index]);
            } else {
                return null;
            }
        }).filter(a => McsUtil.isDefined(a));
    }
    return [];
};

const transformAgendaDbFormat = (webAbsoluteUrl: string, agenda: IComponentAgenda, meetingId: number): IDbAgenda => {
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
    subagenda.MeetingDocuments = agenda.Documents.map((d) => transformDocumentDbFormat(d, subagenda, webAbsoluteUrl));
    subagenda.MeetingPresenters = tranformPresenterListDbFormat(agenda.PresentersLookupId, agenda.Presenters);
    return subagenda;
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
export const transformSpToDb = (webAbsoluteUrl: string, event: ISpEvent, agendaList: ISpAgendaTopic[], materialList: ISpEventMaterial[], presenterList: ISpPresenter[]): IDbMeeting => {
    debugger;
    const modal = transformEventDbFormat(event);
    let lastModifiedDate = new Date(event.Modified as string);

    const allmaterials = sortBy(materialList, a => a.Id);

    const tempAgenda = tranformAgenda(agendaList, materialList, presenterList);
    modal.MeetingAgendas = tempAgenda.map((a): IDbAgenda => {
        const newAgenda = transformAgendaDbFormat(webAbsoluteUrl, a, event.Id);
        var agendaModified = new Date(a.Modified as string);
        if (agendaModified > lastModifiedDate) {
            lastModifiedDate = agendaModified;
        }
        a.Documents.forEach((d) => {
            const docIndex = McsUtil.binarySearch(allmaterials, d.Id, 'Id');
            if (docIndex >= 0) {
                // remove document form list if processed
                allmaterials.splice(docIndex, 1);
            }
        });
        newAgenda.SubAgendaItems = a.SubTopics.map((s) => {
            s.Documents.forEach((d1) => {
                const docIndex = McsUtil.binarySearch(allmaterials, d1.Id, 'Id');
                if (docIndex >= 0) {
                    // remove document form list if processed
                    allmaterials.splice(docIndex, 1);
                }
            });
            return transformAgendaDbFormat(webAbsoluteUrl, s, event.Id);
        });

        return newAgenda;
    });
    modal.LastUpdated = lastModifiedDate;
    modal.MeetingDocuments = allmaterials.map(m => transformDocumentDbFormat(m, null, webAbsoluteUrl));
    return modal;
};
