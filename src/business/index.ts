import { sortBy } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDbCommittee, IDbStaff, IDbMembers } from "../interface/dbmodal";
import { McsUtil } from "../utility/helper";
import service from "../dal/service";
import lobService from "../dal/lobService";
import { ISpEvent, ISpAgendaTopic, ISpCommitteeLink, ISpEventMaterial, ISpPresenter, IDocumentItem } from '../interface/spmodal';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { tranformAgenda } from "./transformAgenda";
import { IFolderCreation } from '../dal/interface';

export interface IBusinessLogicConfig {
    spfxContext: WebPartContext;
}

class BusinessLogic {
    private _eventId: number;
    private _event: ISpEvent;
    // gettting initialized on load
    private _agendaList: ISpAgendaTopic[];
    // gettting initialized on load
    private _documentList: ISpEventMaterial[];
    // gettting initialized on load
    private _presenterList: ISpPresenter[];
    // gettting initialized on load
    private _committeeLists: ISpCommitteeLink[];
    // gettting initialized on load
    private _meetingCommittees: IDbCommittee[];
    private _config: IBusinessLogicConfig;
    private _documentFolderStructure: IFolderCreation;

    private _callbackOnLoaded: Array<(reject: any) => void>;

    private _onloadError: any;

    constructor() {
        this._config = {} as IBusinessLogicConfig;
        this._callbackOnLoaded = [];
    }

    /**
     * Initialize business logic with wepart context and ensure 
     * service is initialized as well. 
     * Also load data if required
     * @param {*} config
     * @memberof BusinessLogic
     */
    public setup(config: any): void {
        this._config = config;
        service.initialize();
        this._initialize();
    }

    /**
     * Wait for data to load
     * @param {() => void} callback
     * @memberof BusinessLogic
     */
    public on_Loaded(callback: (reject: any) => void): void {
        if (McsUtil.isDefined(this._event)) {
            callback(this._onloadError);
        } else {
            this._callbackOnLoaded.push(callback);
        }
    }

    /**
     * Get current Event
     *
     * @returns {ISpEvent}
     * @memberof BusinessLogic
     */
    public get_Event(): ISpEvent {
        return this._event;
    }

    /**
     * Get Event Agenda
     *
     * @returns {ISpAgendaTopic[]}
     * @memberof BusinessLogic
     */
    public get_Agenda(): ISpAgendaTopic[] {
        return this._agendaList;
    }

    /**
     * Get event documents
     *
     * @returns {ISpEventMaterial[]}
     * @memberof BusinessLogic
     */
    public get_Documents(): ISpEventMaterial[] {
        return this._documentList;
    }

    /**
     * Get event presenters
     *
     * @returns {ISpPresenter[]}
     * @memberof BusinessLogic
     */
    public get_Presenters(): ISpPresenter[] {
        return this._presenterList;
    }

    public get_Committee(): IDbCommittee[] {
        return this._meetingCommittees;
    }

    public add_Event(properties: ISpEvent): Promise<ISpEvent> {
        return new Promise((resolve, reject) => {
            service.addItemToSpList(Mcs.WebConstants.committeeCalendarListId, false, properties)
                .then((newevent: ISpEvent) => {
                    this._event = newevent;
                    this._eventId = newevent.Id;
                    const startdate = new Date(this._event.EventDate);
                    return this._loadEventCommittees(startdate.getFullYear());
                }).then(() => {
                    service.setIsSession(this.is_SessionMeeting());
                    resolve();
                }).catch(e => reject(e));
        });
    }

    public edit_Event(id: number, listItemEntityTypeFullName: string, propertiesToUpdate: any): Promise<ISpEvent> {
        return new Promise((resolve, reject) => {
            service.editItemSpList(Mcs.WebConstants.committeeCalendarListId, false, id, listItemEntityTypeFullName, propertiesToUpdate)
                .then((newevent: any) => {
                    if (this._event.IsBudgetHearing !== newevent.IsBudgetHearing) {
                        service.setIsSession(this.is_SessionMeeting());
                    }
                    const oldjcc = JSON.stringify(this._event.JointEventCommitteeId);
                    this._event = { ...this._event, };
                    if (oldjcc !== JSON.stringify(newevent.JointEventCommitteeId)) {
                        const startdate = new Date(this._event.EventDate);
                        return this._loadEventCommittees(startdate.getFullYear());
                    } else {
                        return Promise.resolve(null);
                    }
                }).then(() => {
                    resolve();
                }).catch((e) => reject(e));
        });
    }

    public add_Agenda(properties: any): Promise<ISpAgendaTopic> {
        return new Promise((resolve, reject) => {
            service.addItemToSpList(Mcs.WebConstants.agendaListId, false, properties)
                .then((newagenda: ISpAgendaTopic) => {
                    this._agendaList.push(newagenda);
                    resolve(newagenda);
                }).catch((e) => reject(e));
        });
    }

    public edit_Agenda(id: number, listItemEntityTypeFullName: string, propertiesToUpdate: any): Promise<ISpAgendaTopic> {
        return new Promise((resolve, reject) => {
            service.editItemSpList(Mcs.WebConstants.agendaListId, false, id, listItemEntityTypeFullName, propertiesToUpdate)
                .then((newagenda: ISpAgendaTopic) => {
                    for (let i = 0; i < this._agendaList.length; i++) {
                        if (this._agendaList[i].Id === id) {
                            this._agendaList[i] = newagenda;
                            break;
                        }
                    }
                    resolve(newagenda);
                }).catch((e) => reject(e));
        });
    }

    public delete_Agenda(id: number): Promise<void> {
        return new Promise((resolve, reject) => {
            service.deleteItemFromSpList(Mcs.WebConstants.agendaListId, false, id)
                .then(() => {
                    for (let i = 0; i < this._agendaList.length; i++) {
                        if (this._agendaList[i].Id === id) {
                            this._agendaList.splice(i, 1);
                            break;
                        }
                    }
                    resolve();
                }).catch((e) => reject(e));
        });
    }

    public add_Presenter(properties: ISpPresenter): Promise<ISpPresenter> {
        return new Promise((resolve, reject) => {
            service.addItemToSpList(Mcs.WebConstants.meetingPresenterListId, false, properties)
                .then((newpresenter: ISpPresenter) => {
                    this._presenterList.push(newpresenter);
                    resolve(newpresenter);
                }).catch((e) => reject(e));
        });
    }

    public edit_Presenter(id: number, listItemEntityTypeFullName: string, propertiesToUpdate: any): Promise<ISpPresenter> {
        return new Promise((resolve, reject) => {
            service.editItemSpList(Mcs.WebConstants.meetingPresenterListId, false, id, listItemEntityTypeFullName, propertiesToUpdate)
                .then((newpresenter: ISpPresenter) => {
                    for (let i = 0; i < this._presenterList.length; i++) {
                        if (this._presenterList[i].Id === id) {
                            this._presenterList[i] = newpresenter;
                            break;
                        }
                    }
                    resolve(newpresenter);
                }).catch((e) => reject(e));
        });
    }

    public delete_Presenter(id: number): Promise<void> {
        return new Promise((resolve, reject) => {
            service.deleteItemFromSpList(Mcs.WebConstants.meetingPresenterListId, false, id)
                .then(() => {
                    for (let i = 0; i < this._presenterList.length; i++) {
                        if (this._presenterList[i].Id === id) {
                            this._presenterList.splice(i, 1);
                            break;
                        }
                    }
                    resolve();
                }).catch((e) => reject(e));
        });
    }

    /**
     * Upload document to meeting folder in Intrim Document library
     *
     * @param {string} fileName
     * @param {IDocumentItem} propertiesToUpdate
     * @param {Blob} blob
     * @returns {Promise<IDocumentItem>}
     * @memberof BusinessLogic
     */
    public upLoad_Document(folderName: string, fileName: string, propertiesToUpdate: IDocumentItem, blob: Blob): Promise<IDocumentItem> {
        return new Promise((resolve, reject) => {
            let ensureFolderCreation = Promise.resolve({});
            const folderRelativeUrl = this._findServerRelativeUrl(folderName, this._documentFolderStructure);
            if (this._documentList.length === 0) {
                const startdate = new Date(this._event.EventDate);
                ensureFolderCreation = this._ensure_Folders(startdate.getFullYear(), this._event.Id);
            }
            ensureFolderCreation.then(() => {
                return service.get_MaterialService().addOrUpdateDocument(folderRelativeUrl, fileName, propertiesToUpdate, blob);
            }).then((newdocument: ISpEventMaterial) => {
                this._documentList.push(newdocument);
                resolve(newdocument);
            }).catch((e) => reject(e));
        });
    }

    public edit_Document(id: number, listItemEntityTypeFullName: string, propertiesToUpdate: any): Promise<IDocumentItem> {
        return new Promise((resolve, reject) => {
            service.editItemSpList(Mcs.WebConstants.meetingMaterialListId, false, id, listItemEntityTypeFullName, propertiesToUpdate)
                .then((newdocument: ISpEventMaterial) => {
                    for (let i = 0; i < this._presenterList.length; i++) {
                        if (this._documentList[i].Id === id) {
                            this._documentList[i] = newdocument;
                            break;
                        }
                    }
                    resolve(newdocument);
                }).catch((e) => reject(e));
        });
    }

    public find_Agency(name: string): Promise<any[]> {
        return service.searchAgencyList(name);
    }

    public is_SessionMeeting(): boolean {
        return this._event.IsBudgetHearing;
    }

    public get_FolderNameToUpload(documentType: string): string {
        if (this.is_SessionMeeting() && /^bill$/i.test(documentType)) {
            return `Bill Drafts for ${this._event.Id}`;
        }
        return `Material for ${this._event.Id}`;
    }

    /**
     * Try to read calendar item id from query string. If event id greater than 0 then load
     * event, agenda, and committee 
     * @private
     * @returns {Promise<void>}
     * @memberof BusinessLogic
     */
    private _initialize(): Promise<void> {
        return new Promise((resolve, reject) => {
            const queryParameters: UrlQueryParameterCollection = new UrlQueryParameterCollection(window.location.href);
            this._eventId = 0;
            if (McsUtil.isNumberString(queryParameters.getValue("calendaritemid"))) {
                this._eventId = parseInt(queryParameters.getValue("calendaritemid"), 10);
            }
            this._loadCommitteeLinks().then(() => {
                return Promise.all([this._loadEvent(), this._loadAgenda()]);
            }).then(() => {
                const startdate = new Date(this._event.EventDate);
                return Promise.all([this._loadIntrimDocuments(), this._loadPresenters(), this._loadEventCommittees(startdate.getFullYear())]);
            }).then(() => {
                tranformAgenda(this._agendaList, this._documentList, this._presenterList);
                while (this._callbackOnLoaded.length > 0) {
                    const cb = this._callbackOnLoaded.shift();
                    cb(undefined);
                }
                resolve();
            }).catch((e) => {
                this._onloadError = e;
                while (this._callbackOnLoaded.length > 0) {
                    const cb = this._callbackOnLoaded.shift();
                    cb(e);
                }
            });
        });
    }

    private _loadEvent(): Promise<void> {
        return new Promise((resolve, reject) => {
            if (this._eventId > 0) {
                service.getEvent(this._eventId).then((value) => {
                    let hasParentCommittee = false;
                    if (McsUtil.isNumeric(value.CommitteeEventLookupId) && value.CommitteeEventLookupId > 0 &&
                        McsUtil.isNumeric(value.CommitteeLookupId) && value.CommitteeLookupId > 0) {
                        const parentCommittee = this._committeeLists.filter(c => c.Id === value.CommitteeLookupId);
                        if (parentCommittee.length > 0) {
                            hasParentCommittee = true;
                            var url = parentCommittee[0].URL.Url + "/" + parentCommittee[0].CommitteeDesktopUrl + "?calendarItemId=" + value.CommitteeEventLookupId;
                            window.location.href = url;
                        }
                    } else {
                        this._event = value;
                        resolve();
                    }
                }).catch((e) => reject(e));
            } else {
                this._event = this._getDefaultEvent();
                resolve();
            }
        });
    }

    /**
     * Load agenda for event if event id is greater than 0
     * @returns {Promise<void>}
     * @memberof BusinessLogic
     */
    private _loadAgenda(): Promise<void> {
        return new Promise((resolve, reject) => {
            this._agendaList = [];
            if (this._eventId > 0) {
                service.getAgenda(this._eventId)
                    .then((result) => {
                        this._agendaList = result;
                        resolve();
                    }).catch(() => resolve());
            } else {
                resolve();
            }
        });
    }

    private _loadIntrimDocuments(): Promise<void> {
        return new Promise((resolve, reject) => {
            this._documentList = [];
            if (this._eventId > 0) {
                service.getMaterials(this._event)
                    .then((result) => {
                        this._documentList = result;
                        resolve();
                    }).catch(() => resolve());
            } else {
                resolve();
            }
        });
    }

    private _loadPresenters(): Promise<void> {
        return new Promise((resolve, reject) => {
            this._presenterList = [];
            if (this._eventId > 0 && McsUtil.isArray(this._agendaList) && this._agendaList.length > 0) {
                const presentersId: number[] = [];
                this._agendaList.forEach((tempAgenda) => {
                    const presentersLookupIds: number[] = tempAgenda.PresentersLookupId as number[];
                    if (McsUtil.isArray(presentersLookupIds)) {
                        presentersId.push(...presentersLookupIds);
                    }
                });

                service.getPresenters(presentersId)
                    .then((result) => {
                        this._presenterList = result;
                        resolve();
                    }).catch(() => resolve());
            } else {
                resolve();
            }
        });
    }

    /**
     * Load committee list from sharepoint ICS site, and committee link list
     *
     * @private
     * @param {string} [filter]
     * @returns {Promise<void>}
     * @memberof BusinessLogic
     */
    private _loadCommitteeLinks(): Promise<void> {
        return new Promise((resolve, reject) => {
            this._committeeLists = [];
            service.getCommitteeLinks()
                .then((result) => {
                    this._committeeLists = sortBy(result, c => c.Id);
                    resolve();
                }).catch(() => resolve());
        });
    }

    /**
     * Load committee information like members and staff information 
     * @param {number} year
     * @returns {Promise<void>}
     * @memberof BusinessLogic
     */
    private _loadEventCommittees(year: number): Promise<void> {
        return new Promise((resolve, reject) => {
            this._meetingCommittees = [];
            const committeedbConverter = a => {
                return {
                    DisplayOrder: McsUtil.isNumeric(a.DisplayOrder) ? parseInt(a.DisplayOrder) : 999,
                    FullName: a.URL.Description,
                    ShortName: a.CommitteeName,
                    Year: year,
                    Code: a.CommitteeId
                } as IDbCommittee;
            };
            const committeeArray = this._committeeLists.filter(f => f.CommitteeId == Mcs.WebConstants.committeeId).map(committeedbConverter);
            // [this._committeeLists[index]].map(committeedbConverter);
            if (McsUtil.isDefined(this._event.JointEventCommitteeId)) {
                const jointCommitteeIds: number[] = this._event.JointEventCommitteeId.split(";#").map(a => parseInt(a)).filter(a => !isNaN(a)).sort();
                committeeArray.push(...jointCommitteeIds.map(a => McsUtil.binarySearch(this._committeeLists, a, 'Id'))
                    .filter(a => a >= 0)
                    .map(a => committeedbConverter(this._committeeLists[a])));
            }

            Promise.all(committeeArray.map(c => this._getCommitteeMembers(year, c.Code)))
                .then((responses) => {
                    responses.forEach((m, i) => {
                        committeeArray[i].CommitteeMembers = m;
                    });
                    return Promise.all(committeeArray.map(c => this._getCommitteeStaff(year, c.Code)));
                }).then((responses) => {
                    responses.forEach((m, i) => {
                        committeeArray[i].Staff = m;
                    });
                    this._meetingCommittees = committeeArray;
                    resolve();
                }).catch(e => reject(e));
        });
    }

    private _getCommitteeMembers(year: number, committeeCode: string): Promise<IDbMembers[]> {
        return new Promise((resolve, reject) => {
            lobService.getData(this._config.spfxContext.serviceScope,
                McsUtil.combinePaths(Mcs.WebConstants.lsoServiceBase, '/api/Committees/', `${year}`, committeeCode))
                .then((response) => {
                    const sortedMembers = response.sort((a: IDbMembers, b: IDbMembers) => {
                        if (/h/gi.test(a.chamber) || /s/gi.test(a.chamber)) {
                            if (a.chamber === "s" || a.chamber === "S") return 0;
                            return 1;
                        }
                        return 2;
                    });
                    resolve(sortedMembers);
                }).catch((e) => resolve([]));
        });
    }

    private _getCommitteeStaff(year: number, committeeCode: string): Promise<IDbStaff[]> {
        return new Promise((resolve, reject) => {
            lobService.getData(this._config.spfxContext.serviceScope,
                McsUtil.combinePaths(Mcs.WebConstants.lsoServiceBase, 'api/Calendar/Committee', `${year}`, committeeCode))
                .then((response) => {
                    if (McsUtil.isArray(response.employees)) {
                        resolve(response.employees);
                    } else {
                        resolve([]);
                    }
                    response(response.data.employees);
                }).catch((e) => resolve([]));
        });
    }

    private _getCommittesChairperson(): string {
        const chairmenList = [];
        let chairperson = '';
        for (var i = 0; i < this._meetingCommittees.length; i++) {
            var committee = this._meetingCommittees[i];
            for (var j = 0; j < committee.CommitteeMembers.length; j++) {
                var member = committee.CommitteeMembers[j];
                if (member.isChairman) {
                    chairmenList.push({ name: member.legislatureName, chamber: member.chamber });
                }
            }
        }
        // var sortedMembers = chairmenList.sort((a, b) => {
        //     if (a.chamber === "h" || a.chamber === "H") return 0;
        //     if (a.chamber === "s" || a.chamber === "S") return 1;
        //     return 2;
        // });
        for (var k = 0; k < chairmenList.length - 1; k++) {
            chairperson += (chairmenList[k].chamber === "S" ? "Senator" : "Representative") + " " + chairmenList[k].name;
            if (k + 2 === chairmenList.length) {
                chairperson += " and ";
            } else {
                chairperson += ", ";
            }
        }

        if (chairmenList.length > 0) {
            chairperson += (chairmenList[chairmenList.length - 1].chamber === "S" ? "Senator" : "Representative") + " " + chairmenList[chairmenList.length - 1].name;
            if (chairmenList.length > 1) {
                chairperson += ", Co-chairmen of the " + Mcs.WebConstants.committeeFullName + ", have announced the Committee will meet:";
            } else {
                chairperson += ", Chairman of the " + Mcs.WebConstants.committeeFullName + ", has announced the Committee will meet:";
            }
        }
        return chairperson;
    }

    private _getDefaultEvent(): ISpEvent {
        let meetingdate = new Date();
        const queryParameters: UrlQueryParameterCollection = new UrlQueryParameterCollection(window.location.href);
        if (McsUtil.isNumberString(queryParameters.getValue("startdate"))) {
            try {
                const tempdate = new Date(queryParameters.getValue("startdate"));
                if (McsUtil.isDefined(tempdate)) {
                    meetingdate = tempdate;
                }
            } catch{ }
        }
        const startdate = new Date(meetingdate.getFullYear(), meetingdate.getMonth(), meetingdate.getDate(), 8, 30, 0, 0);
        const endate = new Date(meetingdate.getFullYear(), meetingdate.getMonth(), meetingdate.getDate(), 23, 0, 0, 0);
        return {
            Id: 0,
            ApprovedStatus: '(none)',
            CommitteeEventLookupId: null,
            CommitteeStaff: '',
            ConferenceNumber: '',
            Category: 'Tentative',
            WorkState: 'Wyoming',
            EventDate: (new Date(startdate.toLocaleDateString())).toISOString(),
            EndDate: (new Date(endate.toLocaleDateString())).toISOString(),
            MeetingStartTime: startdate.toLocaleTimeString(),
            Description: '',
            HasLiveStream: false,
            IsBudgetHearing: false,
            EventDocumentsLookupId: null,
            JointEventCommitteeId: null,
            Location: '',
            OtherLocationInfo: '',
            CommitteeLookupId: null,
            Title: '',
            WorkAddress: '',
            WorkCity: '',
            fAllDayEvent: true,
        };
    }

    /**
     * Ensure folders are created for Intrim Documents Library
     * @param {number} meetingYear
     * @param {number} meetingId
     * @returns {Promise<any>}
     * @memberof BusinessLogic
     */
    private _ensure_Folders(meetingYear: number, meetingId: number): Promise<any> {
        const folderStructure = this._getFolderStructure(meetingYear, meetingId);
        return new Promise((resolve, reject) => {
            service.folderCreation(folderStructure)
                .then((e) => {
                    this._documentFolderStructure = e;
                    resolve(e);
                })
                .catch((e) => {
                    this._documentFolderStructure = folderStructure;
                    reject(e);
                });
        });
    }

    private _getFolderStructure(meetingYear: number, meetingId: number): IFolderCreation {
        if (this.is_SessionMeeting()) {
            return {
                name: meetingYear.toString(),
                SubFolder: [
                    { name: 'Agency Budget Requests' },
                    { name: 'Agency Handouts' },
                    {
                        name: 'Bill Drafts',
                        SubFolder: [{ name: `Bill Drafts for ${meetingId}` }]
                    },
                    { name: 'Agency' },
                    { name: 'LSO Analysis' },
                    { name: 'Citizen or Lobbyist Handouts' },
                    { name: 'Executive Letters' },
                    { name: 'Post-Session Summaries' },
                    { name: 'Pre-Session Materials' },
                    {
                        name: 'Meetings',
                        SubFolder: [{ name: `Material for ${meetingId}` }]
                    }
                ]
            };
        } else {
            return {
                name: meetingYear.toString(),
                SubFolder: [
                    { name: 'Correspondence' },
                    {
                        name: 'Meetings',
                        SubFolder: [{ name: `Material for ${meetingId}` }]
                    },
                    { name: 'Work Products' },
                    { name: 'Reports' },
                ]
            };
        }
    }

    private _findServerRelativeUrl(name: string, folder: IFolderCreation): string {
        if (McsUtil.isDefined(folder)) {
            if (folder.name == name) {
                return folder.ServerRelativeUrl;
            } else {
                if (McsUtil.isArray(folder.SubFolder) && folder.SubFolder.length > 0) {
                    for (let i = 0; i < folder.SubFolder.length; i++) {
                        const serverRelativeUrl = this._findServerRelativeUrl(name, folder.SubFolder[i]);
                        if (McsUtil.isString(serverRelativeUrl)) {
                            return serverRelativeUrl;
                        }
                        if (i + 1 == folder.SubFolder.length) {
                            return undefined;
                        }
                    }
                } else {
                    return undefined;
                }
            }
        } else {
            return undefined;
        }
    }

}

export const business: BusinessLogic = new BusinessLogic();
