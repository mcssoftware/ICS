import { sortBy, findIndex } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDbCommittee, IDbStaff, IDbMembers, IDbMeeting } from "../interface/dbmodal";
import { McsUtil } from "../utility/helper";
import service from "../dal/service";
import lobService from "../dal/lobService";
import { ISpEvent, ISpAgendaTopic, ISpCommitteeLink, ISpEventMaterial, ISpPresenter, IDocumentItem, IBillVersion } from '../interface/spmodal';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { tranformAgenda } from "./transformAgenda";
import { transformSpToDb } from "./transformSpToDb";
import { IFolderCreation } from '../dal/interface';
import IcsAppConstants from '../configuration';
import { BlobParser } from '../utility/parser';

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
    private _currentLmsUrl: string;

    private _callbackOnLoaded: Array<(reject: any) => void>;

    private _onloadError: any;

    constructor() {
        this._config = {} as IBusinessLogicConfig;
        this._callbackOnLoaded = [];
        this._currentLmsUrl = "";
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
        McsUtil.setTimeOffset(config.spfxContext.pageContext.web.timeZoneInfo.offset + config.spfxContext.pageContext.web.timeZoneInfo.daylightOffset);
        service.initialize();
        this._initialize();
    }

    public getWebUrl(): string {
        return this._config.spfxContext.pageContext.web.absoluteUrl;
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

    public loadEvent(): Promise<void> {
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
     * Get event document by type
     *
     * @returns {ISpEventMaterial[]}
     * @memberof BusinessLogic
     */
    public get_DocumentByType(documenttype: string): ISpEventMaterial {
        const filterDoc = this._documentList.filter(a => a.lsoDocumentType == documenttype);
        return filterDoc.length > 0 ? filterDoc[0] : null;
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

    public get_CommitteeList(): ISpCommitteeLink[] {
        return this._committeeLists;
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
                    resolve();
                }).catch(e => reject(e));
        });
    }

    public edit_Event(id: number, listItemEntityTypeFullName: string, propertiesToUpdate: any): Promise<ISpEvent> {
        return new Promise((resolve, reject) => {
            service.editItemSpList(Mcs.WebConstants.committeeCalendarListId, false, id, listItemEntityTypeFullName, propertiesToUpdate)
                .then((newevent: any) => {
                    const oldjcc = JSON.stringify(this._event.JointEventCommitteeId);
                    this._event = newevent;
                    if (oldjcc !== JSON.stringify(newevent.JointEventCommitteeId)) {
                        const startdate = new Date(newevent.EventDate);
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
            let ensureDocumentPromise: Promise<any>;

            if (this._documentList.length === 0 && this._agendaList.length < 4) {
                const startdate = new Date(this._event.EventDate);
                ensureDocumentPromise = this._ensure_Folders(startdate.getFullYear(), undefined);
            } else {
                ensureDocumentPromise = Promise.resolve({});
            }

            Promise.all([ensureDocumentPromise, service.addItemToSpList(Mcs.WebConstants.agendaListId, false, properties)])
                .then((responses) => {
                    const newagenda: ISpAgendaTopic = responses[1] as ISpAgendaTopic;
                    this._agendaList.push(newagenda);
                    tranformAgenda(this._agendaList, this._documentList, this._presenterList);
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
                    tranformAgenda(this._agendaList, this._documentList, this._presenterList);
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
                    tranformAgenda(this._agendaList, this._documentList, this._presenterList);
                    resolve();
                }).catch((e) => reject(e));
        });
    }

    public add_Presenter(properties: ISpPresenter): Promise<ISpPresenter> {
        return new Promise((resolve, reject) => {
            service.addItemToSpList(Mcs.WebConstants.meetingPresenterListId, false, properties)
                .then((newpresenter: ISpPresenter) => {
                    this._presenterList.push(newpresenter);
                    tranformAgenda(this._agendaList, this._documentList, this._presenterList);
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
                    tranformAgenda(this._agendaList, this._documentList, this._presenterList);
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
                    tranformAgenda(this._agendaList, this._documentList, this._presenterList);
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
            service.setIsSession(this.is_SessionMeeting());
            let ensureFolderCreation = Promise.resolve({});

            if (!McsUtil.isDefined(this._documentFolderStructure) || this._documentList.length === 0) {
                const startdate = new Date(this._event.EventDate);
                ensureFolderCreation = this._ensure_Folders(startdate.getFullYear(), this._event.Id);
            }
            ensureFolderCreation.then(() => {
                const folderRelativeUrl = this._findServerRelativeUrl(folderName, this._documentFolderStructure);
                return service.get_MaterialService().addOrUpdateDocument(folderRelativeUrl, fileName, propertiesToUpdate, blob);
            }).then((newdocument: ISpEventMaterial) => {
                if (!/preview.pdf/i.test(fileName)) {
                    this._documentList.push(newdocument);
                    tranformAgenda(this._agendaList, this._documentList, this._presenterList);
                }
                resolve(newdocument);
            }).catch((e) => reject(e));
        });
    }

    public edit_Document(id: number, listItemEntityTypeFullName: string, propertiesToUpdate: any): Promise<IDocumentItem> {
        return new Promise((resolve, reject) => {
            service.setIsSession(this.is_SessionMeeting());
            service.editItemSpList(service.get_MaterialService().getListTitle(), false, id, listItemEntityTypeFullName, propertiesToUpdate)
                .then((newdocument: ISpEventMaterial) => {
                    for (let i = 0; i < this._documentList.length; i++) {
                        if (this._documentList[i].Id === id) {
                            this._documentList[i] = newdocument;
                            break;
                        }
                    }
                    tranformAgenda(this._agendaList, this._documentList, this._presenterList);
                    resolve(newdocument);
                }).catch((e) => reject(e));
        });
    }

    public delete_Document(id: number): void {
        if (McsUtil.isUnsignedInt(id)) {
            const index = findIndex(this._documentList, a => a.Id === id);
            if (index > -1) {
                this._documentList.splice(index, 1);
            }
        }
    }

    public find_Agency(name: string): Promise<any[]> {
        return service.searchAgencyList(name);
    }

    public find_Bill(lsonumber?: string): Promise<any[]> {
        return service.getLmsBill(this._currentLmsUrl, lsonumber);
    }

    public find_BillItemVersion(billId: number): Promise<IBillVersion[]> {
        return service.getBillVersion(this._currentLmsUrl, billId);
    }

    public find_Document(agencyName?: string, agencyNumber?: string, search?: string): Promise<any[]> {
        return new Promise((resolve, reject) => {
            service.setIsSession(this.is_SessionMeeting());
            service.get_MaterialService().loadList().then((listresult) => {
                const startdate = new Date(this._event.EventDate);
                const folderNameToSearch = McsUtil.combinePaths(listresult.RootFolder.ServerRelativeUrl, startdate.getFullYear().toString());
                let filter = `startswith(FileDirRef, '${folderNameToSearch}')`;
                if (McsUtil.isString(agencyName)) {
                    filter = `${filter} and substringof('${agencyName.replace("'", "''")}',AgencyName)`;
                }
                const tempFilter = [];
                if (McsUtil.isString(search)) {
                    tempFilter.push(`substringof('${search}',FileLeafRef) or substringof('${search}',Title)`);
                }
                if (McsUtil.isString(agencyNumber)) {
                    tempFilter.push(`substringof('${agencyNumber}',FileLeafRef) or substringof('${agencyNumber}',Title)`);
                }
                if (tempFilter.length > 0) {
                    filter = `${filter} and (${tempFilter.join(' or ')})`;
                }
                return service.get_MaterialService().getListItems(filter, null, null, null, null, 100);
            }).then((searchResult) => {
                const filteredSearchResult = searchResult.filter(a => a.FSObjType == 0 && (findIndex(this._documentList, d => d.Id == a.Id) < 0));
                resolve(filteredSearchResult);
            }).catch((e) => reject(e));
        });
    }

    public is_SessionMeeting(): boolean {
        return this._event.IsBudgetHearing;
    }

    public can_CreateBudgetMeeting(): boolean {
        return Mcs.WebConstants.committeeId == '02';
    }

    public get_FolderNameToUpload(documentType: string): string {
        if (documentType === IcsAppConstants.getPreviewFolder()) {
            return documentType;
        }
        if (this.is_SessionMeeting() && /^bill$/i.test(documentType)) {
            return `Bill Drafts for ${this._event.Id}`;
        }
        return `Material for ${this._event.Id}`;
    }

    public getDocumentFromIntranet(serverRelativeUrl: string, versionLabel?: string): Promise<Blob> {
        const url = McsUtil.combinePaths(Mcs.WebConstants.lmsServiceBase,
            'api/Intranet/GetIntranetDocument?&clean=true' + `&webUrl=${this._currentLmsUrl}&serverRelativeUrl=${serverRelativeUrl}` +
            (McsUtil.isString(versionLabel) ? `&version=${versionLabel}` : ''));
        return lobService.getBlob(this._config.spfxContext.serviceScope, url);
    }

    public get_EventDocumentLookupField(): string {
        if (this.is_SessionMeeting()) {
            return "EventSessionDocLookupId";
        } else {
            return "EventDocumentsLookupId";
        }
    }

    public get_AgendaDocumentLookupField(): string {
        if (this.is_SessionMeeting()) {
            return "AgendaSessionDocLookupId";
        } else {
            return "AgendaDocumentsLookupId";
        }
    }

    public generateMeetingDocument(partialUrl: string, contentType: string, data?: any): Promise<Blob> {
        return new Promise((resolve, reject) => {
            if (!McsUtil.isDefined(data)) {
                data = this.get_publishingMeeting();
            }
            if (typeof contentType === "string" && contentType.length === 0) {
                contentType = "application/json";
            }
            lobService.postData(this._config.spfxContext.serviceScope, McsUtil.combinePaths(Mcs.WebConstants.icsServiceBase, partialUrl),
                data, contentType, "Blob", new BlobParser())
                .then((response) => {
                    resolve(response);
                }).catch((e) => reject(e));
        });
    }

    public get_publishingMeeting(): IDbMeeting {
        const dbPostObject = transformSpToDb(this._config.spfxContext.pageContext.web.absoluteUrl, this._event, this._agendaList, this._documentList, this._presenterList);
        dbPostObject.Chairmen = this._getCommittesChairperson();
        dbPostObject.CommitteeList = this._meetingCommittees;
        return dbPostObject;
    }

    public publishDocument(item: IDocumentItem): Promise<void> {
        return new Promise((resolve) => {
            if (McsUtil.isDefined(item) && item.File.CheckOutType !== 2) {
                service.approveFile(item.File.ServerRelativeUrl, "publish(comment='auto publish file')", "approve(comment='auto approve file')")
                    .then(() => {
                        return service.get_MaterialService().getListItemById(item.Id);
                    }).then((newitem) => {
                        const index = findIndex(this._documentList, d => d.Id == item.Id);
                        this._documentList[index] = newitem;
                        resolve();
                    }).catch(() => {
                        resolve();
                    });
            } else {
                resolve();
            }
        });
    }

    public get_MeetingApprovalListService() {
        return service.get_MeetingApprovalListService();
    }

    public get_folderServerRelativeUrl(name: string): Promise<string> {
        return new Promise((resolve) => {
            if (this._eventId > 0 && this._documentList.length > 0) {
                if (McsUtil.isDefined(this._documentFolderStructure)) {
                    resolve(McsUtil.makeAbsUrl(this._findServerRelativeUrl(name, this._documentFolderStructure)));
                } else {
                    const startdate = new Date(this._event.EventDate);
                    this._ensure_Folders(startdate.getFullYear(), this._event.Id)
                        .then(() => {
                            if (McsUtil.isDefined(this._documentFolderStructure)) {
                                resolve(McsUtil.makeAbsUrl(this._findServerRelativeUrl(name, this._documentFolderStructure)));
                            } else {
                                resolve('');
                            }
                        }).catch(() => resolve(''));
                }
            } else {
                resolve('');
            }
        });
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
            this._getSiteConfiguration();
            this._loadCommitteeLinks()
                .then(() => {
                    return Promise.all([this.loadEvent(), this._loadAgenda()]);
                }).then(() => {
                    const startdate = new Date(this._event.EventDate);
                    return Promise.all([this._loadIntrimDocuments(), this._loadPresenters(), this._loadEventCommittees(startdate.getFullYear())]);
                }).then(() => {
                    if (this._event.Id == 0) {
                        const staff = [];
                        this._meetingCommittees.forEach((c) => {
                            if (McsUtil.isArray(c.Staff) && c.Staff.length > 0) {
                                staff.push(...c.Staff.filter(s => McsUtil.isString(s.name)).map(s => s.name));
                            }
                        });
                        if (staff.length > 1) {
                            const lastStaff = staff.pop();
                            this._event.CommitteeStaff = staff.join(', ') + ' and ' + lastStaff;
                        } else {
                            this._event.CommitteeStaff = staff.join('');
                        }
                    }

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
                service.setIsSession(this.is_SessionMeeting());
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
                McsUtil.combinePaths(Mcs.WebConstants.lsoServiceBase, IcsAppConstants.getCommitteesPartial(), `${year}`, committeeCode))
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
                McsUtil.combinePaths(Mcs.WebConstants.lsoServiceBase, IcsAppConstants.getCommitteesStaffPartial(), `${year}`, committeeCode))
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

    private _getSiteConfiguration(): Promise<void> {
        return new Promise((resolve, reject) => {
            lobService.getData(this._config.spfxContext.serviceScope,
                McsUtil.combinePaths(Mcs.WebConstants.lmsServiceBase, 'api/Intranet/GetSiteConfiguration'))
                .then((response) => {
                    this._currentLmsUrl = response["CURRENT LMS URL"];
                    resolve();
                }).catch((e) => resolve());
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
        if (McsUtil.isString(queryParameters.getValue("startdate"))) {
            try {
                const tempdate = new Date(queryParameters.getValue("startdate"));
                if (McsUtil.isDefined(tempdate)) {
                    meetingdate = tempdate;
                }
            } catch{ }
        }
        const startdate = new Date(meetingdate.getFullYear(), meetingdate.getMonth(), meetingdate.getDate(), 8, 30, 0, 0);
        const isostring = startdate.toISOString().split('T')[0];
        return {
            Id: 0,
            ApprovedStatus: '(none)',
            CommitteeEventLookupId: null,
            CommitteeStaff: '',
            ConferenceNumber: '',
            Category: 'Tentative',
            WorkState: 'Wyoming',
            EventDate: `${isostring}T00:00:00Z`,
            EndDate: `${isostring}T23:59:00Z`,
            MeetingStartTime: startdate.toLocaleTimeString(),
            Description: '',
            HasLiveStream: false,
            IsBudgetHearing: false,
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
    private _ensure_Folders(meetingYear: number, meetingId?: number): Promise<any> {
        const folderStructure = this._getFolderStructure(meetingYear, meetingId);
        return new Promise((resolve, reject) => {
            service.setIsSession(this.is_SessionMeeting());
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

    private _getFolderStructure(meetingYear: number, meetingId?: number): IFolderCreation {
        let meetingFolders;
        if (this.is_SessionMeeting()) {
            meetingFolders = {
                name: meetingYear.toString(),
                SubFolder: [
                    { name: 'Agency Budget Requests' },
                    { name: 'Agency Handouts' },
                    { name: 'Agency' },
                    { name: 'LSO Analysis' },
                    { name: 'Citizen or Lobbyist Handouts' },
                    { name: 'Executive Letters' },
                    { name: 'Post-Session Summaries' },
                    { name: 'Pre-Session Materials' },
                    { name: IcsAppConstants.getPreviewFolder() }
                ]
            };
            if (McsUtil.isUnsignedInt(meetingId)) {
                meetingFolders.SubFolder.push(
                    {
                        name: 'Bill Drafts',
                        SubFolder: [{ name: `Bill Drafts for ${meetingId}` }],
                    },
                    {
                        name: 'Meetings',
                        SubFolder: [{ name: `Material for ${meetingId}` }]
                    });
            } else {
                meetingFolders.SubFolder.push({ name: 'Bill Drafts' }, { name: 'Meetings' });
            }
        } else {
            meetingFolders = {
                name: meetingYear.toString(),
                SubFolder: [
                    { name: 'Correspondence' },
                    { name: 'Work Products' },
                    { name: 'Reports' },
                    { name: IcsAppConstants.getPreviewFolder() }
                ]
            };
            if (McsUtil.isUnsignedInt(meetingId)) {
                meetingFolders.SubFolder.push(
                    {
                        name: 'Meetings',
                        SubFolder: [{ name: `Material for ${meetingId}` }],
                    });
            } else {
                meetingFolders.SubFolder.push({ name: 'Meetings' });
            }
        }
        return meetingFolders;
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
