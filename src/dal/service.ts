import { ISpEvent, ISpAgendaTopic, ISpPresenter, ISpEventMaterial, ISpCommitteeLink, IListItem } from "../interface/spmodal";
import SpListService from "./spListService";
import { uniq, sortBy } from "@microsoft/sp-lodash-subset";
import { McsUtil } from "../utility/helper";
import { IFolderCreation } from "./interface";

const sessionDocumentLibrary: string = 'Session Document';

class Service {
    private _eventService: SpListService<ISpEvent>;
    private _agendaService: SpListService<ISpAgendaTopic>;
    private _presenterService: SpListService<ISpPresenter>;
    private _meetingMaterialService: SpListService<ISpEventMaterial>;
    private _committeeLinkService: SpListService<ISpCommitteeLink>;
    private _agencyList: SpListService<any>;


    constructor() {
    }

    public initialize(): void {
        this._eventService = new SpListService<ISpEvent>(Mcs.WebConstants.committeeCalendarListId, false);
        this._agendaService = new SpListService<ISpAgendaTopic>(Mcs.WebConstants.agendaListId, false);
        this._presenterService = new SpListService<ISpPresenter>(Mcs.WebConstants.meetingPresenterListId, false);
        // initialize meeting material service after event is loaded
        // this._meetingMaterialService = new SpListService<ISpEventMaterial>(Mcs.WebConstants.meetingPresenterListId, false);
        this._committeeLinkService = new SpListService<ISpCommitteeLink>('Committee Links', true);

        //initialize agencyList
        this._agencyList = new SpListService<any>('Agency Contact', false);
        this._agencyList.setWebUrl('https://wyoleg.sharepoint.com/sites/lms');

        this._initializeSelects();
    }

    public setIsSession(isSession: boolean): void {
        if (isSession) {
            this._meetingMaterialService = new SpListService<ISpEventMaterial>(sessionDocumentLibrary, false);
        } else {
            this._meetingMaterialService = new SpListService<ISpEventMaterial>(Mcs.WebConstants.meetingPresenterListId, false);
        }

        this._meetingMaterialService.getSelects = () => {
            return ['Id', 'AgencyName', 'IncludeWithAgenda', 'SortNumber', 'Title', 'lsoDocumentType'];
        };
        this._meetingMaterialService.getExpands = () => {
            return ['File'];
        };
    }

    public getEvent(id: number): Promise<ISpEvent> {
        return this._eventService.getListItemById(id);
    }

    public getAgenda(eventId: number): Promise<ISpAgendaTopic[]> {
        return new Promise((resolve, reject) => {
            let agendaList: ISpAgendaTopic[] = [];
            this._agendaService.getListItems(`EventLookup/Id eq ${eventId}`, null, null,
                [{ Field: 'Id', IsAscending: true }])
                .then((response) => {
                    agendaList = response;
                    if (agendaList.length > 0) {
                        let presenterIds: number[] = [];
                        agendaList.forEach((e) => {
                            const temppresenters = e.PresentersLookupId as number[];
                            if (McsUtil.isArray(temppresenters) && temppresenters) {
                                presenterIds = presenterIds.concat(temppresenters);
                            }
                        });
                        if (presenterIds.length > 0) {
                            return this.getPresenters(presenterIds);
                        } else {
                            return Promise.resolve([]);
                        }
                    } else {
                        return Promise.resolve([]);
                    }
                }).then((presenters) => {
                    if (McsUtil.isArray(presenters) && presenters.length > 0) {
                        agendaList.forEach(a => {
                            a.Presenters = [];
                            const temppresenters = a.PresentersLookupId as number[];
                            if (McsUtil.isArray(temppresenters) && temppresenters.length > 0) {
                                temppresenters.forEach(p => {
                                    const index = McsUtil.binarySearch(presenters, p, 'Id');
                                    if (index > -1) {
                                        a.Presenters.push(presenters[index]);
                                    }
                                });
                            }
                        });
                    } else {
                        agendaList.forEach((a) => a.Presenters = []);
                    }
                    resolve(agendaList);
                }).catch((e) => reject(e));
        });

    }

    public getPresenters(presenterIds: number[]): Promise<ISpPresenter[]> {
        return new Promise((resolve, reject) => {
            if (McsUtil.isArray(presenterIds) && presenterIds.length > 0) {
                Promise.all(McsUtil.chunkArray(uniq(presenterIds), 30).map((e) => {
                    if (e.length > 0) {
                        const filter = e.map((id) => `Id eq ${id}`).join(' or ');
                        return this._presenterService.getListItems(filter, null, null, [{ Field: 'Id', IsAscending: true }]);
                    } else {
                        return Promise.resolve([]);
                    }
                })).then((result) => {
                    let presenterLists: ISpPresenter[] = [];
                    result.forEach((e) => {
                        presenterLists = presenterLists.concat(e);
                    });
                    resolve(sortBy(presenterLists, (e) => e.Id));
                }).catch((e) => reject(e));
            } else {
                resolve([]);
            }
        });
    }

    public getMaterials(event: ISpEvent): Promise<ISpEventMaterial[]> {
        return new Promise((resolve, reject) => {
            if (McsUtil.isDefined(event) && McsUtil.isDefined(event.EventDocumentsLookupId)
                && McsUtil.isArray(event.EventDocumentsLookupId) && (event.EventDocumentsLookupId as number[]).length > 0) {
                Promise.all(McsUtil.chunkArray(event.EventDocumentsLookupId as number[], 30).map((d) => {
                    const filter = d.map((id) => `Id eq ${id}`).join(' or ');
                    return this._meetingMaterialService.getListItems(filter);
                })).then((responses) => {
                    let documentList: ISpEventMaterial[] = [];
                    responses.forEach((d) => {
                        documentList = documentList.concat(d);
                    });
                    resolve(sortBy(documentList, d => d.Id));
                }).catch((e) => reject(e));
            } else {
                resolve([]);
            }
        });
    }

    public getCommitteeLinks(filter?: string): Promise<ISpCommitteeLink[]> {
        if (McsUtil.isString(filter)) {
            return this._committeeLinkService.getListItems(filter);
        } else {
            return this._committeeLinkService.getListItems();
        }
    }

    public addItemToSpList(listTitle: string, isRootSiteList: boolean, properties: any): Promise<IListItem> {
        return new Promise((resolve, reject) => {
            const listService: SpListService<IListItem> = new SpListService<IListItem>(listTitle, isRootSiteList);
            listService.addNewItem(properties)
                .then((result) => {
                    const service = this.get_SpService(listTitle);
                    if (service === null) {
                        return Promise.resolve(result);
                    } else {
                        return service.getListItemById(result.Id);
                    }
                }).then((newitem) => {
                    resolve(newitem);
                }).catch((e) => { reject(e); });
        });
    }

    public editItemSpList(listTitle: string, isRootSiteList: boolean, id: number, listItemEntityTypeFullName: string, propertiesToUpdate: any): Promise<IListItem> {
        return new Promise((resolve, reject) => {
            const listService: SpListService<IListItem> = new SpListService<IListItem>(listTitle, isRootSiteList);
            listService.updateItem(id, listItemEntityTypeFullName, propertiesToUpdate)
                .then(() => {
                    const service = this.get_SpService(listTitle);
                    if (service === null) {
                        return Promise.resolve(propertiesToUpdate);
                    } else {
                        return service.getListItemById(id);
                    }
                }).then((newitem) => {
                    resolve(newitem);
                }).catch((e) => { reject(e); });
        });
    }

    public deleteItemFromSpList(listTitle: string, isRootSiteList: boolean, id: number): Promise<void> {
        return new Promise((resolve, reject) => {
            const listService: SpListService<IListItem> = new SpListService<IListItem>(listTitle, isRootSiteList);
            listService.deleteItem(id)
                .then(() => {
                    resolve();
                }).catch((e) => { reject(e); });
        });
    }

    public folderCreation(folderStructure: IFolderCreation): Promise<IFolderCreation> {
        return new Promise((resolve, reject) => {
            this._meetingMaterialService.loadList()
                .then((listresult) => {
                    return this._folderCreation(listresult.RootFolder.ServerRelativeUrl, folderStructure);
                }).then((result) => {
                    resolve(result);
                }).catch((e) => reject(e));
        });
    }

    public searchAgencyList(agencyName: string): Promise<any[]> {
        if (!McsUtil.isString(agencyName)) {
            return this._agencyList.getListItems(null, null, null, [{ Field: 'AgencyName', IsAscending: true }], 0, 50);
        } else {
            return this._agencyList.getListItems(`IsAgencyDirector eq 1 and (substringof('${agencyName}',AgencyName) or substringof('${agencyName}',Title))`, null, null,
                [{ Field: 'AgencyName', IsAscending: true }], 0, 50);
        }
    }

    public get_MaterialService(): SpListService<ISpEventMaterial> {
        return this._meetingMaterialService;
    }

    public get_SpService(listTitle: string): SpListService<any> {
        if ((listTitle === Mcs.WebConstants.committeeCalendarListId) || (listTitle === this._eventService.getListTitle())) {
            return this._eventService;
        } else {
            if ((listTitle === Mcs.WebConstants.agendaListId) || (listTitle === this._agendaService.getListTitle())) {
                return this._agendaService;
            } else {
                if ((listTitle === Mcs.WebConstants.meetingPresenterListId) || (listTitle === this._presenterService.getListTitle())) {
                    return this._presenterService;
                } else {
                    return null;
                }
            }
        }
    }

    private _folderCreation(folderServerUrl: string, folderStructure: IFolderCreation): Promise<IFolderCreation> {
        return new Promise((resolve, reject) => {
            this._meetingMaterialService.getFolder(folderServerUrl, folderStructure.name)
                .then((result) => {
                    if (result.length > 0) {
                        return Promise.resolve(result[0]);
                    } else {
                        return this._meetingMaterialService.createFolder(folderServerUrl, folderStructure.name);
                    }
                }).then((newfolder) => {
                    folderStructure.ServerRelativeUrl = newfolder.ServerRelativeUrl;
                    if (McsUtil.isArray(folderStructure.SubFolder) && folderStructure.SubFolder.length > 0) {
                        Promise.all(folderStructure.SubFolder
                            .map(f => this._folderCreation(folderStructure.ServerRelativeUrl, f)))
                            .then(() => {
                                resolve(folderStructure);
                            }).catch((e) => reject(e));

                    } else {
                        resolve(folderStructure);
                    }
                }).catch((e) => { reject(e); });

        });
    }

    private _initializeSelects(): void {
        this._eventService.getSelects = () => {
            return ['Id', 'ApprovedStatus', 'Category', 'CommitteeEventLookupId', 'CommitteeLookupId', 'CommitteeStaff', 'ConferenceNumber', 'Description',
                'EndDate', 'EventDate', 'EventDocumentsLookupId', 'Id', 'IsBudgetHearing', 'JointEventCommitteeId', 'Location', 'MeetingStartTime',
                'OtherLocationInfo', 'ParticipantsPickerId', 'Title', 'WorkAddress', 'WorkCity', 'WorkState', 'fAllDayEvent', 'fRecurrence', 'Modified'];
        };

        this._agendaService.getSelects = () => {
            return ['Id', 'AgendaDate', 'AgendaDocumentsLookupId', 'AgendaNumber', 'AgendaTitle', 'AllowPublicComments', 'EventLookupId',
                'ParentTopicId', 'PresentersLookupId', 'Title', 'Modified'];
        };

        // this._committeeLinks.getSelects = () => {
        //     return ['Id', 'CommitteeDesktopUrl', 'CommitteeId', 'CommitteeName', 'DisplayOrder'];
        // };

        this._presenterService.getSelects = () => {
            return ['Id', 'PresenterName', 'Title', 'OrganizationName', 'SortNumber'];
        };

        this._committeeLinkService.getSelects = () => {
            return ['Id', 'Title', 'URL', 'CommitteeId', 'CommitteeName', 'DisplayOrder', 'CommitteeDesktopUrl'];
        };

        this._agencyList.getSelects = () => {
            return ['AgencyName', 'IsAgencyDirector', 'Title'];
        };
    }
}

export default new Service();
