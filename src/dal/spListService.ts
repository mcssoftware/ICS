require("@pnp/logging");
require("@pnp/common");
require("@pnp/odata");
import { sp, List, PagedItemCollection, Items, Web, FileAddResult, Item, ItemUpdateResult, FolderAddResult } from "@pnp/sp";
import { IOrderOption } from "./interface";
import { McsUtil } from "../utility/helper";
import { TypedHash } from "@pnp/common";
import { IDocumentItem } from "../interface/spmodal";

export default class SpListService<T> {
    private _isListTitleGuid: boolean;
    private _webUrl: string;

    constructor(private _listTitle: string, private _fromRootWeb: boolean) {
        this._isListTitleGuid = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/i.test(_listTitle);
    }

    public setWebUrl(url: string): void {
        this._webUrl = url;
    }

    public setListTitle(listTitle: string): void {
        this._listTitle = listTitle;
    }

    public getListTitle(): string{
        return this._listTitle;
    }

    public getListItems(filter?: string, select?: string[], expand?: string[], orderBy?: IOrderOption[], skip?: number, top?: number): Promise<T[]> {
        // SB: Disable caching because it does not seem to work properly.
        // if (this.useCaching) {
        //     return this._getRestDataUsingCaching(filter, select, expand, orderBy, skip, top);
        // } else {
        return this._getRestData(filter, select, expand, orderBy, skip, top);
        // }
    }

    public getSelects(): string[] {
        return ["Id", "Title"];
    }

    public getExpands(): string[] {
        return [];
    }

    public getUserSelectForExpand(field: string): string[] {
        return [field + "/Id", field + "/Title", field + "/EMail"];
    }

    public getListItemById(id: number): Promise<T> {
        return new Promise<T>((resolve, reject) => {
            let select: string[] = this.getSelects();
            // tslint:disable-next-line:prefer-const
            let expand: string[] = this.getExpands();
            if (McsUtil.isArray(select)) {
                if (!McsUtil.isArray(select)) {
                    select = [];
                }
                this._getList().items.getById(id)
                    .select(...select)
                    .expand(...expand)
                    .get()
                    .then((value) => {
                        resolve(value);
                    }, (error) => {
                        reject(error);
                    });
            } else {
                this._getList().items.getById(id)
                    .get()
                    .then((value) => {
                        resolve(value);
                    }, (error) => {
                        reject(error);
                    });
            }
        });
    }

    public addNewItem(properties: TypedHash<any>): Promise<T> {
        return new Promise<T>((resolve, reject) => {
            this._getList().items.add(properties).then((result) => {
                resolve(result.data);
            }, (error) => {
                reject(error);
            });
        });
    }

    public updateItem(id: number, listItemEntityTypeFullName: string, properties: TypedHash<any>): Promise<boolean> {
        return new Promise((resolve, reject) => {
            this._getList().items.getById(id).update(properties, "*", listItemEntityTypeFullName).then((result) => {
                resolve();
            }, (error) => {
                reject(error);
            });
        });
    }

    public deleteItem(id: number): Promise<void> {
        return this._getList().items.getById(id).delete();
    }

    public loadList(): Promise<any> {
        return this._getList().expand(...['RootFolder']).get();
    }

    public getFolder(sourceFolderUr: string, name: string): Promise<any[]> {
        return this._getWeb().getFolderByServerRelativePath(sourceFolderUr).folders
            .filter(`Name eq '${name}'`)
            .select('ServerRelativeUrl')
            .get();
    }

    public createFolder(serverRelativeUrl: string, foldername: string): Promise<FolderAddResult> {
        return new Promise((resolve, reject) => {
            this._getWeb()
                .getFolderByServerRelativeUrl(serverRelativeUrl)
                .folders
                .add(foldername)
                .then((value) => { resolve(value.data); })
                .catch(e => reject(e));
        });
    }

    public addOrUpdateDocument(folderServerRelativeUrl: string, fileName: string, propertiesToUpdate: IDocumentItem, blob: Blob): Promise<IDocumentItem> {
        return new Promise<IDocumentItem>((resolve, reject) => {
            this._getWeb().getFolderByServerRelativePath(folderServerRelativeUrl).files.add(fileName, blob, true)
                .then((fileAdded: FileAddResult) => {
                    fileAdded.file.getItem().then((item: Item) => {
                        item.update(propertiesToUpdate).then((value: ItemUpdateResult) => {
                            value.item.select(...this.getSelects()).expand(...this.getExpands()).get().then((documentItem: IDocumentItem) => {
                                resolve(documentItem);
                            }, (err) => { reject(err); });
                        }, (err) => { reject(err); });
                    }, (err) => { reject(err); });
                }, (err) => {
                    reject(err);
                });
        });
    }

    private _getRestData(filter?: string, select?: string[], expand?: string[], orderBy?: IOrderOption[], skip?: number, itemCount?: number): Promise<T[]> {
        return new Promise<T[]>((resolve, reject) => {
            let result: T[] = [];
            if (!McsUtil.isString(filter)) {
                filter = "";
            }
            let listSelect: string[] = select;
            if (!McsUtil.isArray(select)) {
                listSelect = this.getSelects();
            }
            let listExpand: string[] = expand;
            if (!McsUtil.isArray(expand)) {
                listExpand = this.getExpands();
            }
            const getAllItems: boolean = typeof itemCount === "undefined" || itemCount === null;
            itemCount = itemCount || 50;
            skip = skip || 0;
            // tslint:disable-next-line:typedef
            let listItems: Items = this._getList().items
                .filter(filter)
                .select(...listSelect)
                .expand(...listExpand);
            if (McsUtil.isArray(orderBy)) {
                orderBy.forEach((ele: IOrderOption) => {
                    listItems = listItems.orderBy(ele.Field, ele.IsAscending);
                });
            } else {
                listItems = listItems.orderBy("Id");
            }
            if (McsUtil.isArray(listSelect)) {
                if (listSelect.indexOf("Id") < 0) {
                    listSelect.push("Id");
                }
                if (!McsUtil.isArray(listExpand)) {
                    listExpand = [];
                }
                listItems.select(...listSelect)
                    .expand(...listExpand)
                    .skip(skip)
                    .top(itemCount).getPaged().then((value) => {
                        result = value.results as T[];
                        if (getAllItems && value.hasNext) {
                            this._getNextPages(value, []).then((pagedResult) => {
                                result = result.concat(pagedResult);
                                resolve(result);
                            });
                        } else {
                            resolve(result);
                        }
                    }, (error) => {
                        reject(error);
                    });
            } else {
                listItems.select(...listSelect)
                    .expand(...listExpand)
                    .skip(skip)
                    .top(itemCount).getPaged().then((value) => {
                        result = value.results as T[];
                        if (getAllItems && value.hasNext) {
                            this._getNextPages(value, []).then((pagedResult) => {
                                result = result.concat(pagedResult);
                                resolve(result);
                            });
                        } else {
                            resolve(result);
                        }
                    }, (error) => {
                        reject(error);
                    });
            }
        });
    }

    private _getWeb(): Web {
        if (McsUtil.isDefined(this._webUrl)) {
            return new Web(this._webUrl);
        } else {
            if (this._fromRootWeb) {
                return sp.site.rootWeb;
            }
            return sp.web;
        }
    }

    private _getList(): List {
        if (this._isListTitleGuid) {
            return this._getWeb().lists.getById(this._listTitle);
        }
        return this._getWeb().lists.getByTitle(this._listTitle);
    }

    private _getNextPages(paged: PagedItemCollection<any>, items: T[]): Promise<T[]> {
        if (paged.hasNext) {
            return paged.getNext().then((value) => {
                return this._getNextPages(value, items.concat(value.results));
            });
        }
        return Promise.resolve(items);
        // return new Promise<T[]>((resolve, reject) => {
        //     paged.getNext().then((value) => {
        //         let result: T[] = value.results;
        //         if (value.hasNext) {
        //             return
        //             this._getNextPages(value).then((nextPage) => {
        //                 result = result.concat(nextPage);
        //                 resolve(result);
        //             });
        //         } else {
        //             resolve(result);
        //         }
        //     }, (error) => {
        //         resolve([]);
        //     });
        // });
    }
}

