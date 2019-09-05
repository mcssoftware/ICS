export interface IOrderOption {
    Field: string;
    IsAscending: boolean;
}

export interface IFolderCreation {
    name: string;
    SubFolder?: IFolderCreation[];
    ServerRelativeUrl?: string;
}