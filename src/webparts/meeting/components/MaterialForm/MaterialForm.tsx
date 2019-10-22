import * as React from 'react';
import { IMaterialFormProp, IMaterialFormState, DocumentUploadType } from './IMaterialForm';
import { TextField, Dropdown, IDropdownOption, Checkbox, DefaultButton, Label } from "office-ui-fabric-react";
import { Waiting } from '../../../../controls/waiting';
import styles from '../Meeting.module.scss';
import css from '../../../../utility/css';
import { McsUtil } from '../../../../utility/helper';
import AsyncSelect from 'react-select/async';
import Select from 'react-select';
import { cloneDeep } from '@microsoft/sp-lodash-subset';
import { business } from '../../../../business';
import { ISpEventMaterial, OperationType, IBillVersion } from '../../../../interface/spmodal';

export default class MaterialForm extends React.Component<IMaterialFormProp, IMaterialFormState> {

    private _billVersions: IBillVersion[];
    private _enableDocumentAttachment: boolean;

    constructor(props: Readonly<IMaterialFormProp>) {
        super(props);
        this._billVersions = [];
        this._enableDocumentAttachment = business.can_CreateBudgetMeeting();
        const documentUploadType = this._enableDocumentAttachment ? DocumentUploadType.SessionDocuments : DocumentUploadType.InterimDocument;
        this.state = {
            selectedSubTopic: null,
            loadingBillVersion: false,
            // workingDoc: McsUtil.isDefined(props.document) ? { ...props.document } : {} as ISpEventMaterial,
            agenda: props.requireAgendaSelection ? null : (McsUtil.isArray(props.agenda) && props.agenda.length == 1 ? props.agenda[0] : null),
            documentUploadType,
            documentId: McsUtil.isDefined(props.document) ? props.document.Id : 0,
            defaultAsyncDocuments: [],
            workingDocument: McsUtil.isDefined(props.document) ?
                {
                    selectedAgency: {
                        label: props.document.AgencyName,
                        value: props.document.AgencyName
                    },
                    includeWithAgenda: props.document.IncludeWithAgenda,
                    title: props.document.Title,
                    billVersion: '',
                    lsonumber: '',
                    sessionDocumentId: 0,
                    uploadFile: null,
                    sortNumber: props.document.SortNumber,
                    lsoDocumentType: props.document.lsoDocumentType,
                } : this._getFormDefaultValue()
        };
    }

    public componentDidMount(): void {
        this._loadDefaultDocumentOptions(this.state.documentUploadType);
    }

    public render(): React.ReactElement<IMaterialFormProp> {
        const { selectedSubTopic, agenda, documentUploadType, workingDocument, documentId, waitingMessage } = this.state;
        const marginClassName = css.combine(styles["ml-2"], styles["mr-2"]);
        const selectTopic = this.props.requireAgendaSelection && !McsUtil.isDefined(this.state.agenda);
        return (<div className={styles["container-fluid"]}>
            <div className={styles.row}>
                <div className={styles["col-2"]}>
                    <TextField label="Order #"
                        name="sortNumber"
                        className={marginClassName}
                        onGetErrorMessage={(value) => /\d+/.test(value) ? undefined : 'Must be numberic'}
                        value={workingDocument.sortNumber}
                        onChange={this._onTextFieldChanged}
                    />
                </div>
                {selectTopic && <div className={styles["col-5"]}>
                    <Dropdown
                        label="Topic"
                        className={marginClassName}
                        selectedKey={McsUtil.isDefined(agenda) ? agenda.Id : 0}
                        onChange={this._onTopicSelected}
                        options={this._getTopicOptions()}
                    />
                </div>}
                <div className={selectTopic ? styles["col-5"] : styles["col-10"]}>
                    <Dropdown
                        label="SubTopic"
                        className={marginClassName}
                        selectedKey={McsUtil.isDefined(selectedSubTopic) ? selectedSubTopic.Id.toString() : "0"}
                        onChange={this._onSubTopicSelected}
                        placeholder="Select an option"
                        disabled={!McsUtil.isDefined(this.state.agenda) || (documentId !== 0)}
                        options={this._getSubTopicOptions()}
                    />
                </div>
            </div>
            {documentId == 0 && <div className={styles.row}>
                <div className={styles["col-12"]}>
                    <div className={marginClassName}>
                        <Dropdown
                            label="Document Upload Type"
                            selectedKey={documentUploadType}
                            onChange={this._onUploadTypeSelected}
                            placeholder="Select an option"
                            options={[
                                { key: DocumentUploadType.InterimDocument, text: 'Upload Document', disabled: this._enableDocumentAttachment },
                                { key: DocumentUploadType.LSOBill, text: 'Bills From LMS' },
                                { key: DocumentUploadType.SessionDocuments, text: 'Attached Session Document', disabled: !this._enableDocumentAttachment },
                            ]}
                        />
                    </div>
                </div>
            </div>}
            {documentId == 0 && documentUploadType == DocumentUploadType.InterimDocument &&
                <div>
                    <div className={styles.row}>
                        <div className={styles["col-6"]}>
                            <TextField label="Attachment Title"
                                name="title"
                                className={marginClassName}
                                value={workingDocument.title} onChange={this._onDocTextPropertyChange} />
                        </div>
                        <div className={styles["col-6"]}>
                            <div className={marginClassName}>
                                <Label>Providing Agency</Label>
                                <AsyncSelect defaultOptions={true}
                                    isClearable={true}
                                    value={workingDocument.selectedAgency}
                                    onChange={this._agencySelectChange}
                                    loadOptions={this.loadAgencyOptions} />
                            </div>
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles["col-12"]}>
                            <Checkbox label="Include with agenda"
                                name="includeWithAgenda"
                                className={css.combine(marginClassName, styles["mt-3"])}
                                checked={workingDocument.includeWithAgenda}
                                onChange={this._onIncludeWithAgendaChange} />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles["col-12"]}>
                            <TextField label="File"
                                type="File"
                                className={marginClassName}
                                multiple={false}
                                id="fileUpload"
                                onChange={this._onFileChange}
                                placeholder="Select file to upload ..."
                            />
                        </div>
                    </div>
                    <div className={css.combine(styles.row, styles["mt-2"])}>
                        <div className={styles["col-6"]}>
                            <DefaultButton text={"Upload Document"}
                                disabled={!this._canUploadDocument()}
                                className={css.combine(styles["mr-2"], styles["bg-primary"], styles["text-white"])}
                                onClick={this._uploadFileToSp} />
                        </div>
                    </div>
                </div>
            }
            {documentId == 0 && documentUploadType == DocumentUploadType.LSOBill &&
                <div>
                    <div className={styles.row}>
                        <div className={styles["col-6"]}>
                            <Label>Bill</Label>
                            <AsyncSelect defaultOptions={true}
                                isClearable={true}
                                value={workingDocument.selectedBill}
                                onChange={this._billSelectChange}
                                loadOptions={this.loadBillOptions} />
                        </div>
                        <div className={styles["col-6"]}>
                            <Label>Bill Version</Label>
                            <Select
                                value={workingDocument.selectedBillVersion}
                                isLoading={this.state.loadingBillVersion}
                                isSearchable={true}
                                isClearable={true}
                                name="billversion"
                                onChange={this._billVersionSelectChange}
                                options={this._billVersions.map(a => { return { value: a.VersionLabel, label: `${a.DocumentStatus} (${a.DocumentVersion})` }; })}
                            />
                        </div>
                    </div>
                    <div className={css.combine(styles.row, styles["mt-2"])}>
                        <div className={styles["col-6"]}>
                            <DefaultButton text="Get Bill from LMS"
                                className={css.combine(styles["mr-2"], styles["bg-primary"], styles["text-white"])}
                                disabled={!McsUtil.isString(workingDocument.lsonumber) || !McsUtil.isDefined(workingDocument.billVersion)}
                                onClick={this._uploadBillToSp} />
                        </div>
                    </div>
                </div>
            }
            {documentId == 0 && documentUploadType == DocumentUploadType.SessionDocuments &&
                <div>
                <div className={styles.row}>
                    <div className={styles["col-6"]}>
                        <div className={marginClassName}>
                            <Label>Providing Agency</Label>
                            <AsyncSelect defaultOptions={true}
                                isClearable={true}
                                value={workingDocument.searchSelectedAgency}
                                onChange={this._searchAgencySelectChange}
                                loadOptions={this.loadAgencyOptionsForSearch} />
                        </div>
                    </div>
                    <div className={styles["col-6"]}>
                        <div className={marginClassName}>
                            <Label>Document</Label>
                            <AsyncSelect 
                                isClearable={true}
                                value={workingDocument.searchSelectedDocument}
                                onChange={this._searchDocumentSelectChange}
                                defaultOptions={this.state.defaultAsyncDocuments}
                                loadOptions={this.loadDocumentOptions} />
                        </div>
                    </div>
                </div>
                <div>
                    <div className={styles["col-6"]}>
                        <div className={css.combine(marginClassName, styles["mt-2"], styles["d-flex"])}>
                            <DefaultButton text="Attach Session Document"
                                disabled={!this._canAttachDocument()}
                                className={css.combine(styles["mr-2"], styles["bg-primary"], styles["text-white"])}
                                style={{ maxWidth: '60%' }}
                                onClick={this._attachMaterial} />
                        </div>
                    </div>
                </div>
            </div>
            }
            {documentId != 0 && <div>
                <div className={styles.row}>
                    <div className={styles["col-6"]}>
                        <TextField label="Attachment Title"
                            name="title"
                            className={marginClassName}
                            value={workingDocument.title}
                            onChange={this._onDocTextPropertyChange} />
                    </div>
                    <div className={styles["col-6"]}>
                        <div className={marginClassName}>
                            <Label>Providing Agency</Label>
                            <AsyncSelect defaultOptions={true}
                                isClearable={true}
                                value={workingDocument.selectedAgency}
                                onChange={this._agencySelectChange}
                                loadOptions={this.loadAgencyOptions} />
                        </div>
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles["col-12"]}>
                        <Checkbox label="Include with agenda"
                            name="includeWithAgenda"
                            className={css.combine(marginClassName, styles["mt-3"])}
                            checked={workingDocument.includeWithAgenda}
                            onChange={this._onIncludeWithAgendaChange} />
                    </div>
                </div>
                <div className={css.combine(styles.row, styles["mt-2"])}>
                    <div className={styles["col-6"]}>
                        <div className={css.combine(styles["d-flex"])}>
                            <DefaultButton text={"Update Document"}
                                disabled={!this._canUpdateDocument()}
                                className={css.combine(styles["mr-2"], styles["bg-primary"], styles["text-white"])}
                                onClick={this._updateMaterial} />
                            <DefaultButton text={"Delete Document"}
                                className={css.combine(styles["mr-2"], styles["bg-primary"], styles["text-white"])}
                                onClick={this._deleteMaterial} />
                        </div>
                    </div>
                </div>
            </div>}
            <Waiting message={waitingMessage} />
        </div>);
    }

    private _onTextFieldChanged = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value: string) => {
        const { workingDocument } = this.state;
        workingDocument[(event.target as HTMLInputElement).name] = value;
        this.setState({ workingDocument });
    }

    private _onTopicSelected = (event: any, option?: IDropdownOption) => {
        if (option.key == "0") {
            this.setState({ agenda: null });
        } else {
            this.setState({ agenda: this.props.agenda.filter(a => a.Id == (option.key as number))[0] });
        }
    }

    private _onSubTopicSelected = (event: any, option?: IDropdownOption) => {
        if (option.key == "0") {
            this.setState({ selectedSubTopic: null });
        } else {
            const selectedSubTopic = this.state.agenda.SubTopics.filter(a => a.Id == (option.key as number))[0];
            this.setState({ selectedSubTopic });
        }
    }

    private _onUploadTypeSelected = (event: any, option?: IDropdownOption) => {
        const workingDocument = this._getFormDefaultValue();
        const documentUploadType = option.key as DocumentUploadType;
        this.setState({ documentUploadType, workingDocument });
        this._loadDefaultDocumentOptions(documentUploadType);
    }

    private _onDocTextPropertyChange = (event: React.FormEvent<HTMLInputElement>, newValue?: string): void => {
        const workingDoc = cloneDeep(this.state.workingDocument);
        workingDoc[(event.target as HTMLInputElement).name] = newValue;
        this.setState({ workingDocument: workingDoc });
    }

    private _onIncludeWithAgendaChange = (ev?: any, checked?: boolean): void => {
        const workingDocument = cloneDeep(this.state.workingDocument);
        workingDocument.includeWithAgenda = checked || false;
        this.setState({ workingDocument });
    }

    private _onFileChange = (event: React.FormEvent<HTMLInputElement>): void => {
        const fileList: FileList = (event.target as HTMLInputElement).files;
        const workingDocument = cloneDeep(this.state.workingDocument);
        workingDocument.uploadFile = fileList;
        this.setState({ workingDocument });
    }

    private _canUploadDocument = (): boolean => {
        const { workingDocument } = this.state;
        if (McsUtil.isString(workingDocument.title) && McsUtil.isUnsignedInt(workingDocument.sortNumber) &&
            McsUtil.isDefined(workingDocument.selectedAgency) && McsUtil.isString(workingDocument.selectedAgency.value)) {
            if (this.state.documentId === 0) {
                return this.state.workingDocument.uploadFile !== null;
            }
            return true;
        }
        return false;
    }

    private _canAttachDocument = (): boolean => {
        const { workingDocument } = this.state;
        const document: ISpEventMaterial = (workingDocument.searchSelectedDocument || {}).item;
        if (McsUtil.isDefined(document)) {
            if (McsUtil.isString(document.AgencyName) || (McsUtil.isDefined(workingDocument.searchSelectedAgency) && McsUtil.isString(workingDocument.searchSelectedAgency.label))) {
                return McsUtil.isUnsignedInt(workingDocument.sortNumber);
            }
        }
        return false;
    }

    private _canUpdateDocument = (): boolean => {
        const { workingDocument, documentId } = this.state;
        if (McsUtil.isUnsignedInt(documentId) && McsUtil.isUnsignedInt(workingDocument.sortNumber) && McsUtil.isString(workingDocument.title) &&
            McsUtil.isDefined(workingDocument.selectedAgency) && McsUtil.isString(workingDocument.selectedAgency.value)) {
            return true;
        }
        return false;
    }

    private _getTopicOptions = (): IDropdownOption[] => {
        const options: IDropdownOption[] = [{
            key: '0',
            text: 'Select SubTopic'
        }];
        if (McsUtil.isArray(this.props.agenda)) {
            this.props.agenda.forEach((a) => {
                options.push({
                    key: a.Id.toString(),
                    text: a.AgendaTitle
                });
            });
        }
        return options;
    }

    private _agencySelectChange = (value): void => {
        let { workingDocument } = this.state;
        workingDocument = { ...workingDocument, selectedAgency: value };
        this.setState({ workingDocument });
    }

    private _searchAgencySelectChange = (value): void => {
        let { workingDocument } = this.state;
        workingDocument = { ...workingDocument, searchSelectedAgency: value };
        this.setState({ workingDocument });
        this._loadDefaultDocumentOptions(this.state.documentUploadType);
    }

    private _searchDocumentSelectChange = (value): void => {
        let { workingDocument } = this.state;
        workingDocument = { ...workingDocument, searchSelectedDocument: value };
        this.setState({ workingDocument });
    }

    private _billSelectChange = (value): void => {
        const lsonumber = McsUtil.isDefined(value) ? value.label : '';
        let { workingDocument } = this.state;
        workingDocument = { ...workingDocument, lsonumber, selectedBill: value, selectedBillVersion: undefined, billVersion: undefined };
        this.setState({ workingDocument, loadingBillVersion: true });
        this._billVersions = [];
        this.loadBillVersionsOptions(value);
    }

    private _billVersionSelectChange = (value): void => {
        const billVersion = McsUtil.isDefined(value) ? value.value : undefined;
        let { workingDocument } = this.state;
        workingDocument = { ...workingDocument, selectedBillVersion: value, billVersion };
        this.setState({ workingDocument });
    }

    private loadAgencyOptions = (inputValue) =>
        new Promise((resolve) => {
            business.find_Agency(inputValue)
                .then((val) => {
                    resolve(
                        val.map(a => {
                            return {
                                value: a.AgencyName === "." ? "LSO" : a.AgencyName,
                                label: a.AgencyName === "." ? "LSO" : a.AgencyName
                            };
                        }));
                });
        })

    private loadAgencyOptionsForSearch = (inputValue) =>
        new Promise((resolve) => {
            business.find_Agency(inputValue)
                .then((val) => {
                    resolve(
                        val.map(a => {
                            if (a.AgencyName === ".") {
                                return {
                                    value: "LSO",
                                    label: "LSO"
                                };
                            } else {
                                return {
                                    value: a.Title,
                                    label: a.AgencyName
                                };
                            }
                        }));
                });
        })

    private loadBillOptions = (inputValue) =>
        new Promise((resolve) => {
            business.find_Bill(inputValue)
                .then((val) => {
                    resolve(
                        val.map(a => {
                            return {
                                value: a.Id,
                                label: a.LSONumber,
                                item: a
                            };
                        }));
                });
        })

    private loadBillVersionsOptions = (selectedBill: any) =>
        new Promise(() => {
            if (McsUtil.isDefined(selectedBill) && selectedBill.value > 0) {
                business.find_BillItemVersion(selectedBill.value)
                    .then((val) => {
                        this._billVersions = val;
                        this.setState({ loadingBillVersion: false });
                    });
            } else {
                this._billVersions = [];
                this.setState({ loadingBillVersion: false });
            }
        })

    private _loadDefaultDocumentOptions = (documentType: DocumentUploadType) => {
        if (documentType === DocumentUploadType.SessionDocuments) {
            this.loadDocumentOptions('').then((result: any[]) => {
                this.setState({ defaultAsyncDocuments: result });
            });
        } else {
            if (this.state.defaultAsyncDocuments.length > 0) {
                this.setState({ defaultAsyncDocuments: [] });
            }
        }
    }

    private loadDocumentOptions = (inputValue) =>
        new Promise((resolve) => {
            let agencyname = McsUtil.isDefined(this.state.workingDocument) && McsUtil.isDefined(this.state.workingDocument.searchSelectedAgency) &&
                McsUtil.isDefined(this.state.workingDocument.searchSelectedAgency.value) ? this.state.workingDocument.searchSelectedAgency.label : undefined;
            let agencynumber = McsUtil.isDefined(this.state.workingDocument) && McsUtil.isDefined(this.state.workingDocument.searchSelectedAgency) &&
                McsUtil.isDefined(this.state.workingDocument.searchSelectedAgency.value) ? this.state.workingDocument.searchSelectedAgency.value : undefined;
            business.find_Document(agencyname, agencynumber, inputValue)
                .then((val) => {
                    resolve(
                        val.map(a => {
                            return {
                                value: a.Id,
                                label: a.FileLeafRef,
                                item: a
                            };
                        }));
                });
        })

    private _getSubTopicOptions = (): IDropdownOption[] => {
        const options: IDropdownOption[] = [{
            key: '0',
            text: 'Select SubTopic'
        }];
        if (McsUtil.isDefined(this.state.agenda) && McsUtil.isArray(this.state.agenda.SubTopics)) {
            this.state.agenda.SubTopics.forEach((a) => {
                options.push({
                    key: a.Id.toString(),
                    text: a.AgendaTitle
                });
            });
        }
        return options;
    }

    private _getFormDefaultValue = (): any => {
        return {
            title: '',
            billVersion: '',
            lsonumber: '',
            sessionDocumentId: 0,
            uploadFile: null,
            selectedAgency: null,
            sortNumber: this.props.sortNumber,
            includeWithAgenda: false,
            lsoDocumentType: 'Meeting Attachments',
        };
    }

    private _uploadFileToSp = (): void => {
        const { workingDocument, agenda, selectedSubTopic } = this.state;
        const uploadProperties: ISpEventMaterial = {
            lsoDocumentType: workingDocument.lsoDocumentType,
            AgencyName: workingDocument.selectedAgency.label,
            Title: workingDocument.title,
            IncludeWithAgenda: workingDocument.includeWithAgenda,
            SortNumber: parseInt(workingDocument.sortNumber),
        };
        const file: File = workingDocument.uploadFile[0];
        this.setState({ waitingMessage: "Uploading file" });
        business.upLoad_Document(business.get_FolderNameToUpload(uploadProperties.lsoDocumentType), file.name, uploadProperties, file)
            .then((value: ISpEventMaterial) => {
                this.setState({ waitingMessage: "" });
                this.props.onChange(value, McsUtil.isDefined(selectedSubTopic) ? selectedSubTopic : agenda, OperationType.Add);
            });
    }

    private _uploadBillToSp = (): void => {
        const { workingDocument, agenda, selectedSubTopic } = this.state;
        const selectedVersion = this._billVersions.filter(a => a.VersionLabel == workingDocument.billVersion);
        const isCurrentVersion = selectedVersion[0].IsCurrentVersion;
        const selectedBill: any = workingDocument.selectedBill.item;
        this.setState({ waitingMessage: "" });
        business.getDocumentFromIntranet(selectedBill.File.ServerRelativeUrl, isCurrentVersion ? undefined : workingDocument.billVersion)
            .then((value: Blob) => {
                const uploadProperties: ISpEventMaterial = {
                    lsoDocumentType: "Meeting Attachments",
                    AgencyName: "LSO",
                    Title: selectedBill.LSONumber + " v" + selectedVersion[0].DocumentVersion + " " + selectedVersion[0].CatchTitle,
                    IncludeWithAgenda: workingDocument.includeWithAgenda,
                    SortNumber: parseInt(workingDocument.sortNumber),
                };
                var filename = selectedVersion[0].FileLeafRef; // uriparts[uriparts.length - 1];
                var lastindex = filename.lastIndexOf(".");
                var extension = "docx"; //"pdf";
                if (lastindex > 0) {
                    //extension = filename.substring(lastindex + 1);
                    filename = filename.substring(0, lastindex);
                }
                var fileNameToUpload: string = filename + " v" + selectedVersion[0].DocumentVersion + "." + extension;
                return business.upLoad_Document(business.get_FolderNameToUpload(uploadProperties.lsoDocumentType), fileNameToUpload, uploadProperties, value);
            }).then((value: ISpEventMaterial) => {
                this.setState({ waitingMessage: "" });
                this.props.onChange(value, McsUtil.isDefined(selectedSubTopic) ? selectedSubTopic : agenda, OperationType.Add);
            }).catch();
    }

    private _updateMaterial = (): void => {
        const { document } = this.props;
        const { workingDocument, agenda, selectedSubTopic } = this.state;
        business.edit_Document(document.Id, document["odata.type"], {
            AgencyName: workingDocument.selectedAgency.label,
            Title: workingDocument.title,
            IncludeWithAgenda: workingDocument.includeWithAgenda,
            SortNumber: parseInt(workingDocument.sortNumber),
        }).then((value: ISpEventMaterial) => {
            this.setState({ waitingMessage: "" });
            this.props.onChange(value, McsUtil.isDefined(selectedSubTopic) ? selectedSubTopic : agenda, OperationType.Edit);
        });
    }

    private _deleteMaterial = (): void => {
        const { document } = this.props;
        const { agenda } = this.state;
        business.delete_Document(document.Id);
        this.props.onChange(document, agenda, OperationType.Delete);
    }

    private _attachMaterial = (): void => {
        const { workingDocument, agenda, selectedSubTopic } = this.state;
        const document: ISpEventMaterial = workingDocument.searchSelectedDocument.item;
        business.edit_Document(document.Id, document["odata.type"], {
            AgencyName: McsUtil.isString(document.AgencyName) ? document.AgencyName : workingDocument.searchSelectedAgency.label,
            Title: McsUtil.isString(document.Title) ? document.Title : document.File.Name,
            IncludeWithAgenda: workingDocument.includeWithAgenda,
            SortNumber: parseInt(workingDocument.sortNumber),
        }).then((value: ISpEventMaterial) => {
            this.setState({ waitingMessage: "" });
            this.props.onChange(value, McsUtil.isDefined(selectedSubTopic) ? selectedSubTopic : agenda, OperationType.Add);
        });
    }
}