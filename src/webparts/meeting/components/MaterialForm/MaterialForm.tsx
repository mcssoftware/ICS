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

    constructor(props: Readonly<IMaterialFormProp>) {
        super(props);
        this._billVersions = [];
        this.state = {
            selectedSubTopic: null,
            loadingBillVersion: false,
            // workingDoc: McsUtil.isDefined(props.document) ? { ...props.document } : {} as ISpEventMaterial,
            agenda: props.requireAgendaSelection ? null : (McsUtil.isArray(props.agenda) && props.agenda.length == 1 ? props.agenda[0] : null),
            documentUploadType: business.is_SessionMeeting() ? DocumentUploadType.SessionDocuments : DocumentUploadType.InterimDocument,
            documentId: McsUtil.isDefined(props.document) ? props.document.Id : 0,
            workingDocument: McsUtil.isDefined(props.document) ?
                {
                    agency: props.document.AgencyName,
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
                        disabled={!McsUtil.isDefined(this.state.agenda)}
                        options={this._getSubTopicOptions()}
                    />
                </div>
            </div>
            <div className={styles.row}>
                <div className={styles["col-12"]}>
                    <div className={marginClassName}>
                        <Dropdown
                            label="Document Upload Type"
                            selectedKey={documentUploadType}
                            onChange={this._onUploadTypeSelected}
                            placeholder="Select an option"
                            disabled={documentId > 0}
                            options={[
                                { key: DocumentUploadType.InterimDocument, text: 'Upload Document', disabled: !business.is_SessionMeeting() },
                                { key: DocumentUploadType.LSOBill, text: 'Bills From LMS' },
                                { key: DocumentUploadType.SessionDocuments, text: 'Attached Session Document' },
                            ]}
                        />
                    </div>

                </div>
            </div>
            {documentUploadType == DocumentUploadType.InterimDocument &&
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
                                    value={workingDocument.selectedAgency}
                                    onChange={this._agencySelectChange}
                                    loadOptions={this.loadAgencyOptions} />
                            </div>
                            {/* <TextField label="Providing Agency"
                                name="agency"
                                className={marginClassName}
                                value={workingDocument.agency} onChange={this._onDocTextPropertyChange} /> */}
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles["col-12"]}>
                            <Checkbox label="Include with agenda"
                                name="includeWithAgenda"
                                className={css.combine(marginClassName, styles["mt-3"])}
                                checked={workingDocument.includeWithAgenda} onChange={this._onIncludeWithAgendaChange} />
                        </div>
                    </div>
                    {documentId === 0 && <div className={styles.row}>
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
                    </div>}
                    <div className={css.combine(styles.row, styles["mt-2"])}>
                        <div className={styles["col-6"]}>
                            <DefaultButton text={documentId === 0 ? "Upload Document" : "Update Document"}
                                disabled={!this._canUploadDocument()}
                                onClick={this._uploadFileToSp} />
                        </div>
                    </div>
                </div>
            }
            {documentUploadType == DocumentUploadType.LSOBill &&
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
            {documentUploadType == DocumentUploadType.SessionDocuments &&
                <div>
                    <div className={styles.row}>
                        <div className={styles["col-6"]}>
                            <div className={marginClassName}>
                                <Label>Providing Agency</Label>
                                <AsyncSelect defaultOptions={true}
                                    value={workingDocument.selectedAgency}
                                    onChange={this._agencySelectChange}
                                    loadOptions={this.loadAgencyOptions} />
                            </div>
                        </div>
                        <div className={styles["col-6"]}>
                            <div className={marginClassName}>
                                <Label>Document</Label>
                                <AsyncSelect defaultOptions={true} />
                            </div>
                        </div>
                    </div>
                    <div>
                        <div className={styles["col-6"]}>
                            <div className={css.combine(marginClassName, styles["mt-2"], styles["d-flex"])}>
                                <DefaultButton text="Attach Session Document"
                                    disabled={!McsUtil.isUnsignedInt(workingDocument.sessionDocumentId)}
                                    className={css.combine(styles["mr-2"], styles["bg-primary"], styles["text-white"])}
                                    style={{ maxWidth: '60%' }} />
                                <DefaultButton text="Cancel" className={css.combine(styles["ml-2"], styles["bg-light"], styles["text-dark"])} />
                            </div>
                        </div>
                    </div>
                </div>
            }
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
        this.setState({ documentUploadType: option.key as DocumentUploadType, workingDocument });
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
        if (McsUtil.isString(workingDocument.title) && McsUtil.isString(workingDocument.sortNumber) && McsUtil.isString(workingDocument.agency)) {
            if (this.state.documentId === 0) {
                return this.state.workingDocument.uploadFile !== null;
            }
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
        const agency = McsUtil.isDefined(value) ? value.label : '';
        let { workingDocument } = this.state;
        workingDocument = { ...workingDocument, selectedAgency: value, agency };
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
                                value: a.Title,
                                label: a.AgencyName
                            };
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
        new Promise((resolve) => {
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
            agency: '',
            title: '',
            billVersion: '',
            lsonumber: '',
            sessionDocumentId: 0,
            uploadFile: null,
            selectedAgency: null,
            sortNumber: this.props.sortNumber,
            lsoDocumentType: 'Meeting Attachments',
        };
    }

    private _uploadFileToSp = (): void => {
        const { workingDocument, agenda } = this.state;
        const uploadProperties: ISpEventMaterial = {
            lsoDocumentType: workingDocument.lsoDocumentType,
            AgencyName: workingDocument.selectedAgency,
            Title: workingDocument.title,
            IncludeWithAgenda: workingDocument.includeWithAgenda,
            SortNumber: McsUtil.isUnsignedInt(workingDocument.sortNumber) ? parseInt(workingDocument.sortNumber) : 1,
        };
        const file: File = workingDocument.uploadFile.item[0];
        this.setState({ waitingMessage: "Uploading file" });
        business.upLoad_Document(business.get_FolderNameToUpload(uploadProperties.lsoDocumentType), file.name, uploadProperties, file)
            .then((value: ISpEventMaterial) => {
                this.setState({ waitingMessage: "" });
                this.props.onChange(value, agenda, OperationType.Add);
            });
    }

    private _uploadBillToSp = (): void => {
        const { workingDocument, agenda } = this.state;
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
                    SortNumber: McsUtil.isUnsignedInt(workingDocument.sortNumber) ? parseInt(workingDocument.sortNumber) : 1,
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
                this.props.onChange(value, agenda, OperationType.Add);
            }).catch();
    }
}