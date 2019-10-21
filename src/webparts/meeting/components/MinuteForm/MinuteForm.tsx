import * as React from 'react';
import { business } from '../../../../business';
import { ISpEvent, ISpEventMaterial } from "../../../../interface/spmodal";
import styles from '../Meeting.module.scss';
import { McsUtil } from '../../../../utility/helper';
import { PrimaryButton } from 'office-ui-fabric-react';
import { Waiting } from '../../../../controls/waiting';
import { sp } from '@pnp/sp';
import IcsAppConstants from '../../../../configuration';
import css from '../../../../utility/css';

export interface IMinuteFormProps {
    event: ISpEvent;
}

export interface IMinuteFormState {
    minute: ISpEventMaterial;
    waitingMessage: string;
    folderToInsert: string;
}

const minuteDocumentType = "Meeting Minutes";

export class MinuteForm extends React.Component<IMinuteFormProps, IMinuteFormState> {
    constructor(props: IMinuteFormProps) {
        super(props);

        this.state = {
            minute: business.get_DocumentByType(minuteDocumentType),
            waitingMessage: '',
            folderToInsert: ''
        };

        business.get_folderServerRelativeUrl(business.get_FolderNameToUpload(minuteDocumentType))
            .then((value) => {
                this.setState({ folderToInsert: value });
            });
    }

    public render() {
        const { minute, folderToInsert } = this.state;
        return (
            <div>
                {McsUtil.isDefined(minute) && McsUtil.isDefined(minute.File) && <div className={css.combine(styles.row, styles["m-2"])}>
                    <div className={styles["col-4"]}>
                        <PrimaryButton title="Edit Minute" onClick={this._openInWordApp}
                            style={{ color: "#fff", textDecoration: "none" }}>Edit Minute</PrimaryButton>
                    </div>
                    <div className={styles["col-4"]}>
                        <PrimaryButton text="Approve meeting minutes for publishing" disabled={minute.File.CheckOutType === 2} onClick={this._approveMinuteClicked} />
                    </div>
                </div>}
                <div className={css.combine(styles.row, styles["m-2"])}>
                    <div className={styles["col-4"]}>
                        <PrimaryButton text={McsUtil.isDefined(minute) ? "Update Material Index in Minutes" : "Generate Template"}
                            onClick={this._createMinuteClicked} />
                    </div>
                    <div className={styles["col-4"]}>
                        <PrimaryButton href={folderToInsert} target="_blank" title="Open Document Folders"
                            style={{ color: "#fff", textDecoration: "none" }}
                            disabled={!McsUtil.isString(folderToInsert)} >Open Document Folders</PrimaryButton>
                    </div>
                    <div className={styles["col-4"]}></div>
                </div>
                <Waiting message={this.state.waitingMessage} />
            </div>
        );
    }

    private _openInWordApp = (): void => {
        const { minute } = this.state;
        const serverRelativeUrl: string = minute.File.ServerRelativeUrl.split("?")[0];
        _WriteDocEngagement("DocLibECB_Click_ID_EditIn_Word", "OneDrive_DocLibECB_Click_ID_EditIn_Word");
        editDocumentWithProgID2(serverRelativeUrl, "", "SharePoint.OpenDocuments", "0", business.getWebUrl(), "0", "ms-word");
    }

    private _approveMinuteClicked = (): void => {
        this.setState({ waitingMessage: 'Publishing minute' });
        business.publishDocument(this.state.minute)
            .then(() => {
                this.setState({ minute: business.get_DocumentByType("Meeting Minutes"), waitingMessage: '' });
            })
            .catch((e) => { });
    }

    private _createMinuteClicked = (): void => {
        const { minute } = this.state;
        this.setState({ waitingMessage: McsUtil.isDefined(minute) ? 'Updating minute' : 'Creating minute' });
        if (McsUtil.isDefined(minute)) {
            sp.web.getFileByServerRelativePath(minute.File.ServerRelativeUrl).getBlob()
                .then((blob) => {
                    this._createOrUpdateMinute([blob]);
                }).catch((e) => { });
        } else {
            this._createOrUpdateMinute([]);
        }

    }

    private _createOrUpdateMinute = (fileBlob: Blob[]): void => {
        const { minute } = this.state;
        const formData = new FormData();
        const modalData = JSON.stringify(business.get_publishingMeeting());
        formData.append("model", modalData);
        if (McsUtil.isArray(fileBlob) && fileBlob.length > 0) {
            fileBlob.forEach((b, i) => {
                formData.append("file" + i, b);
            });
        }
        business.generateMeetingDocument(IcsAppConstants.getCreateMinutePreviewPartial(), "multipart", formData)
            .then((blob) => {
                var materialMetaData = {
                    Title: "Meeting Minutes",
                    AgencyName: "LSO",
                    lsoDocumentType: "Meeting Minutes",
                    IncludeWithAgenda: false,
                    SortNumber: McsUtil.isDefined(minute) ? minute.SortNumber : 999, //to ensure this is always last document.
                };
                return business.upLoad_Document(business.get_FolderNameToUpload(minuteDocumentType), "Meeting Minutes.docx", materialMetaData, blob);
            }).then((document) => {
                const event = business.get_Event();
                const eventPropToUpdate = {};
                const eventLookupField = business.get_EventDocumentLookupField();
                const docIds: number[] = event[eventLookupField] || [];
                if (docIds.indexOf(document.Id) < 0) {
                    docIds.push(document.Id);
                    eventPropToUpdate[eventLookupField] = {
                        __metadata: {
                            type: "Collection(Edm.Int32)"
                        },
                        results: [...docIds]
                    };
                    return business.edit_Event(event.Id, event["odata.type"], eventPropToUpdate);
                } else {
                    return Promise.resolve(event);
                }
            }).then(() => {
                this.setState({ minute: business.get_DocumentByType(minuteDocumentType), waitingMessage: '' });
            }).catch();
    }
}

declare function _WriteDocEngagement(a: string, b: string): void;
declare function editDocumentWithProgID2(a: string, b: string, c: string, d: string, e: string, f: string, g: string): void;