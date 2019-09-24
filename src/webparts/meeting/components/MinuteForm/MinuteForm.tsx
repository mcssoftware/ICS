import * as React from 'react';
import { business } from '../../../../business';
import { ISpEvent, ISpEventMaterial } from "../../../../interface/spmodal";
import styles from '../Meeting.module.scss';
import { McsUtil } from '../../../../utility/helper';
import { Link, PrimaryButton } from 'office-ui-fabric-react';
import { Waiting } from '../../../../controls/waiting';
import { sp } from '@pnp/sp';
import IcsAppConstants from '../../../../configuration';

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
                {McsUtil.isDefined(minute) && <div className={styles.row}>
                    <div className={styles["col-4"]}>
                        <Link href={minute.File.LinkingUrl}>Edit Minute</Link>
                    </div>
                    <div className={styles["col-4"]}>
                        <PrimaryButton text="Primary" disabled={minute.File.CheckOutType === 2} onClick={this._approveMinuteClicked} />
                    </div>
                </div>}
                <div className={styles.row}>
                    <div className={styles["col-4"]}>
                        <PrimaryButton text={McsUtil.isDefined(minute) ? "Update Material Index in Minutes" : "Generate Template"}
                            onClick={this._createMinuteClicked} />
                    </div>
                    <div className={styles["col-4"]}>
                        <PrimaryButton href={folderToInsert} target="_blank" title="Open Document Folders"
                            disabled={!McsUtil.isString(folderToInsert)} >Open Document Folders</PrimaryButton>
                    </div>
                    <div className={styles["col-4"]}></div>
                </div>
                <Waiting message={this.state.waitingMessage} />
            </div>
        );
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
        formData.append("modal", JSON.stringify(business.get_publishingMeeting()));
        if (McsUtil.isArray(fileBlob) && fileBlob.length > 0) {
            fileBlob.forEach((b, i) => {
                formData.append("file" + i, b);
            });
        }
        business.generateMeetingDocument(IcsAppConstants.getCreateMinutePreviewPartial(), "multipart/form-data", formData)
            .then((blob) => {
                var materialMetaData = {
                    Title: "Meeting Minutes",
                    AgencyName: "LSO",
                    lsoDocumentType: "Meeting Minutes",
                    IncludeWithAgenda: false,
                    SortNumber: McsUtil.isDefined(minute) ? minute.SortNumber : 999, //to ensure this is always last document.
                };
                return business.upLoad_Document(business.get_FolderNameToUpload(minuteDocumentType), "Meeting Minutes.docx", materialMetaData, blob);
            }).then(() => {
                this.setState({ minute: business.get_DocumentByType(minuteDocumentType), waitingMessage: '' });
            }).catch();
    }
}
