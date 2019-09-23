import * as React from 'react';
import { ChoiceGroup, IChoiceGroupOption, DefaultButton } from 'office-ui-fabric-react';
import { McsUtil } from '../../../../utility/helper';
import { business } from '../../../../business';
import IcsAppConstants from '../../../../configuration';

export interface IEventPublisherProps {
    onComplete: () => void;
}

export interface IEventPublisherState {
    waitingMessage: string;
    publishingType: string;
}

enum PublishingType {
    MeetingNoticeOnly = 1,
    Minutes,
    All
}

export default class EventPublisher extends React.Component<IEventPublisherProps, IEventPublisherState>
{
    constructor(props: Readonly<IEventPublisherProps>) {
        super(props);
        this.state = {
            waitingMessage: '',
            publishingType: ''
        };
    }

    public render(): React.ReactElement<IEventPublisherProps> {
        return (
            <div>
                <ChoiceGroup
                    value={this.state.publishingType}
                    options={[
                        {
                            key: 'Tentative',
                            text: 'Publish Tentative Meeting Notice'
                        },
                        {
                            key: 'FormalMeetingNotice',
                            text: 'First Formal Meeting Notice for Director Approval'
                        },
                        {
                            key: 'FormalAgenda',
                            text: 'First Formal Meeting Notice & Agenda for Director Approval',
                        },
                        {
                            key: 'MeetingNotice',
                            text: 'Update Meeting Notice'
                        },
                        {
                            key: 'Agenda',
                            text: 'Update Meeting Notice & Agenda'
                        }
                    ]}
                    onChange={this._onPublishTypeSelected}
                    label="Select Publish Type"
                    required={true}
                />
                <DefaultButton text="Publish" disabled={!McsUtil.isString(this.state.publishingType)} onClick={this._onPublishButtonClicked} checked={true} />
            </div>
        );
    }

    private _onPublishTypeSelected = (ev?: any, option?: IChoiceGroupOption): void => {
        this.setState({ publishingType: option.key });
    }

    private _onPublishButtonClicked = (): void => {
        this.setState({ waitingMessage: 'Publishing Meeting' });
        let lastUpdated = null;
        business.loadEvent().then(() => {
            const postdata = business.get_publishingMeeting();
            lastUpdated = postdata.LastUpdated;
            return Promise.all([business.generateMeetingDocument(IcsAppConstants.getCreateMeetingNoticePartial(), postdata),
            business.generateMeetingDocument(IcsAppConstants.getCreateAgendaPreviewPartial(), postdata),
            this._initiateFlow(PublishingType.All)
            ]);
        }).then((responses) => {
            const meetingnoticeData = {
                Title: "Meeting Notice",
                AgencyName: "LSO",
                lsoDocumentType: "Meeting Notice",
                IncludeWithAgenda: false,
                SortNumber: 1
            };
            const meetingNotice = business.get_DocumentByType(meetingnoticeData.lsoDocumentType);
            let meetingNoticePromise;
            if (!McsUtil.isDefined(meetingNotice) || lastUpdated <= new Date(meetingNotice.Modified as string)) {
                meetingNoticePromise = business.upLoad_Document(business.get_FolderNameToUpload(meetingnoticeData.lsoDocumentType), "Meeting Notice.pdf", meetingnoticeData, responses[0]);
            } else {
                meetingNoticePromise = Promise.resolve(meetingNotice);
            }
            const agendaMetadata = {
                Title: "Agenda",
                AgencyName: "LSO",
                lsoDocumentType: "Agenda Preview",
                IncludeWithAgenda: true,
                SortNumber: 1
            };

            const agenda = business.get_DocumentByType(agendaMetadata.lsoDocumentType);
            let agendaPromise;
            if (!McsUtil.isDefined(agenda) || lastUpdated <= new Date(agenda.Modified as string)) {
                agendaPromise = business.upLoad_Document(business.get_FolderNameToUpload(agendaMetadata.lsoDocumentType), "AgendaPreview.pdf", meetingnoticeData, responses[1]);
            } else {
                agendaPromise = Promise.resolve(agenda);
            }
            return Promise.all([meetingNoticePromise, agendaPromise]);
        }).then(files => {
            return Promise.all([
                business.publishDocument(files[0]),
                business.publishDocument(files[1]),
                business.publishDocument(business.get_DocumentByType("Meeting Minutes"))
            ]);
        }).then(() => {
            this.setState({ waitingMessage: '' });
            this.props.onComplete();
        });
    }

    private _initiateFlow = (publishType: PublishingType): Promise<void> => {
        return new Promise((resolve, reject) => {
            var approvalType = ""; //"Tentative";
            var flowMessage = "";

            if (publishType === PublishingType.Minutes) {
                approvalType = "Minute";
                flowMessage = "Minutes approval for " + Mcs.WebConstants.committeeFullName;
            }
            else {
                approvalType = this.state.publishingType;
                if (approvalType.indexOf("Formal") < 0) {
                    flowMessage = "Confirm publishing of " + Mcs.WebConstants.committeeFullName + " meeting.";
                } else {
                    flowMessage = "Approve " + Mcs.WebConstants.committeeFullName + " meeting.";
                }
            }
            const eventId = business.get_Event().Id.toString();
            var baseItem = {
                //Title: "Approve " + WebConstants.committeeFullName + " meeting.",
                Title: flowMessage,
                PublishSiteUrl: business.getWebUrl(),
                CommitteeId: Mcs.WebConstants.committeeId,
                PublishEventId: eventId,
                ApprovalType: approvalType,
                ApprovalItemUrl: window.location.href.split("?")[0] + "?publishing=true&calendarItemId=" + eventId
            };
            business.get_MeetingApprovalListService().addNewItem(baseItem).then(() => {
                resolve();
            }).catch((e) => reject(e));
        });
    }

}