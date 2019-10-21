import * as React from 'react';
import { business } from '../../../../business';
import { ISpEvent, ISpCommitteeLink } from '../../../../interface/spmodal';
import { findIndex } from "@microsoft/sp-lodash-subset";
import { Checkbox, PrimaryButton } from 'office-ui-fabric-react';
import { Promise } from 'es6-promise';
import { McsUtil } from '../../../../utility/helper';
import SpListService from '../../../../dal/spListService';
import styles from "./AddCommittee.module.scss";
import { Waiting } from '../../../../controls/waiting';


interface ISpCommitteeLinkTemp extends ISpCommitteeLink {
    selected: boolean;
}

export interface IAddCommitteeProps {
    onComplete: () => void;
}

export interface IAddCommitteeState {
    event: ISpEvent;
    currentCommittee: ISpCommitteeLink;
    options: ISpCommitteeLinkTemp[];
    waitingMessage: string;
}

export default class AddCommittee extends React.Component<IAddCommitteeProps, IAddCommitteeState> {
    constructor(props: IAddCommitteeProps) {
        super(props);
        const committeeList = business.get_CommitteeList();
        const committee = business.get_Committee();

        const options: ISpCommitteeLinkTemp[] = [];
        let currentCommittee: ISpCommitteeLink;
        committeeList.forEach(c => {
            if (c.CommitteeId !== Mcs.WebConstants.committeeId) {
                const index = findIndex(committee, a => a.Code == c.CommitteeId);
                options.push({ ...c, selected: index >= 0 });
            } else {
                currentCommittee = c;
            }
        });

        this.state = {
            event: business.get_Event(),
            options,
            currentCommittee,
            waitingMessage: ''
        };
    }

    public render() {
        const { options } = this.state;

        return (
            <div className={styles.addCommittee}>
                {options.map((c, i) => {
                    return (<Checkbox label={c.CommitteeName}
                        checked={c.selected}
                        onChange={(ev: any, isChecked: boolean) => { this._comitteeCheckboxOnchange(i, isChecked); }} />);
                })}
                <PrimaryButton text="Add committees" onClick={this._saveClicked} allowDisabledFocus />
                <Waiting message={this.state.waitingMessage} />
            </div>
        );
    }

    private _comitteeCheckboxOnchange = (index: number, isChecked: boolean): void => {
        const options = [...this.state.options];
        options[index].selected = isChecked;
        this.setState({ options });
    }

    private _saveClicked = (): void => {
        const { options } = this.state;
        this.setState({ waitingMessage: 'Adding or removing committees to this meeting.' });
        const oldCommiteeList = business.get_Committee().filter(a => a.Code != Mcs.WebConstants.committeeId).map(a => a.Code);
        const selectedCommitee = options.filter(a => a.selected);
        const newCommitteeSelected = selectedCommitee.map(a => a.CommitteeId);
        if (JSON.stringify(oldCommiteeList) === JSON.stringify(newCommitteeSelected)) {
            this._onComplete();
        } else {
            const committeeToAdd = selectedCommitee.filter(a => findIndex(oldCommiteeList, b => b == a.CommitteeId) < 0);
            const committeeToRemove = options.filter(a => !a.selected && (findIndex(oldCommiteeList, b => b == a.CommitteeId) >= 0));
            Promise.all([this._addToCommittee(committeeToAdd), this._removeFromCommittee(committeeToRemove)])
                .then(() => {
                    const jointCommittee = selectedCommitee.map(a => a.Id).join(";#");
                    const event = business.get_Event();
                    return business.edit_Event(event.Id, event["odata.type"], { JointEventCommitteeId: jointCommittee });
                }).then(() => this._onComplete())
                .catch(() => this._onComplete());
        }
    }

    private _onComplete = (): void => {
        this.setState({ waitingMessage: '' });
        this.props.onComplete();
    }

    private _addToCommittee = (committeeList: ISpCommitteeLinkTemp[]): Promise<void> => {
        return new Promise((resolve, reject) => {
            if (McsUtil.isArray(committeeList) && committeeList.length > 0) {
                const event = business.get_Event();
                const promises = committeeList.map(a => {
                    const committeeService = new SpListService<any>("Committee%20Calendar", false);
                    committeeService.setWebUrl(a.URL.Url);
                    var newEventTitle = {
                        Title: a.CommitteeName + " Committee Meeting - with " + Mcs.WebConstants.committeeFullName,
                        MeetingStartTime: event.MeetingStartTime.replace(/[^\d\s:APM]/gi, ""),
                        Location: event.Location,
                        Description: event.Description,
                        fAllDayEvent: event.fAllDayEvent,
                        fRecurrence: event.fRecurrence,
                        Category: event.Category,
                        WorkAddress: event.WorkAddress,
                        WorkCity: event.WorkCity,
                        WorkState: event.WorkState,
                        OtherLocationInfo: event.OtherLocationInfo,
                        CommitteeStaff: event.CommitteeStaff,
                        ConferenceNumber: event.ConferenceNumber,
                        CommitteeLookupId: this.state.currentCommittee.Id,
                        CommitteeEventLookupId: event.Id,
                        HasLiveStream: event.HasLiveStream,
                        IsBudgetHearing: event.IsBudgetHearing || false,
                        ApprovedStatus: "(none)",
                        EventDate: event.EventDate,
                        EndDate: event.EndDate,
                    };
                    return committeeService.addNewItem(newEventTitle);
                });

                Promise.all(promises).then((a) => {
                    resolve();
                }).catch((e) => reject(e));

            } else {
                resolve();
            }
        });
    }

    private _removeFromCommittee = (committeeList: ISpCommitteeLinkTemp[]): Promise<void> => {
        return new Promise((resolve, reject) => {
            if (McsUtil.isArray(committeeList) && committeeList.length > 0) {
                const event = business.get_Event();
                const committeeid = this.state.currentCommittee.Id;
                const promises = committeeList.map(a => {
                    return this._deleteFromCommitee(a, committeeid, event.Id);
                });

                Promise.all(promises).then((a) => {
                    resolve();
                }).catch((e) => reject(e));
            } else {
                resolve();
            }
        });
    }

    private _deleteFromCommitee = (committeeList: ISpCommitteeLinkTemp, committeeLookupId: number, eventId: number): Promise<void> => {
        return new Promise((resolve, reject) => {
            const committeeService = new SpListService<any>("Committee%20Calendar", false);
            committeeService.setWebUrl(committeeList.URL.Url);
            const filter = `CommitteeLookupId eq '${committeeLookupId}' and CommitteeEventLookupId eq '${eventId}'`;
            committeeService.getListItems(filter, null, null, null, 0, 1)
                .then((items) => {
                    if (McsUtil.isArray(items) && items.length == 1) {
                        committeeService.deleteItem(items[0].Id).then(() => resolve()).catch(() => resolve());
                    } else {
                        resolve();
                    }
                }).then(() => resolve()).catch((e) => reject(e));
        });
    }
}
