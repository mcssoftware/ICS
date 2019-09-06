import * as React from 'react';
import styles from '../Meeting.module.scss';
import { IEventProps, IEventState } from './IEvent';
import { ChoiceGroup, TextField, DatePicker, DayOfWeek, Toggle, PrimaryButton, Label, IChoiceGroupOption } from 'office-ui-fabric-react';
import css from '../../../../utility/css';
import datePickerStrings from '../../../../utility/datePickerStrings';
import Select from 'react-select';
import { findStateOption, usStates } from '../../../../utility/usstates';
import { Timepicker } from '../../../../controls/timepicker';
import { McsUtil } from '../../../../utility/helper';
import { ISpEvent } from '../../../../interface/spmodal';
import { business } from '../../../../business';
import { SPEvent } from '@microsoft/sp-core-library';

export default class Event extends React.Component<IEventProps, IEventState> {

    constructor(props: Readonly<IEventProps>) {
        super(props);
        var s1 = McsUtil.convertUtcDateToLocalDate(new Date(props.event.EventDate));
        var s2 = props.event.MeetingStartTime;
        const startDate = new Date(s1.toLocaleDateString() + ", " + s2.replace(/[^\d\s:APM]/gi, ""));
        const endDate = McsUtil.convertUtcDateToLocalDate(new Date(props.event.EndDate));
        this.state = {
            event: { ...props.event },
            startDate,
            endDate,
            selectedState: findStateOption(props.event.WorkState),
            isDirty: false,
        };
    }

    public render(): React.ReactElement<IEventProps> {
        const { event, startDate, endDate, selectedState } = this.state;
        const marginClassName = css.combine(styles["ml-2"], styles["mr-2"]);
        let minStartDate: Date;
        let maxStartDate: Date;
        if (event.Id > 0) {
            minStartDate = new Date(startDate.getFullYear(), 1, 1, 0, 0, 0);
            maxStartDate = new Date(startDate.getFullYear(), 12, 31, 23, 59, 59);
        }
        return (
            <div>
                <div className={styles.row}>
                    <div className={styles["col-12"]}>
                        <div className={marginClassName}>
                            <ChoiceGroup
                                className={css.combine(styles["checkbox-label-inline"], styles["checkbox-items-inline"])}
                                selectedKey={event.Category}
                                options={[
                                    {
                                        key: 'Tentative',
                                        text: 'Tentative Meeting'
                                    },
                                    {
                                        key: 'Formal',
                                        text: 'Formal Meeting'
                                    }
                                ]}
                                label="Meeting Type"
                                onChange={this._onCategoryChange}
                                required={true}
                            />
                        </div>
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles["col-12"]}>
                        <TextField label="The purpose of the meeting is to"
                            multiline={true}
                            className={marginClassName}
                            rows={5}
                            value={event.Description}
                            onChange={this._onDescriptionChange} />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={css.combine(styles["col-12"], styles["col-md-4"], styles["label-mg-b-10"])}>
                        <div className={marginClassName}>
                            <DatePicker
                                label="Start Date"
                                className={styles.dateTimePicker}
                                firstDayOfWeek={DayOfWeek.Monday}
                                strings={datePickerStrings}
                                showWeekNumbers={true}
                                firstWeekOfYear={1}
                                showMonthPickerAsOverlay={true}
                                placeholder="Select start date..."
                                ariaLabel="Select start date"
                                onSelectDate={this._onStartDateSelected}
                                value={startDate}
                                minDate={minStartDate}
                                maxDate={maxStartDate}
                            />
                        </div>
                    </div>
                    <div className={css.combine(styles["col-12"], styles["col-md-4"])}>
                        <div className={marginClassName}>
                            <Timepicker time={event.MeetingStartTime} onChange={this._onTimeChanged} />
                        </div>
                    </div>
                    <div className={css.combine(styles["col-12"], styles["col-md-4"], styles["label-mg-b-10"])}>
                        <div className={marginClassName}>
                            <DatePicker
                                label="End Date"
                                className={styles.dateTimePicker}
                                firstDayOfWeek={DayOfWeek.Monday}
                                strings={datePickerStrings}
                                showWeekNumbers={true}
                                firstWeekOfYear={1}
                                showMonthPickerAsOverlay={true}
                                placeholder="Select end date..."
                                ariaLabel="Select end date"
                                onSelectDate={this._onEndDateSelected}
                                value={endDate}
                                minDate={startDate}
                            />
                        </div>
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles["col-6"]}>
                        <TextField label="Telephone Conference" name="ConferenceNumber" className={marginClassName}
                            value={event.ConferenceNumber} onChange={this._onInputTxtChanged} />
                    </div>
                    <div className={styles["col-6"]}>
                        <TextField label="Committee Staff(s)" placeholder="" className={marginClassName}
                            value={event.CommitteeStaff} name="CommitteeStaff" onChange={this._onInputTxtChanged} />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles["col-5"]}>
                        <TextField label="Location 1" className={marginClassName} placeholder="Gillette College, Tech Building, Auditorium"
                            value={event.Location} name="Location" onChange={this._onInputTxtChanged} />
                    </div>
                    <div className={styles["col-2"]}>
                        <TextField label="Location 2" className={marginClassName} placeholder="Room 100"
                            value={event.OtherLocationInfo}
                            name="OtherLocationInfo" onChange={this._onInputTxtChanged} />
                    </div>
                    <div className={styles["col-5"]}>
                        <TextField label="Street Address" className={marginClassName} placeholder="200 W 24th St #213"
                            value={event.WorkAddress} name="WorkAddress" onChange={this._onInputTxtChanged} />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles["col-3"]}>
                        <TextField label="City" placeholder="Gillette"
                            name="WorkCity"
                            className={marginClassName}
                            value={event.WorkCity}
                            onChange={this._onInputTxtChanged}
                            required={event.Category == 'Formal'} />
                    </div>
                    <div className={styles["col-4"]}>
                        <div className={marginClassName}>
                            <Label>State</Label>
                            <Select
                                defaultValue={selectedState}
                                isClearable={false}
                                isSearchable={true}
                                name="USStates"
                                options={usStates}
                                onChange={this._onUsStateSelected}
                            />
                        </div>
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles["col-sm-6"]}>
                        <Toggle className={marginClassName} label="Will meeting be live streamed?" />
                    </div>
                    <div className={styles["col-sm-6"]}>
                        <Toggle className={marginClassName} label="Is budget hearing?" />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={css.combine(styles["col-12"], styles["d-flex"], styles["justify-content-around"], styles["mt-2"])}>
                        <PrimaryButton text="Save" />
                        <PrimaryButton text="Print View" />
                        <PrimaryButton text="Publish" />
                        <PrimaryButton text="Add committees to this meeting" />
                    </div>
                </div>
            </div>
        );
    }

    private _onCategoryChange = (ev?: any, option?: IChoiceGroupOption): void => {
        const event = { ...this.state.event };
        event.Category = option.key;
        this.setState({ event, isDirty: this._isDirty(event) });
    }

    private _onDescriptionChange = (ev?: any, newValue?: string): void => {
        const event = { ...this.state.event };
        event.Description = newValue;
        this.setState({ event, isDirty: this._isDirty(event) });
    }

    private _onStartDateSelected = (date: Date | null | undefined): void => {
        if (McsUtil.isDefined(date)) {
            this.setState({ startDate: date, isDirty: true });
        }
    }

    private _onEndDateSelected = (date: Date | null | undefined): void => {
        if (McsUtil.isDefined(date)) {
            this.setState({ endDate: date, isDirty: true });
        }
    }

    private _onInputTxtChanged = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        const event = { ...this.state.event };
        event[(ev.target as HTMLInputElement).name] = newValue;
        this.setState({ event, isDirty: this._isDirty(event) });
    }

    private _onTimeChanged = (newValue: string): void => {
        const event = { ...this.state.event };
        event.MeetingStartTime = newValue;
        this.setState({ event, isDirty: this._isDirty(event) });
    }

    private _onUsStateSelected = (value: any): void => {
        this.setState({ selectedState: value, isDirty: this._isDirty(event) });
    }

    private _isDirty = (newEvent: any): boolean => {
        return JSON.stringify(this.props.event) !== JSON.stringify(newEvent);
    }

    private _saveEvent = (): void => {
        const { event } = this.state;

        const propertiesToUpdate: ISpEvent = {
            EventDate: (new Date(this.state.startDate.toLocaleDateString())).toISOString(),
            EndDate: (new Date(this.state.endDate.toLocaleDateString())).toISOString(),
            MeetingStartTime: event.MeetingStartTime,
            Location: event.Location,
            Description: event.Description,
            fAllDayEvent: true,
            Title: `${Mcs.WebConstants.webTitle} Committee Meeting ${this.state.startDate.toLocaleDateString()}`,
            Category: event.Category,
            WorkAddress: event.WorkAddress,
            WorkCity: event.WorkCity,
            WorkState: event.WorkState,
            OtherLocationInfo: event.OtherLocationInfo,
            CommitteeStaff: event.CommitteeStaff,
            ConferenceNumber: event.ConferenceNumber,
            ApprovedStatus: event.ApprovedStatus,
            JointEventCommitteeId: event.JointEventCommitteeId,
            CommitteeLookupId: event.CommitteeLookupId,
            CommitteeEventLookupId: event.CommitteeEventLookupId,
            HasLiveStream: event.HasLiveStream,
            IsBudgetHearing: event.IsBudgetHearing
        };
        let promise: Promise<ISpEvent>;
        if (event.Id > 0) {
            promise = business.edit_Event(event.Id, event["odata.type"], propertiesToUpdate);
        } else {
            promise = business.add_Event(propertiesToUpdate);
        }
        promise.then((newevent) => {
            this.setState({ event: newevent });
            this.props.onChange();
        }).catch(() => { });
    }
}
