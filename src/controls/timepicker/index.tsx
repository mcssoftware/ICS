import * as React from 'react';
import styles from './timepicker.module.scss';
import { IconButton, DefaultButton, TextField } from 'office-ui-fabric-react';
import { McsUtil } from '../../utility/helper';

export interface ITimepickerProps {
    time: string;
    onChange: (newValue: string) => void;
}

export interface ITimepickerState {
    hour: number;
    minutes: number;
    minutestr: string;
    ampm: string;
}

export class Timepicker extends React.Component<ITimepickerProps, ITimepickerState>{
    constructor(props: Readonly<ITimepickerProps>) {
        super(props);
        let hour = 12;
        let minutes = 0;
        let ampm = 'AM';
        try {
            var tempDate = new Date(`2001-01-01 ${props.time}`);
            if (typeof tempDate !== 'undefined' && tempDate !== null) {
                hour = tempDate.getHours();
                minutes = tempDate.getMinutes();
                ampm = hour >= 12 ? 'PM' : 'AM';
                hour = hour % 12;
                hour = hour ? hour : 12; // the hour '0' should be '12'
            }
        } catch{ }
        this.state = {
            hour,
            minutes,
            minutestr: McsUtil.padNumber(minutes, 2),
            ampm
        };
        // this._onTimeChanged({ hour, minutes, ampm });
    }

    public render(): React.ReactElement<ITimepickerProps> {
        return (
            <div className={styles.timePicker}>
                <table className={styles.timePickerTable}>
                    <tr>
                        <td className={styles.colHour}>
                            <IconButton iconProps={{ iconName: 'ChevronUp' }}
                                onClick={() => this._changeHour(1)}
                                title="add hour" ariaLabel="add hour" />
                        </td>
                        <td className={styles.colEmpty}>&nbsp;</td>
                        <td className={styles.colSecond}>
                            <IconButton iconProps={{ iconName: 'ChevronUp' }}
                                onClick={() => this._changeMinute(1)}
                                title="add second" ariaLabel="add second" />
                        </td>
                        <td className={styles.colEmpty}>&nbsp;</td>
                        <td className={styles.colbtn}>&nbsp;</td>
                    </tr>
                    <tr>
                        <td className={styles.colHour}>
                            <TextField value={this.state.hour.toString()} onChange={this._hourChanged} />
                        </td>
                        <td className={styles.colEmpty}>:</td>
                        <td className={styles.colSecond}>
                            <TextField value={this.state.minutestr} onChange={this._minutesChanged} />
                        </td>
                        <td className={styles.colEmpty}>&nbsp;</td>
                        <td className={styles.colbtn}>
                            <DefaultButton text={this.state.ampm} onClick={this._btnClicked} className={styles.btn} />
                        </td>
                    </tr>
                    <tr>
                        <td className={styles.colHour}>
                            <IconButton iconProps={{ iconName: 'ChevronDown' }}
                                onClick={() => this._changeHour(-1)}
                                title="substract hour" ariaLabel="substract hour" />
                        </td>
                        <td className={styles.colEmpty}>&nbsp;</td>
                        <td className={styles.colSecond}>
                            <IconButton iconProps={{ iconName: 'ChevronDown' }}
                                onClick={() => this._changeMinute(-1)}
                                title="substract second" ariaLabel="substract second" />
                        </td>
                        <td className={styles.colEmpty}>&nbsp;</td>
                        <td className={styles.colbtn}>&nbsp;</td>
                    </tr>
                </table>
            </div>
        );
    }

    private _hourChanged = (event: any, newValue?: string): void => {
        const re = /^[0-9\b]+$/;
        if (newValue === '' || re.test(newValue)) {
            const hour = parseInt(newValue);
            this.setState({ hour });
            this._onTimeChanged({ hour, minutes: this.state.minutes, ampm: this.state.ampm });
        }
    }

    private _minutesChanged = (event: any, newValue?: string): void => {
        const re = /^[0-9\b]+$/;
        if (newValue === '' || re.test(newValue)) {
            const minutes = parseInt(newValue);
            this.setState({ minutestr: newValue, minutes });
            this._onTimeChanged({ hour: this.state.hour, minutes, ampm: this.state.ampm });
        }
    }

    private _changeHour = (val: number) => {
        let hour = this.state.hour + val;
        hour = hour % 12;
        hour = hour ? hour : 12; // the hour '0' should be '12'
        this.setState({ hour });
        this._onTimeChanged({ hour, minutes: this.state.minutes, ampm: this.state.ampm });
    }

    private _changeMinute = (val: number) => {
        let minutes = (this.state.minutes + val) % 60;
        if (minutes < 0) {
            minutes = 59;
        }
        this.setState({ minutes, minutestr: minutes.toString() });
        this._onTimeChanged({ hour: this.state.hour, minutes, ampm: this.state.ampm });
    }

    private _btnClicked = (): void => {
        let value = 'AM';
        if (this.state.ampm === 'AM') {
            value = 'PM';
        }
        this.setState({ ampm: value });
        this._onTimeChanged({ hour: this.state.hour, minutes: this.state.minutes, ampm: value });
    }

    private _onTimeChanged = ({ hour, minutes, ampm }): void => {
        if (typeof this.props.onChange === 'function') {
            // const { hour, minutes, ampm } = this.state;
            const dateTime = new Date(`2001:01-01 ${hour}:${minutes} ${ampm}`);
            this.props.onChange(dateTime.toLocaleTimeString());
        }
    }
}
