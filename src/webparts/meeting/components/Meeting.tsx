import * as React from 'react';
import styles from './Meeting.module.scss';
import { IMeetingProps, IMeetingState } from './IMeeting';
import { DefaultButton } from 'office-ui-fabric-react';
import Event from './EventForm/Event';
import Agenda from './AgendaForm/Agenda';
import css from '../../../utility/css';
import { business } from '../../../business';
import { McsUtil } from '../../../utility/helper';
import { MaterialForm } from './MaterialForm/Material';
import { Waiting } from '../../../controls/waiting';
import { MinuteForm } from "./MinuteForm/MinuteForm";

export default class Meeting extends React.Component<IMeetingProps, IMeetingState> {

  constructor(props: Readonly<IMeetingProps>) {
    super(props);
    this.state = {
      isLoaded: false,
      selectedTab: 'Event',
      isNewEvent: true,
      message: 'Loading'
    };
    business.on_Loaded((error: any) => {
      if (!McsUtil.isDefined(error)) {
        this._onDataLoaded();
      } else {
        this.setState({ isLoaded: false, message: error.response.statusText });
      }
    });
  }

  public render(): React.ReactElement<IMeetingProps> {
    const { isLoaded, selectedTab, isNewEvent, message } = this.state;
    const event = business.get_Event();
    let minDate: Date = undefined;
    let maxDate: Date = undefined;

    return (
      <div className={styles["container-fluid"]}>
        <div className={styles.row}>
          <div className={styles["col-12"]}>
            <ul className={css.combine(styles["list-group"], styles["list-group-horizontal"], styles["justify-content-between"])} style={{ marginBottom: '15px' }}>
              {this._getTopNavList(isNewEvent).map((list) => {
                const liclassnames = css.combine(styles["list-group-item"], styles.topnav, list.text == selectedTab ? styles.active : '', list.disabled ? styles.disabled : '');
                const btnclassnames = css.combine(styles["topnav-btn"], list.text == selectedTab ? styles.active : '');
                return (<li className={liclassnames}>
                  <DefaultButton text={list.text} className={btnclassnames} disabled={list.disabled} onClick={() => this._onTabChoiceClicked(list)} />
                </li>);
              })}
            </ul>
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles["col-12"]}>
            {!isLoaded && <Waiting message={message} />}
            {isLoaded && selectedTab === 'Event' &&
              <Event event={event}
                committees={business.get_Committee()}
                onChange={this._eventAddedOrUpdated} />}
            {isLoaded && !isNewEvent && selectedTab === 'Agenda' &&
              <Agenda minDate={minDate} maxDate={maxDate} eventLookupId={event.Id} />}
            {isLoaded && !isNewEvent && selectedTab === 'Materials' && <MaterialForm />}
            {isLoaded && !isNewEvent && selectedTab === 'Minutes' && <MinuteForm event={event} />}
          </div>
        </div>
        <a href="" style={{ display: "none" }} id="noticePreview"></a>
      </div>
    );
  }


  private _onTabChoiceClicked = (option: any): void => {
    this.setState({ selectedTab: option.key });
  }

  private _onDataLoaded = (): void => {
    const event = business.get_Event();
    let isNewEvent = true;
    if (McsUtil.isDefined(event)) {
      if (event.Id > 0) {
        isNewEvent = false;
      }
    }
    this.setState({
      isLoaded: true,
      isNewEvent
    });
  }

  private _eventAddedOrUpdated = (): void => {
    // const event = business.get_Event();
    // if (this.state.isNewEvent) {
    //   if (event.Id > 0) {
    //     const startdate = new Date(event.EventDate);
    //     business.ensure_Folders(startdate.getFullYear(), event.Id);
    //   }
    // }
    // to ensure component is refreshed.
    this.setState({ isLoaded: true });
  }

  private _getTopNavList = (isNewEvent: boolean, committeeId?: string): any[] => {
    const topNavList = [
      {
        key: 'Event',
        text: 'Event'
      },
      {
        key: 'Agenda',
        text: 'Agenda',
        disabled: isNewEvent
      },
      {
        key: 'Materials',
        text: 'Materials',
        disabled: isNewEvent
      },
      {
        key: 'Minutes',
        text: 'Minutes',
        disabled: isNewEvent
      }
    ];
    return topNavList;
  }
}
