import * as React from 'react';
import { McsUtil } from '../../../../utility/helper';
import { List, IconButton, CommandBarButton, IIconProps } from 'office-ui-fabric-react';
import css from '../../../../utility/css';
import styles from '../Meeting.module.scss';
import { IComponentAgenda } from '../../../../business/transformAgenda';

export interface ITopicDisplayProps {
    agenda: IComponentAgenda;
    onAddOrEditBtnClicked: (parentTopic: IComponentAgenda, item: IComponentAgenda | null | undefined) => void;
}
const presenterDiplay = (agenda: IComponentAgenda): string => {
    var presenterValue = "";
    if (McsUtil.isArray(agenda.Presenters) && agenda.Presenters.length > 0) {
        var tempPresenters = [];
        var presentersList = agenda.Presenters.sort((a, b) => {
            return a.SortNumber - b.SortNumber;
        });
        for (var i = 0; i < presentersList.length; i++) {
            var meetingPresenters = presentersList[i];
            var temp = meetingPresenters.PresenterName.trim();
            if (typeof meetingPresenters.Title === "string" && meetingPresenters.Title.trim().length > 0) {
                temp += ", " + meetingPresenters.Title.trim();
            }
            if (typeof meetingPresenters.OrganizationName === "string" && meetingPresenters.OrganizationName.trim().length > 0) {
                temp += ", " + meetingPresenters.OrganizationName.trim();
            }
            tempPresenters.push(temp);
        }
        presenterValue = " - " + tempPresenters.join("; ");
    }
    return presenterValue;
};

const subTopicItemRender = (parentTopic: IComponentAgenda, item: IComponentAgenda, index: number | undefined,
    callback: (parentTopic: IComponentAgenda, item: IComponentAgenda | null | undefined) => void): JSX.Element => {
    return (
        <div className={css.combine(styles["d-flex"], styles["justify-content-between"])}
            style={{ marginLeft: '10px' }} data-is-focusable={true}>
            <div style={{ fontSize: '14px', paddingTop: '8px' }}>{item.AgendaTitle}</div>
            <div style={{ marginLeft: 'auto!important' }}>
                <IconButton iconProps={{ iconName: 'PageEdit' }} title="Edit" ariaLabel="Edit" onClick={() => callback(parentTopic, item)} />
            </div>
        </div>
    );
};

const addIcon: IIconProps = { iconName: 'Add' };

const topicDisplay: React.SFC<ITopicDisplayProps> = (props) => {

    return (
        <div className={styles.card}>
            <div className={styles["card-header"]}>
                {props.agenda.AgendaTitle} {presenterDiplay(props.agenda)}
            </div>
            <div className={styles["card-body"]}>
                <List items={props.agenda.SubTopics} onRenderCell={(item, index) => {
                    return subTopicItemRender(props.agenda, item, index, props.onAddOrEditBtnClicked);
                }} />
            </div>
            <div className={styles["card-footer"]}>
                <CommandBarButton iconProps={addIcon} text="New sub topic" onClick={() => { props.onAddOrEditBtnClicked(props.agenda, void (0)); }} />
            </div>
        </div>
    );
};

export { topicDisplay as TopicDisplay };
