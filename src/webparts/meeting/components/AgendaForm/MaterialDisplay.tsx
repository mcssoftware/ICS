import * as React from 'react';
import { Link, IconButton, CommandBarButton, IIconProps } from 'office-ui-fabric-react';
import { ISpEventMaterial } from "../../../../interface/spmodal";
import { McsUtil } from '../../../../utility/helper';
import css from '../../../../utility/css';
import styles from '../Meeting.module.scss';
import { IComponentAgenda } from '../../../../business/transformAgenda';

export interface IMaterialDisplayProps {
    agenda: IComponentAgenda;
    onAddOrUpdateMaterial: (agenda: IComponentAgenda, item: ISpEventMaterial | null | undefined) => void;
}

const uploadIcon: IIconProps = { iconName: 'CloudUpload' };

const materialDisplayPart: React.SFC<IMaterialDisplayProps> = (props) => {

    const displayAgendaMaterial = (agenda: IComponentAgenda, m: ISpEventMaterial) => {
        return (<li className={css.combine(styles["list-group-item"], styles["d-flex"], styles["p-0"])} style={{ whiteSpace: "normal" }}>
            <Link href={m.File.ServerRelativeUrl}>{m.SortNumber} - {m.Title}</Link>
            <div style={{ marginLeft: 'auto!important' }}>
                <IconButton iconProps={{ iconName: 'PageEdit' }} title="Edit" ariaLabel="Edit" onClick={() => props.onAddOrUpdateMaterial(agenda, m)} />
            </div>
        </li>);
    };

    return (
        <div className={styles.card}>
            {McsUtil.isArray(props.agenda.Documents) && props.agenda.Documents.length > 0 && <div className={styles["card-body"]}>
                <ul className={styles["list-group"]}>
                    {props.agenda.Documents.map((m) => { return displayAgendaMaterial(props.agenda, m); })}
                    {McsUtil.isArray(props.agenda.SubTopics) && props.agenda.SubTopics.length > 0 &&
                        props.agenda.SubTopics.map((subTopic) => {
                            return subTopic.Documents.map((doc) => {
                                return displayAgendaMaterial(subTopic, doc);
                            });
                        })
                    }
                </ul>
            </div>}
            <div className={styles["card-footer"]}>
                <CommandBarButton iconProps={uploadIcon} text="Upload Material" onClick={() => { props.onAddOrUpdateMaterial(props.agenda, void (0)); }} />
            </div>
        </div>
    );
};

export { materialDisplayPart as MaterialDisplay };
