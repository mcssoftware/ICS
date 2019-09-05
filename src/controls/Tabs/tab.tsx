import * as React from 'react';
import { Link, IconButton, CommandBarButton, IIconProps } from 'office-ui-fabric-react';
import styles from './tabs.module.scss';
export interface ITabProps {
    activeTab: string;
    label: string;
    onClick: (tab: string) => void;
}


const Tab: React.SFC<ITabProps> = (props) => {

    const { activeTab, label, onClick } = props;
    let className = styles["tab-list-item"];
    if (activeTab === label) {
        className += ' ' + styles["tab-list-active"];
    }

    return (
        <li className={styles["tab-list-active"]} onClick={() => { onClick(label); }}>
            {label}
        </li>
    );
};


export default Tab;