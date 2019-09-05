import * as React from 'react';
import styles from './tabs.module.scss';
import Tab from './tab';
import { McsUtil } from '../../utility/helper';

export interface ITabsProps {

}

export interface ITabsState {
    activeTab: string;
    canCreateTab: boolean;
}

export default class Tabs extends React.Component<ITabsProps, ITabsState>{
    /**
     *
     */
    constructor(props) {
        super(props);
        let canCreateTab = false;
        let activeTab = '';
        if (McsUtil.isArray(props.children)) {
            const children = props.children as Array<React.ReactElement>;
            if (children.length > 0) {
                canCreateTab = true;
                activeTab = children[0].props.title;
            }
        }
        this.state = {
            activeTab,
            canCreateTab
        };
    }

    public render(): React.ReactElement<ITabsProps> {
        const { children } = this.props;
        const { activeTab, canCreateTab } = this.state;
        return (<div className={styles.tabs}>
            {canCreateTab && <div>
                <ol className={styles["tab-list"]}>
                    {
                        (children as Array<React.ReactElement>).map((child) => { 
                            const { title } = child.props;
                            return (
                                <Tab
                                    activeTab={activeTab}
                                    key={title}
                                    label={title}
                                    onClick={this._onClickTabItem}
                                />
                            );
                        })
                    }
                </ol>
                <div className={styles["tab-content"]}>
                    {(children as Array<React.ReactElement>).map((child) => {
                        if (child.props.label !== this.state.activeTab) return undefined;
                        return child.props.children;
                    })}
                </div>
            </div>}
        </div>);
    }

    private _onClickTabItem = (tab: string) => {
        this.setState({ activeTab: tab });
    }
}