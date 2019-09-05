import * as React from 'react';
import styles from '../Meeting.module.scss';
import { IAgendaProps, IAgendaState, AgendaPanelType } from './IAgenda';
import css from '../../../../utility/css';
import { CommandBar, SelectionMode, Selection, DetailsList, DetailsListLayoutMode, IColumn, Panel, PanelType } from 'office-ui-fabric-react';
import { IComponentAgenda, get_tranformAgenda } from "../../../../business/transformAgenda";
// import tranformAgenda from './tranformAgenda';
// import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { McsUtil } from '../../../../utility/helper';
import { TopicDisplay } from './TopicDisplay';
import { MaterialDisplay } from './MaterialDisplay';
import { ISpEventMaterial } from '../../../../interface/spmodal';
import AgendaForm from './AgendaForm';
import MaterialForm from '../MaterialForm/MaterialForm';

export default class Agenda extends React.Component<IAgendaProps, IAgendaState> {

    private _selection: Selection;

    constructor(props: Readonly<IAgendaProps>) {
        super(props);
        this._selection = new Selection({
            onSelectionChanged: () => {
                this._setSelectedAgenda();
            }
        });
        this.state = {
            agendaItems: get_tranformAgenda(),
            showPanel: false,
            selectedAgendaItem: null,
            panelHeaderText: '',
            panelItem: null
        };
    }

    public render(): React.ReactElement<IAgendaProps> {
        const { agendaItems, panelHeaderText, panelType, panelItem, selectedAgendaItem } = this.state;
        const agendaSelected = McsUtil.isDefined(selectedAgendaItem);
        return (
            <div className={styles["container-fluid"]}>
                <div className={styles.row}>
                    <div className={styles["col-12"]}>
                        <CommandBar
                            items={this._getCommandBarItems(agendaSelected)}
                            overflowItems={[]}
                            farItems={[]}
                            ariaLabel={'Use left and right arrow keys to navigate between commands'}
                        />
                    </div>
                </div>
                <DetailsList
                    key="ListViewControl"
                    items={agendaItems}
                    columns={this._getListColumns()}
                    selectionMode={SelectionMode.single}
                    selection={this._selection}
                    layoutMode={DetailsListLayoutMode.justified}
                    compact={false}
                    setKey="ListViewControl" />
                <Panel
                    isOpen={this.state.showPanel}
                    onDismiss={this._hidePanel}
                    type={PanelType.custom}
                    customWidth="600px"
                    headerText={panelHeaderText}>
                    <div>
                        {(panelType == AgendaPanelType.topic || panelType == AgendaPanelType.subtopic) &&
                            <AgendaForm
                                parentTopicId={McsUtil.isDefined(selectedAgendaItem) ? selectedAgendaItem.Id : void (0)}
                                isSubTopic={panelType == AgendaPanelType.subtopic}
                                agendaNumber={this._getNextAgendaNumber(selectedAgendaItem, panelType)}
                                agenda={panelItem}
                                minDate={this.props.minDate}
                                maxDate={this.props.maxDate}
                                eventLookupId={this.props.eventLookupId}
                            />
                        }
                        {panelType == AgendaPanelType.uploadDocument &&
                            <MaterialForm requireAgendaSelection={false} document={panelItem} agenda={[selectedAgendaItem]} />}
                    </div>

                </Panel>
            </div>
        );
    }

    private _hidePanel = () => {
        this.setState({ showPanel: false });
    }

    private _onTopicAdd = (agenda: IComponentAgenda): void => {
        this.setState({
            showPanel: true,
            panelType: AgendaPanelType.topic,
            panelHeaderText: McsUtil.isDefined(agenda) ? "Edit Topic" : "Add Topic",
            panelItem: agenda,
        });
    }

    private _onTopicDisplayBtnsClicked = (agenda: IComponentAgenda, item: IComponentAgenda | null | undefined): void => {
        if (McsUtil.isDefined(item)) {
            this.setState({
                showPanel: true,
                panelType: AgendaPanelType.subtopic,
                panelHeaderText: 'Edit Topic',
                panelItem: item,
                selectedAgendaItem: agenda
            });
        } else {
            this.setState({
                showPanel: true,
                panelType: AgendaPanelType.subtopic,
                panelHeaderText: 'Add Topic',
                panelItem: null,
                selectedAgendaItem: agenda
            });
        }
    }

    private _onMaterialDisplayBtnClicked = (agenda: IComponentAgenda, item: ISpEventMaterial | null | undefined): void => {
        if (McsUtil.isDefined(item)) {
            this.setState({
                showPanel: true,
                panelType: AgendaPanelType.uploadDocument,
                panelHeaderText: 'Edit Material Properties',
                panelItem: item,
                selectedAgendaItem: agenda
            });
        } else {
            this.setState({
                showPanel: true,
                panelType: AgendaPanelType.uploadDocument,
                panelHeaderText: 'Upload Material',
                panelItem: null,
                selectedAgendaItem: agenda
            });
        }
    }

    private _setSelectedAgenda = (): void => {
        const selectionCount = this._selection.getSelectedCount();
        if (selectionCount == 0) {
            this.setState({ selectedAgendaItem: null });
        } else {
            this.setState({ selectedAgendaItem: (this._selection.getSelection()[0] as IComponentAgenda) });
        }
    }

    private _getNextAgendaNumber = (selectedAgendaItem: IComponentAgenda, panelType: AgendaPanelType): number => {
        if (panelType === AgendaPanelType.topic) {
            return this.state.agendaItems.length + 1;
        }
        if (McsUtil.isArray(selectedAgendaItem.SubTopics)) {
            if (selectedAgendaItem.SubTopics.length == 0) {
                return 1;
            } else {
                return selectedAgendaItem.SubTopics[selectedAgendaItem.SubTopics.length - 1].AgendaNumber + 1;
            }
        }
        return 1;
    }

    private _getListColumns = (): IColumn[] => {
        return [{
            name: 'Time',
            key: 'Time',
            isSorted: false,
            isResizable: true,
            minWidth: 115,
            maxWidth: 150,
            onRender: (item?: IComponentAgenda, index?: number) => {
                let dateval = '';
                if (McsUtil.isDefined(item.AgendaDate)) {
                    try {
                        dateval = (new Date(item.AgendaDate as any) as any).format('MM/dd/yyyy @ h:mma');
                    } catch{ dateval = ''; }
                }
                return (<div className={css.combine(styles["d-flex"], styles["flex-column"], styles["justify-content-around"])}>
                    {dateval.length > 0 && <div>dateval</div>}
                    <div>Agenda Number: {item.AgendaNumber}</div>
                </div>);
            }
        },
        {
            name: 'Topic/SubTopic',
            key: 'Title',
            isSorted: false,
            isResizable: true,
            minWidth: 50,
            onRender: (item?: IComponentAgenda, index?: number) => {
                return (<TopicDisplay agenda={item} onAddOrEditBtnClicked={this._onTopicDisplayBtnsClicked} />);
            }
        },
        {
            name: 'Material',
            key: 'Id',
            isSorted: false,
            isResizable: true,
            minWidth: 200,
            maxWidth: 300,
            onRender: (item?: IComponentAgenda, index?: number) => {
                return (<MaterialDisplay material={item.Documents}
                    agenda={item}
                    onAddOrUpdateMaterial={this._onMaterialDisplayBtnClicked}
                />);
            }
        }];
    }

    private _getCommandBarItems = (agendaSelected: boolean): any[] => {
        return [
            {
                key: 'newItem',
                name: 'New Agenda',
                // cacheKey: 'myCacheKey', // changing this key will invalidate this items cache
                iconProps: {
                    iconName: 'Add'
                },
                ariaLabel: 'New Agenda',
                // onclick: () => {
                //     this._onTopicAdd();
                // }
                onClick: () => {
                    this._onTopicAdd(null);
                }
            },
            {
                key: 'editItem',
                name: 'Edit Agenda',
                // cacheKey: 'myCacheKey', // changing this key will invalidate this items cache
                iconProps: {
                    iconName: 'Edit'
                },
                ariaLabel: 'Edit Agenda',
                disabled: !agendaSelected,
                onClick: () => {
                    this._onTopicAdd(this.state.selectedAgendaItem);
                }
            },
            {
                key: 'upload',
                name: 'Upload Material',
                iconProps: {
                    iconName: 'Upload'
                },
                disabled: !agendaSelected,
                onClick: () => {
                    this._onMaterialDisplayBtnClicked(this.state.selectedAgendaItem, null);
                }
            }
        ];
    }
}
