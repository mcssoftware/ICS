import * as React from 'react';
import styles from '../Meeting.module.scss';
import { IAgendaProps, IAgendaState, AgendaPanelType } from './IAgenda';
import css from '../../../../utility/css';
import { CommandBar, SelectionMode, Selection, DetailsList, DetailsListLayoutMode, IColumn, Panel, PanelType } from 'office-ui-fabric-react';
import { IComponentAgenda, get_tranformAgenda } from "../../../../business/transformAgenda";
import { Waiting } from '../../../../controls/waiting';
// import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { McsUtil } from '../../../../utility/helper';
import { TopicDisplay } from './TopicDisplay';
import { MaterialDisplay } from './MaterialDisplay';
import { ISpEventMaterial, OperationType } from '../../../../interface/spmodal';
import AgendaForm from './AgendaForm';
import MaterialForm from '../MaterialForm/MaterialForm';
import { findIndex, cloneDeep } from '@microsoft/sp-lodash-subset';
import { business } from '../../../../business';
import { Informational, InformationalType } from '../../../../controls/informational';

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
            panelItem: null,
            waitingMessage: '',
            message: '',
            messageType: InformationalType.none
        };
    }

    public render(): React.ReactElement<IAgendaProps> {
        const { agendaItems, panelHeaderText, panelType, panelItem, selectedAgendaItem, waitingMessage } = this.state;
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
                <div className={styles.row}>
                    <div className={styles["col-12"]}>
                        <DetailsList
                            key="ListViewControl"
                            items={agendaItems}
                            columns={this._getListColumns()}
                            selectionMode={SelectionMode.single}
                            selection={this._selection}
                            layoutMode={DetailsListLayoutMode.justified}
                            compact={false}
                            setKey="ListViewControl" />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles["col-12"]}>
                        <Informational message={this.state.message} type={this.state.messageType} />
                    </div>
                </div>
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
                                onChange={this._onNewAgendaAddedOrEdited}
                                onCancel={this._hidePanel}
                            />
                        }
                        {panelType == AgendaPanelType.uploadDocument &&
                            <MaterialForm requireAgendaSelection={false}
                                meetingId={this.props.eventLookupId}
                                document={panelItem}
                                agenda={[selectedAgendaItem]}
                                sortNumber={selectedAgendaItem.Documents.length > 0 ? selectedAgendaItem.Documents[selectedAgendaItem.Documents.length - 1].SortNumber : 1}
                                onChange={this._onMaterialUploaded}
                            />}
                    </div>
                </Panel>

                <Waiting message={waitingMessage} />
            </div>
        );
    }

    private _hidePanel = () => {
        this.setState({ showPanel: false });
    }

    private _onNewAgendaAddedOrEdited = (topic: IComponentAgenda, parentTopicId?: number): void => {
        const agendaCopy = [...this.state.agendaItems];
        let message = '';
        if (McsUtil.isDefined(parentTopicId)) {
            const agendaIndex = findIndex(agendaCopy, a => a.Id === parentTopicId);
            const subtopicIndex = findIndex(agendaCopy[agendaIndex].SubTopics, topic.Id);
            if (subtopicIndex >= 0) {
                agendaCopy[agendaIndex].SubTopics[subtopicIndex] = topic;
                message = 'Subtopic updated';
            } else {
                agendaCopy[agendaIndex].SubTopics.push(topic);
                message = 'Subtopic added';
            }
        } else {
            const agendaIndex = findIndex(agendaCopy, a => a.Id === topic.Id);
            if (agendaIndex >= 0) {
                agendaCopy[agendaIndex] = topic;
                message = 'Topic updated';
            } else {
                message = 'Subtopic added';
                agendaCopy.push(topic);
            }
        }
        this.setState({ agendaItems: agendaCopy, showPanel: false, waitingMessage: '', message, messageType: InformationalType.Info });
    }

    private _onTopicAddButtonClicked = (agenda: IComponentAgenda): void => {
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

    /**
     * on material has been uploaded.
     * @private
     * @memberof Agenda
     */
    private _onMaterialUploaded = (document: ISpEventMaterial, agenda: IComponentAgenda, type: OperationType): void => {
        if (McsUtil.isDefined(document)) {
            const event = business.get_Event();
            this.setState({ waitingMessage: "Attaching document to meeting." });
            business.edit_Event(event.Id, event["odata.type"],
                {
                    __metadata: {
                        type: "Collection(Edm.Int32)"
                    },
                    results: this._getDocumentLookupIds((McsUtil.isDefined(event.EventDocumentsLookupId) ? event.EventDocumentsLookupId.results : []), document.Id, type)
                }).then((e) => {
                    if (McsUtil.isDefined(agenda)) {
                        business.edit_Agenda(agenda.Id, agenda["odata.type"],
                            {
                                __metadata: {
                                    type: "Collection(Edm.Int32)"
                                },
                                results: this._getDocumentLookupIds((McsUtil.isDefined(agenda.AgendaDocumentsLookupId) ? agenda.AgendaDocumentsLookupId.results : []), document.Id, type)
                            }).then((updatedAgenda: IComponentAgenda) => {
                                updatedAgenda.SubTopics = [...agenda.SubTopics];
                                updatedAgenda.Presenters = [...agenda.Presenters];
                                updatedAgenda.Documents = [...agenda.Documents];

                                if (type === OperationType.Delete || type === OperationType.Edit) {
                                    const index = findIndex(updatedAgenda.Documents, a => a.Id == document.Id);
                                    if (index > -1) {
                                        if (type === OperationType.Edit) {
                                            updatedAgenda.Documents[index] = document;
                                        } else {
                                            updatedAgenda.Documents.splice(index, 1);
                                        }
                                    }
                                } else {
                                    updatedAgenda.Documents.push(document);
                                }
                                const tempAgenda = cloneDeep(this.state.agendaItems);
                                let found = false;
                                for (let i = 0; i < tempAgenda.length && !found; i++) {
                                    if (tempAgenda[i].Id === agenda.Id) {
                                        tempAgenda[i] = agenda;
                                        found = true;
                                        break;
                                    }
                                    for (let j = 0; j < tempAgenda[i].SubTopics.length && !found; j++) {
                                        if (tempAgenda[i].SubTopics[j].Id == agenda.Id) {
                                            tempAgenda[i].SubTopics[j] = agenda;
                                            found = true;
                                            break;
                                        }
                                    }
                                }
                                this.setState({ agendaItems: tempAgenda });
                            }).catch(() => { });
                    } else {
                        return Promise.resolve(null);
                    }
                }).then();
            // need to update event object as well.
            this.setState({ showPanel: false, waitingMessage: '' });
        } else {
            this.setState({ showPanel: false, waitingMessage: '' });
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
                    this._onTopicAddButtonClicked(null);
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
                    this._onTopicAddButtonClicked(this.state.selectedAgendaItem);
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

    private _getDocumentLookupIds = (oldDocIds: number[], docId: number, type: OperationType): number[] => {
        const documentLookupIds: number[] = McsUtil.isArray(oldDocIds) ? oldDocIds : [];
        if (type === OperationType.Add) {
            documentLookupIds.push(docId);
        }
        if (type === OperationType.Delete) {
            const docIndex = findIndex(documentLookupIds, a => a == docId);
            if (docIndex > -1) {
                documentLookupIds.splice(docIndex, 1);
            }
        }
        return documentLookupIds;
    }
}
