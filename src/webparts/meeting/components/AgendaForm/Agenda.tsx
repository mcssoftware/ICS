import * as React from 'react';
import styles from '../Meeting.module.scss';
import agendaDisplayStyles from './AgendaDisplay.module.scss';
import { IAgendaProps, IAgendaState, AgendaPanelType } from './IAgenda';
import css from '../../../../utility/css';
import { CommandBar, SelectionMode, Selection, DetailsList, DetailsListLayoutMode, IColumn, Panel, PanelType, IDragDropEvents, IDragDropContext, mergeStyles } from 'office-ui-fabric-react';
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
import IcsAppConstants from '../../../../configuration';

export default class Agenda extends React.Component<IAgendaProps, IAgendaState> {

    private _selection: Selection;
    private _dragDropEvents: IDragDropEvents;


    constructor(props: Readonly<IAgendaProps>) {
        super(props);
        this._selection = new Selection({
            onSelectionChanged: () => {
                this._setSelectedAgenda();
            }
        });
        this._dragDropEvents = this._getDragDropEvents();

        this.state = {
            agendaItems: get_tranformAgenda(),
            showPanel: false,
            selectedAgendaItem: null,
            panelHeaderText: '',
            panelItem: null,
            waitingMessage: '',
            message: '',
            messageType: InformationalType.none,
            orderChanged: false,
        };
    }

    public render(): React.ReactElement<IAgendaProps> {
        const { agendaItems, panelHeaderText, panelType, panelItem, selectedAgendaItem, waitingMessage, orderChanged } = this.state;
        const agendaSelected = McsUtil.isDefined(selectedAgendaItem);
        return (
            <div className={css.combine(styles["container-fluid"], agendaDisplayStyles.agendaDisplay)}>
                <div className={styles.row}>
                    <div className={styles["col-12"]}>
                        <CommandBar
                            items={this._getCommandBarItems(agendaSelected, orderChanged)}
                            overflowItems={[]}
                            farItems={this._getFarItems()}
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
                            dragDropEvents={this._dragDropEvents}
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
                                parentTopicId={McsUtil.isDefined(selectedAgendaItem) && (panelType == AgendaPanelType.subtopic) ? selectedAgendaItem.Id : void (0)}
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

    private _draggedItem: IComponentAgenda | undefined;
    private _draggedIndex: number;

    private _getDragDropEvents = (): IDragDropEvents => {
        const dragEnterClass = mergeStyles({
            backgroundColor: "black"
        });
        return {
            canDrop: (dropContext?: IDragDropContext, dragContext?: IDragDropContext) => {
                return true;
            },
            canDrag: (item?: any) => {
                return true;
            },
            onDragEnter: (item?: any, event?: DragEvent) => {
                // return string is the css classes that will be added to the entering element.
                return dragEnterClass;
            },
            onDragLeave: (item?: any, event?: DragEvent) => {
                return;
            },
            onDrop: (item?: any, event?: DragEvent) => {
                if (this._draggedItem) {
                    this._insertBeforeItem(item);
                }
            },
            onDragStart: (item?: any, itemIndex?: number, selectedItems?: any[], event?: MouseEvent) => {
                this._draggedItem = item;
                this._draggedIndex = itemIndex!;
            },
            onDragEnd: (item?: any, event?: DragEvent) => {
                this._draggedItem = undefined;
                this._draggedIndex = -1;
            }
        };
    }

    private _insertBeforeItem = (item: IComponentAgenda): void => {
        const draggedItems = this._selection.isIndexSelected(this._draggedIndex)
            ? (this._selection.getSelection() as IComponentAgenda[])
            : [this._draggedItem!];

        const items = this.state.agendaItems.filter(itm => draggedItems.indexOf(itm) === -1);
        let insertIndex = items.indexOf(item);

        // if dragging/dropping on itself, index will be 0.
        if (insertIndex === -1) {
            insertIndex = 0;
        }

        items.splice(insertIndex, 0, ...draggedItems);
        items.forEach((a, i) => a.AgendaNumber = i + 1);

        this.setState({ agendaItems: items, orderChanged: true });
    }

    private _onNewAgendaAddedOrEdited = (topic: IComponentAgenda, operationType: OperationType, parentTopicId?: number): void => {
        const agendaCopy = [...this.state.agendaItems];
        let message = '';
        if (McsUtil.isDefined(parentTopicId)) {
            const agendaIndex = findIndex(agendaCopy, a => a.Id === parentTopicId);
            const subtopicIndex = findIndex(agendaCopy[agendaIndex].SubTopics, a => a.Id === topic.Id);
            if (subtopicIndex >= 0) {
                if (operationType === OperationType.Delete) {
                    agendaCopy[agendaIndex].SubTopics.splice(subtopicIndex, 1);
                    message = 'Subtopic deleted';
                } else {
                    agendaCopy[agendaIndex].SubTopics[subtopicIndex] = topic;
                    message = 'Subtopic updated';
                }
            } else {
                agendaCopy[agendaIndex].SubTopics.push(topic);
                message = 'Subtopic added';
            }
        } else {
            const agendaIndex = findIndex(agendaCopy, a => a.Id === topic.Id);
            if (agendaIndex >= 0) {
                if (operationType === OperationType.Delete) {
                    agendaCopy.splice(agendaIndex, 1);
                    message = 'Topic deleted';
                } else {
                    agendaCopy[agendaIndex] = topic;
                    message = 'Topic updated';
                }
            } else {

                message = 'Topic added';
                agendaCopy.push(topic);
            }
        }
        this.setState({
            agendaItems: agendaCopy,
            selectedAgendaItem: McsUtil.isDefined(parentTopicId) ? this.state.selectedAgendaItem : topic,
            showPanel: false,
            waitingMessage: '',
            message,
            messageType: InformationalType.Info
        });
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
            if (type === OperationType.Edit) {
                const tempAgenda = cloneDeep(this.state.agendaItems);
                let documentFound = false;
                for (let i = 0; i < tempAgenda.length && !documentFound; i++) {
                    const index = findIndex(tempAgenda[i].Documents || [], a => a.Id == document.Id);
                    if (index >= 0) {
                        documentFound = true;
                        tempAgenda[i].Documents[index] = document;
                        break;
                    }
                    for (let j = 0; j < tempAgenda[i].SubTopics.length && !documentFound; j++) {
                        const subindex = findIndex(tempAgenda[i].SubTopics[j].Documents || [], a => a.Id == document.Id);
                        if (subindex >= 0) {
                            documentFound = true;
                            tempAgenda[i].SubTopics[j].Documents[subindex] = document;
                            break;
                        }
                    }
                }
                this.setState({ agendaItems: tempAgenda, showPanel: false, waitingMessage: '' });
            } else {
                this.setState({ showPanel: false, waitingMessage: 'Attaching document to meeting.' });
                const event = business.get_Event();
                const eventPropToUpdate = {};
                const eventLookupField = business.get_EventDocumentLookupField();
                const lookupids = this._getDocumentLookupIds((McsUtil.isArray(event[eventLookupField]) ? event[eventLookupField] as number[] : []), document.Id, type);
                eventPropToUpdate[eventLookupField] = {
                    __metadata: {
                        type: "Collection(Edm.Int32)"
                    },
                    results: [...lookupids]
                };
                business.edit_Event(event.Id, event["odata.type"], eventPropToUpdate)
                    .then((e) => {
                        if (McsUtil.isDefined(agenda)) {
                            const agendaPropToUpdate = {};
                            const agendaLookupField = business.get_AgendaDocumentLookupField();
                            const docIds = this._getDocumentLookupIds((McsUtil.isArray(agenda[agendaLookupField]) ? agenda[agendaLookupField] as number[] : []), document.Id, type);
                            agendaPropToUpdate[agendaLookupField] = {
                                __metadata: {
                                    type: "Collection(Edm.Int32)"
                                },
                                results: [...docIds]
                            };
                            business.edit_Agenda(agenda.Id, agenda["odata.type"], agendaPropToUpdate)
                                .then((updatedAgenda: IComponentAgenda) => {
                                    updatedAgenda.SubTopics = [...agenda.SubTopics];
                                    updatedAgenda.Presenters = [...agenda.Presenters];
                                    updatedAgenda.Documents = [...agenda.Documents];

                                    if (type === OperationType.Delete) {
                                        const index = findIndex(updatedAgenda.Documents, a => a.Id == document.Id);
                                        if (index > -1) {
                                            updatedAgenda.Documents.splice(index, 1);
                                        }
                                    } else {
                                        updatedAgenda.Documents.push(document);
                                    }
                                    const tempAgenda = cloneDeep(this.state.agendaItems);
                                    let found = false;
                                    for (let i = 0; i < tempAgenda.length && !found; i++) {
                                        if (tempAgenda[i].Id === updatedAgenda.Id) {
                                            tempAgenda[i] = updatedAgenda;
                                            found = true;
                                            break;
                                        }
                                        for (let j = 0; j < tempAgenda[i].SubTopics.length && !found; j++) {
                                            if (tempAgenda[i].SubTopics[j].Id == updatedAgenda.Id) {
                                                tempAgenda[i].SubTopics[j] = updatedAgenda;
                                                found = true;
                                                break;
                                            }
                                        }
                                    }
                                    this.setState({ agendaItems: tempAgenda, waitingMessage: '' });
                                }).catch(() => { });
                        } else {
                            this.setState({ showPanel: false, waitingMessage: '' });
                        }
                    }).catch();
            }
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

    private _saveAgendaNumber = (): void => {
        const { agendaItems } = this.state;
        Promise.all(agendaItems.map((a) => {
            return business.edit_Agenda(a.Id, a["odata.type"], { AgendaNumber: a.AgendaNumber });
        })).then(() => {
            this.setState({ orderChanged: false });
        });
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
                        const agendaDate = McsUtil.convertToISONoAllDay(item.AgendaDate as string);
                        if (!(agendaDate.getHours() === 0 && agendaDate.getMinutes() === 0)) {
                            dateval = (agendaDate as any).format('MM/dd/yyyy @ h:mmA');
                        }
                    } catch{ dateval = ''; }
                }
                return (<div className={css.combine(styles["d-flex"], styles["flex-column"], styles["justify-content-around"])}>
                    {dateval.length > 0 && <div>{dateval}</div>}
                    <div><strong>Agenda Number: </strong>{item.AgendaNumber}</div>
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
                return (<MaterialDisplay agenda={item}
                    onAddOrUpdateMaterial={this._onMaterialDisplayBtnClicked}
                />);
            }
        }];
    }

    private _getCommandBarItems = (agendaSelected: boolean, orderChanged: boolean): any[] => {
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
            },
            {
                key: 'saveOrder',
                name: 'Save Agenda Order',
                iconProps: {
                    iconName: 'Sort'
                },
                disabled: !orderChanged,
                onClick: () => {
                    this._saveAgendaNumber();
                }
            }
        ];
    }

    private _getFarItems = () => {
        return [
            {
                key: 'print',
                name: 'Print',
                ariaLabel: 'Print',
                iconProps: {
                    iconName: 'Print'
                },
                onClick: () => {
                    this.setState({ waitingMessage: 'Generating agenda (PREVIEW)' });
                    business.generateMeetingDocument(IcsAppConstants.getCreateAgendaPreviewPartial(), '')
                        .then((blob) => {
                            const preview = {
                                Title: IcsAppConstants.getPreviewFolder(),
                                AgencyName: "LSO",
                                lsoDocumentType: IcsAppConstants.getPreviewFolder(),
                                IncludeWithAgenda: false,
                                SortNumber: 1
                            };
                            return business.upLoad_Document(business.get_FolderNameToUpload(IcsAppConstants.getPreviewFolder()), "Agenda Preview.pdf", preview, blob);
                        }).then((item) => {
                            window.open(McsUtil.makeAbsUrl(item.File.ServerRelativeUrl), '_blank');
                            this.setState({ waitingMessage: '' });
                        }).catch(() => {
                            this.setState({ waitingMessage: '' });
                        });
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

    private _canDisplayDate = (): boolean => {
        return false;
    }
}
