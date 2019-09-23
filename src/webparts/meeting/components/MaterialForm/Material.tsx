import * as React from 'react';
import { CommandBar, DetailsList, SelectionMode, DetailsListLayoutMode, IColumn, MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { McsUtil } from '../../../../utility/helper';
import materialStyle from './Material.module.scss';
import styles from '../Meeting.module.scss';
import { get_eventDocuments, get_tranformAgenda } from '../../../../business/transformAgenda';
import { sortBy } from '@microsoft/sp-lodash-subset';

interface IMaterialDocuments {
    sortNumber: number;
    displayIndex: string;
    agendaTitle: string;
    description: string;
    provider: string;
    url: string;
}
export interface IMaterialProps {
}

const getFarItems = () => {
    return [
        {
            key: 'print',
            name: 'Print',
            ariaLabel: 'Print',
            iconProps: {
                iconName: 'Print'
            },
            onClick: () => console.log('Sort')
        }
    ];
};

const getDocuments = (): IMaterialDocuments[] => {
    const eventDocuments: IMaterialDocuments[] = sortBy(get_eventDocuments()
        .filter(a => a.lsoDocumentType == 'Agenda Preview' || a.lsoDocumentType == 'Sign-in Sheet' || a.lsoDocumentType == 'Agenda Preview'), b => b.SortNumber)
        .map(a => {
            let sortIndex = 0;
            let numberOfminutesApproval = 0;
            let title = a.Title;
            let displayIndex = '';
            switch (a.lsoDocumentType) {
                case 'Agenda Preview': sortIndex = 1; displayIndex = '1-01'; break;
                case 'Sign-in Sheet': sortIndex = 2; displayIndex = '1-02'; break;
                case 'Minutes Approval': {
                    sortIndex = 3 + numberOfminutesApproval;
                    numberOfminutesApproval++;
                    displayIndex = `1-${McsUtil.padNumber(sortIndex, 2)}`;
                    break;
                }
            }
            return {
                sortNumber: sortIndex,
                displayIndex,
                agendaTitle: '',
                description: a.Title,
                provider: a.AgencyName,
                url: a.File.ServerRelativeUrl
            };
        });
    let lastAgendaIndex = 1;
    sortBy(get_tranformAgenda(), b => b.AgendaNumber).forEach(a => {
        lastAgendaIndex = a.AgendaNumber;
        let counter = 1;
        if (a.Documents.length > 0) {
            sortBy(a.Documents, d => d.SortNumber).forEach(d => {
                const document: IMaterialDocuments = {
                    sortNumber: a.AgendaNumber * 100 + counter++,
                    displayIndex: `${a.AgendaNumber}-${McsUtil.padNumber(d.SortNumber, 2)}`,
                    agendaTitle: a.Title,
                    description: d.Title,
                    provider: d.AgencyName,
                    url: d.File.ServerRelativeUrl
                };
                eventDocuments.push(document);
            });
        }
        if (a.SubTopics.length > 0) {
            sortBy(a.SubTopics, b => b.AgendaNumber).forEach(s => {
                if (s.Documents.length > 0) {
                    sortBy(s.Documents, d => d.SortNumber).forEach(d => {
                        const document: IMaterialDocuments = {
                            sortNumber: s.AgendaNumber * 100 + counter++,
                            displayIndex: `${s.AgendaNumber}-${McsUtil.padNumber(d.SortNumber, 2)}`,
                            agendaTitle: s.Title,
                            description: d.Title,
                            provider: d.AgencyName,
                            url: d.File.ServerRelativeUrl
                        };
                        eventDocuments.push(document);
                    });
                }
            });
        }
    });
    (sortBy(get_eventDocuments().filter(a => a.lsoDocumentType == 'Meeting Attachments'), d => d.SortNumber))
        .forEach((a, i) => {
            let sortIndex = (lastAgendaIndex + 1) * 100 + (i + 1);
            const document = {
                sortNumber: sortIndex,
                displayIndex: `${lastAgendaIndex + 1}-${McsUtil.padNumber(i + 1, 2)}`,
                agendaTitle: '',
                description: a.Title,
                provider: a.AgencyName,
                url: a.File.ServerRelativeUrl
            };
            eventDocuments.push(document);
        });
    return sortBy(eventDocuments, d => d.sortNumber);
};

const getListColumns = (): IColumn[] => {
    return [{
        name: 'Index Number',
        key: 'displayIndex',
        fieldName: 'displayIndex',
        isSorted: false,
        isResizable: true,
        // className: styles["col-2"],
        minWidth: 50
    },
    {
        name: 'Agenda Item',
        key: 'agendaTitle',
        fieldName: 'agendaTitle',
        isSorted: false,
        isResizable: true,
        // className: styles["col-3"],
        minWidth: 150,
    },
    {
        name: 'Document Description',
        key: 'description',
        fieldName: 'description',
        isSorted: false,
        isResizable: true,
        // className: styles["col-4"],
        minWidth: 200,
    },
    {
        name: 'Document Provider',
        key: 'provider',
        fieldName: 'provider',
        isSorted: false,
        isResizable: true,
        // className: styles["col-3"],
        minWidth: 150,
    }];
};

const materialDisplayPart: React.SFC<IMaterialProps> = (props) => {
    return (
        <div className={styles["container-fluid"] + " " + materialStyle.materialDisplay}>
            <div className={styles.row}>
                <div className={styles["col-12"]}>
                    <MessageBar
                        messageBarType={MessageBarType.warning}
                        isMultiline={true}>
                        <strong>Please Note:{' '}</strong>
                        This index is subject to change. It will include documents submitted prior to, during and after the meeting to which it applies. The index may contain documents that were not available prior to the meeting and documents which were not considered at the meeting.{' '}
                    </MessageBar>
                </div>
            </div>
            <div className={styles.row}>
                <div className={styles["col-12"]}>
                    <CommandBar
                        items={[]}
                        overflowItems={[]}
                        farItems={getFarItems()}
                        ariaLabel={'Use left and right arrow keys to navigate between commands'}
                    />
                </div>
            </div>
            <DetailsList
                key="ListViewControl"
                items={getDocuments()}
                columns={getListColumns()}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                compact={false}
                setKey="ListViewControl" />
        </div>
    );
};

export { materialDisplayPart as MaterialForm };
