import * as React from 'react';
import { IAgendaFormProps, IAgendaFormState } from './IAgenda';
import styles from '../Meeting.module.scss';
import { Checkbox, TextField, DefaultButton, IconButton, IIconProps, DatePicker, DayOfWeek } from 'office-ui-fabric-react';
import { Waiting } from '../../../../controls/waiting';
import css from '../../../../utility/css';
import { McsUtil } from '../../../../utility/helper';
import { sortBy, cloneDeep, findIndex } from "@microsoft/sp-lodash-subset";
import { ISpPresenter, ISpAgendaTopic } from '../../../../interface/spmodal';
import datePickerStrings from '../../../../utility/datePickerStrings';
import { Timepicker } from '../../../../controls/timepicker';
import { business } from '../../../../business';
import { IComponentAgenda } from '../../../../business/transformAgenda';

const userRemoveIcon: IIconProps = { iconName: 'Delete' };
const editContactIcon: IIconProps = { iconName: 'EditContact' };


export default class AgendaForm extends React.Component<IAgendaFormProps, IAgendaFormState> {

    constructor(props: Readonly<IAgendaFormProps>) {
        super(props);
        let useTime = true;
        let agendaDate = McsUtil.isDefined(props.minDate) ? props.minDate : new Date();
        let agendaTime = '08:00:00 AM';
        if (props.isSubTopic) {
            useTime = false;
        } else {
            if (McsUtil.isDefined(props.agenda)) {
                const tempdate = new Date(props.agenda.AgendaDate as string);
                if (tempdate.getHours() == 0 && tempdate.getMinutes() == 0) {
                    useTime = false;
                    agendaTime = tempdate.toLocaleTimeString();
                } else {
                    agendaDate = tempdate;
                }
            }
        }
        const defaultAgenda = {
            Id: 0,
            Title: '',
            AgendaTitle: '',
            AgendaNumber: props.agendaNumber,
            EventLookupId: props.eventLookupId,
            ParentTopicId: props.parentTopicId,
            AllowPublicComments: false,
            SubTopics: [],
            Presenters: [],
            Documents: [],
            AgendaDate: new Date()
        };
        const agenda = McsUtil.isDefined(props.agenda) ? cloneDeep(props.agenda) : defaultAgenda;
        this.state = {
            agenda,
            agendaDate,
            useTime,
            presenter: this._getDefaultPresenter(agenda),
            agendaTime,
            waitingMessage: ''
        };
    }

    public render(): React.ReactElement<IAgendaFormProps> {
        const { isSubTopic } = this.props;
        const { useTime, agenda, agendaTime, agendaDate, waitingMessage, presenter } = this.state;
        const cansaveAgenda = this._canSaveAgenda(agenda);
        const canSavePresenter = this._canSavePresenter(presenter);
        const presenter_col_1_Width = '110px';
        return (<div className={css.combine(styles["d-flex"], styles["flex-column"], styles["justify-content-between"])}>
            <div className={styles["mb-3"]}>
                <div className={styles["container-fluid"]}>
                    {!isSubTopic &&
                        <div className={styles.row}>
                            <div className={styles["col-12"]}>
                                <Checkbox label="Use Time?" defaultChecked={useTime} onChange={this._onUseTimeChanged} />
                            </div>
                            <div className={css.combine(styles["col-6"])}>
                                <DatePicker
                                    label="Agenda Date/Time"
                                    className={styles.dateTimePicker}
                                    firstDayOfWeek={DayOfWeek.Monday}
                                    strings={datePickerStrings}
                                    showWeekNumbers={true}
                                    firstWeekOfYear={1}
                                    showMonthPickerAsOverlay={true}
                                    placeholder="Select start date..."
                                    ariaLabel="Select start date"
                                    minDate={this.props.minDate}
                                    maxDate={this.props.maxDate}
                                    onSelectDate={this._onAgendaDateChanged}
                                    value={agendaDate}
                                />
                            </div>
                            {useTime && <div className={css.combine(styles["col-6"])}>
                                <Timepicker time={agendaTime} onChange={this._onTimeChanged} />
                            </div>}
                        </div>
                    }

                    <div className={styles.row}>
                        {isSubTopic &&
                            <div className={styles["col-4"]}>
                                <TextField label="Number"
                                    type="number"
                                    name="AgendaNumber"
                                    onGetErrorMessage={(value) => /\d+/.test(value) ? undefined : 'Must be numberic'}
                                    data-validation={'\d+'}
                                    onChange={this._onAgendaTextChanged}
                                    value={agenda.AgendaNumber.toString()} />
                            </div>
                        }
                        <div className={isSubTopic ? styles["col-9"] : styles["col-12"]}>
                            <TextField label="Title" name="AgendaTitle"
                                value={agenda.AgendaTitle}
                                multiline
                                rows={3}
                                required
                                onChange={this._onAgendaTextChanged} />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles["col-12"]}>
                            <Checkbox label="Allow public comment?" className={styles["mt-2"]}
                                defaultChecked={agenda.AllowPublicComments}
                                onChange={this._onAllowPublicCommentChanged} />
                        </div>
                    </div>
                    {/* Presenter Form */}
                    <table className={css.combine(styles.table, styles["table-borderless"], styles["table-sm"], styles["mt-3"])}>
                        <thead>
                            <tr>
                                <th>Order</th>
                                <th>Presenter Name</th>
                                <th>Presenter Title</th>
                                <th>Organization</th>
                            </tr>
                        </thead>
                        <tbody style={{ minWidth: "20px" }}>
                            {McsUtil.isDefined(agenda) && McsUtil.isArray(agenda.Presenters) &&
                                sortBy(agenda.Presenters, a => a.SortNumber).map((p, index) => {
                                    return (<tr className={p.Id === presenter.Id ? css.combine(styles["bg-gray"]) : ''}>
                                        <td className={css.combine(styles["d-flex"], styles["justify-content-between"])}
                                            style={{ maxWidth: presenter_col_1_Width }}>
                                            <div className={css.combine(styles["d-flex"], styles["justify-content-between"])}>
                                                <IconButton iconProps={userRemoveIcon}
                                                    title="PresenterRemove"
                                                    ariaLabel="PresenterRemove"
                                                    name={'edit' + index.toString()}
                                                    className={styles["p-0"]}
                                                    onClick={() => this._onDeletePresenterClicked(p)}
                                                />
                                                <IconButton
                                                    iconProps={editContactIcon}
                                                    title="PresenterEdit"
                                                    ariaLabel="PresenterEdit"
                                                    className={styles["p-0"]}
                                                    name={'delete' + index.toString()}
                                                    onClick={() => this._onEditPresenterClicked(p)} />
                                            </div>
                                            <div>{p.SortNumber}</div>
                                        </td>
                                        <td>{p.PresenterName}</td>
                                        <td>{p.Title}</td>
                                        <td>{p.OrganizationName}</td>
                                    </tr>);
                                })
                            }
                        </tbody>
                        <tfoot>
                            <tr>
                                <td style={{ maxWidth: presenter_col_1_Width }}><TextField name="SortNumber"
                                    type='number'
                                    data-validation={'\d+'}
                                    value={presenter.SortNumber.toString()}
                                    required
                                    onChange={this._onPresenterTextChanged} /></td>
                                <td><TextField name="PresenterName"
                                    value={presenter.PresenterName}
                                    required
                                    onChange={this._onPresenterTextChanged} /></td>
                                <td><TextField name="Title" value={presenter.Title} onChange={this._onPresenterTextChanged} /></td>
                                <td><TextField name="OrganizationName" value={presenter.OrganizationName} onChange={this._onPresenterTextChanged} /></td>
                            </tr>
                            <tr>
                                <td colSpan={3}>
                                    <DefaultButton text={(presenter.Id !== 0 ? 'Edit' : 'Add') + ' Presenter'}
                                        disabled={!canSavePresenter}
                                        className={css.combine(styles["mr-2"], styles["bg-gray"])}
                                        style={canSavePresenter ? {} : { opacity: .4 }}
                                        onClick={this._onSavePresenterClicked} />
                                    <DefaultButton text="Clear Presenter"
                                        className={css.combine(styles["ml-2"], styles["bg-light"], styles["text-dark"])}
                                        onClick={this._clearPresenterClicked}
                                    />
                                </td>
                            </tr>
                        </tfoot>
                    </table>
                </div >
            </div>
            <div className={styles["mt-3"]}>
                <DefaultButton text="Save"
                    className={css.combine(styles["mr-2"], styles["bg-primary"], styles["text-white"])}
                    disabled={!cansaveAgenda}
                    style={cansaveAgenda ? {} : { opacity: .4 }}
                    onClick={this._onSaveClicked} />
                <DefaultButton text="Cancel"
                    className={css.combine(styles["ml-2"], styles["bg-light"], styles["text-dark"])}
                    onClick={this._dismisModal} />
            </div>
            <Waiting message={waitingMessage} />
        </div>
        );
    }

    private _canSaveAgenda = (agenda: IComponentAgenda): boolean => {
        if (McsUtil.isDefined(agenda) && McsUtil.isString(agenda.AgendaTitle) && !this._canSavePresenter(this.state.presenter)) {
            if (this.props.isSubTopic) {
                return McsUtil.isUnsignedInt(agenda.AgendaNumber);
            }
            return true;
        }
        return false;
    }

    private _canSavePresenter = (presenter: ISpPresenter): boolean => {
        return McsUtil.isUnsignedInt(presenter.SortNumber) && McsUtil.isString(presenter.PresenterName);
    }

    private _onTimeChanged = (newValue: string): void => {
        this.setState({ agendaTime: newValue });
    }

    private _onAgendaDateChanged = (date: Date | null | undefined): void => {
    }

    private _onUseTimeChanged = (ev?: any, checked?: boolean): void => {
        this.setState({ useTime: checked });
    }

    private _onAgendaTextChanged = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
        const { agenda } = this.state;
        const inputElement = ev.target as HTMLInputElement;
        const propertyName = inputElement.name;
        if (agenda.hasOwnProperty(propertyName)) {
            const validation = inputElement.getAttribute('data-validation');
            let isvalid = true;
            if (McsUtil.isString(validation) && value.length > 0) {
                isvalid = new RegExp(validation).test(value);
            }
            if (isvalid) {
                agenda[propertyName] = (value.length > 0) && /number/i.test(inputElement.type) ? parseInt(value) : value;
                this.setState({ agenda });
            }
        }
    }

    private _onPresenterTextChanged = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
        const { presenter } = this.state;
        const inputElement = ev.target as HTMLInputElement;
        const propertyName = inputElement.name;
        if (presenter.hasOwnProperty(propertyName)) {
            const validation = inputElement.getAttribute('data-validation');
            let isvalid = true;
            if (McsUtil.isString(validation) && value.length > 0) {
                isvalid = new RegExp(validation).test(value);
            }
            if (isvalid) {
                presenter[propertyName] = (value.length > 0) && /number/i.test(inputElement.type) ? parseInt(value) : value;
                this.setState({ presenter });
            }
        }
    }

    private _onAllowPublicCommentChanged = (ev?: any, checked?: boolean): void => {
        const agenda = { ...this.state.agenda };
        agenda.AllowPublicComments = checked;
        this.setState({ agenda });
    }

    private _onSaveClicked = (): void => {
        const { agenda } = this.state;
        const presenters = agenda.Presenters;
        let newPresenters: ISpPresenter[] = [];
        // find all presenters in props not in state
        this.setState({ waitingMessage: 'Adding or editing agenda.' });
        const deletedPresenters = McsUtil.isDefined(this.props.agenda) && McsUtil.isArray(this.props.agenda.Presenters) ?
            this.props.agenda.Presenters.filter(a => findIndex(presenters, p => p.Id === a.Id) < 0) : [];
        Promise.all([this._addPresenters(presenters.filter(a => a.Id < 1)),
        this._editPresenters(presenters.filter(a => a.Id > 0,
            this._deletePresenters(deletedPresenters)))])
            .then((responses) => {
                newPresenters = sortBy(responses[0].concat(responses[1]), a => a.SortNumber);
                const agendaPropertiesToUpdate = {
                    Title: agenda.AgendaTitle.length > 20 ? (agenda.AgendaTitle.substr(0, 20) + '...') : agenda.AgendaTitle,
                    AgendaTitle: agenda.AgendaTitle,
                    AgendaNumber: agenda.AgendaNumber,
                    AgendaDate: agenda.AgendaDate,
                    EventLookupId: this.props.eventLookupId,
                    PresentersLookupId: {
                        __metadata: {
                            type: "Collection(Edm.Int32)"
                        },
                        results: newPresenters.map(a => a.Id)
                    },
                    // PresentersLookupId: newPresenters.map(a => a.Id),
                    ParentTopicId: this.props.isSubTopic ? this.props.parentTopicId : null,
                    // AgendaDocumentsLookupId?: IMultipleLookupField; // we are not updating this on agenda
                    // Presenters?: Array<ISpPresenter>;
                    AllowPublicComments: agenda.AllowPublicComments
                };
                let promise = agenda.Id > 0 ? business.edit_Agenda(agenda.Id, agenda["odata.type"], agendaPropertiesToUpdate) :
                    business.add_Agenda(agendaPropertiesToUpdate);
                promise.then((result: IComponentAgenda) => {
                    result.Presenters = newPresenters;
                    result.Documents = [...agenda.Documents];
                    if (this.props.isSubTopic) {
                        result.SubTopics = [];
                    } else {
                        result.SubTopics = [...agenda.SubTopics];
                    }
                    this.setState({ agenda: result, waitingMessage: '' });
                    this.props.onChange(result, this.props.isSubTopic ? this.props.parentTopicId : undefined);
                });

            });
    }

    private _onEditPresenterClicked = (p: ISpPresenter): void => {
        this.setState({ presenter: p });
    }

    private _onDeletePresenterClicked = (p: ISpPresenter): void => {
        if (confirm("Do you want to remove presenter?")) {
            const { agenda } = this.state;
            const index = findIndex(agenda.Presenters, a => a.Id == p.Id);
            if (index > -1) {
                agenda.Presenters.splice(index, 1);
                this.setState({ agenda });
            }
        }
    }

    private _onSavePresenterClicked = (): void => {
        const { agenda, presenter } = this.state;
        if (presenter.Id !== 0) {
            const index = findIndex(agenda.Presenters, presenter.Id);
            if (index > -1) {
                agenda.Presenters[index] = presenter;
            }
        } else {
            const tempPresenter = { ...presenter };
            tempPresenter.Id = -1;
            agenda.Presenters.push(tempPresenter);
        }
        this.setState({ agenda, presenter: this._getDefaultPresenter(agenda) });
    }

    private _clearPresenterClicked = (): void => {
        this.setState({ presenter: this._getDefaultPresenter(this.state.agenda) });
    }

    private _dismisModal = (): void => {
        this.props.onCancel();
    }

    private _getDefaultPresenter = (agenda): ISpPresenter => {
        let sortNumber = 1;
        if (McsUtil.isDefined(agenda)) {
            if (McsUtil.isArray(agenda.Presenters) && agenda.Presenters.length > 0) {
                sortNumber = agenda.Presenters[agenda.Presenters.length - 1].SortNumber + 1;
            } else {
                sortNumber = 1;
            }
        }
        return {
            Id: 0,
            Title: '',
            SortNumber: sortNumber,
            PresenterName: '',
            OrganizationName: '',
        };
    }

    private _addPresenters = (presentersToAdd: ISpPresenter[]): Promise<ISpPresenter[]> => {
        return new Promise((resolve, reject) => {
            if (McsUtil.isArray(presentersToAdd) && presentersToAdd.length > 0) {
                Promise.all(presentersToAdd.map(a => business.add_Presenter({
                    Title: a.Title,
                    SortNumber: a.SortNumber,
                    PresenterName: a.PresenterName,
                    OrganizationName: a.OrganizationName
                }))).then((responses: ISpPresenter[]) => {
                    resolve(responses);
                });
            } else {
                resolve([]);
            }
        });
    }

    private _editPresenters = (presentersToEdit: ISpPresenter[]): Promise<ISpPresenter[]> => {
        return new Promise((resolve, reject) => {
            if (McsUtil.isArray(presentersToEdit) && presentersToEdit.length > 0) {
                Promise.all(presentersToEdit.map(a => business.edit_Presenter(
                    a.Id,
                    a["odata.type"],
                    {
                        Title: a.Title,
                        SortNumber: a.SortNumber,
                        PresenterName: a.PresenterName,
                        OrganizationName: a.OrganizationName
                    }))).then((responses: ISpPresenter[]) => {
                        resolve(responses);
                    });
            } else {
                resolve([]);
            }
        });
    }

    private _deletePresenters = (presentersToAdd: ISpPresenter[]): Promise<void> => {
        return new Promise((resolve, reject) => {
            if (McsUtil.isArray(presentersToAdd) && presentersToAdd.length > 0) {
                Promise.all(presentersToAdd.map(a => business.delete_Presenter(a.Id)))
                    .then(() => {
                        resolve();
                    });
            } else {
                resolve();
            }
        });
    }
}