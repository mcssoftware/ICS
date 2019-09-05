import { ISpAgendaTopic, ISpPresenter, ISpEventMaterial, IListItem } from "../../../../interface/spmodal";
import { IComponentAgenda } from "../../../../business/transformAgenda";

export enum AgendaPanelType {
  uploadDocument,
  topic,
  subtopic,
}

export interface IAgendaProps {
  eventLookupId: number;
  minDate?: Date;
  maxDate?: Date;
}

export interface IAgendaState {
  agendaItems: IComponentAgenda[];
  selectedAgendaItem: IComponentAgenda;
  showPanel: boolean;
  panelType?: AgendaPanelType;
  panelHeaderText: string;
  panelItem: any;
}

export interface IAgendaFormProps {
  eventLookupId: number;
  agenda: IComponentAgenda;
  isSubTopic: boolean;
  parentTopicId?: number;
  agendaNumber: number;
  minDate?: Date;
  maxDate?: Date;
  closeModal?: (newAgenda?: IComponentAgenda) => void;
}

export interface IAgendaFormState {
  useTime?: boolean;
  agenda: IComponentAgenda;
  agendaDate: Date;
  presenter: ISpPresenter;
  agendaTime: string;
}