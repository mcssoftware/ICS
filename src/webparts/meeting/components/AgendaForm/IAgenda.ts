import { ISpPresenter} from "../../../../interface/spmodal";
import { IComponentAgenda } from "../../../../business/transformAgenda";
import { InformationalType } from "../../../../controls/informational";

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
  waitingMessage: string;
  message: string;
  messageType: InformationalType;
  orderChanged: boolean;
}

export interface IAgendaFormProps {
  eventLookupId: number;
  agenda: IComponentAgenda;
  isSubTopic: boolean;
  parentTopicId?: number;
  agendaNumber: number;
  minDate?: Date;
  maxDate?: Date;
  onChange: (topic: IComponentAgenda, parentTopicId?: number) => void;
  onCancel: () => void;
}

export interface IAgendaFormState {
  useTime?: boolean;
  agenda: IComponentAgenda;
  agendaDate: Date;
  presenter: ISpPresenter;
  agendaTime: string;
  waitingMessage: string;
}