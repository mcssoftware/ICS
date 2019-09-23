import { ISpEvent } from "../../../../interface/spmodal";
import { IDbCommittee } from "../../../../interface/dbmodal";
import { InformationalType } from "../../../../controls/informational";

export interface IEventProps {
  event: ISpEvent;
  committees: IDbCommittee[];
  onChange: () => void;
}

export interface IEventState {
  event: ISpEvent;
  startDate?: Date;
  endDate?: Date;
  selectedState: any;
  isDirty: boolean;
  waitingMessage: string;
  message: string;
  messageType: InformationalType;
  publishPanelOpen: boolean;
}
