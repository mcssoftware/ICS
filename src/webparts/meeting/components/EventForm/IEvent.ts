import { ISpEvent } from "../../../../interface/spmodal";
import { IDbCommittee } from "../../../../interface/dbmodal";

export interface IEventProps {
  event: ISpEvent;
  committees: IDbCommittee[];
}

export interface IEventState {
  event: ISpEvent;
  startDate?: Date;
  endDate?: Date;
  selectedState: any;
  isDirty: boolean;
}
