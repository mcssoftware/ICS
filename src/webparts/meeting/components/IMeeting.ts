import { ISpEvent } from "../../../interface/spmodal";
import { IDbCommittee } from "../../../interface/dbmodal";

export interface IMeetingProps {
  description: string;
}

export interface IMeetingState {
  event: ISpEvent;
  committees: IDbCommittee[];
  isLoaded: boolean;
  selectedTab: string;
  minDate?: Date;
  maxDate?: Date;
}
