import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';

export enum InformationalType {
    Error = 0,
    Info,
    none
}

export interface IInformationalProps {
    message: string;
    type: InformationalType;
}

const part: React.SFC<IInformationalProps> = (props) => {
    let messageBarType = MessageBarType.success;
    if (props.type === InformationalType.Error) {
        messageBarType = MessageBarType.severeWarning;
    }
    return (<div style={{ paddingTop: '10px' }}>
        {props.type !== InformationalType.none && <MessageBar messageBarType={messageBarType} isMultiline={true} >{props.message}</MessageBar>}
    </div>);
};

export { part as Informational };