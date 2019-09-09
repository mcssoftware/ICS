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
    let messageBarType = MessageBarType.info;
    if (props.type === InformationalType.Error) {
        messageBarType = MessageBarType.error;
    }
    return (
        <MessageBar messageBarType={messageBarType} isMultiline={true} hidden={props.type === InformationalType.none}>{props.message}
        </MessageBar>
    );
};

export { part as Informational };