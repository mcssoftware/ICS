import * as React from 'react';
import { Spinner } from 'office-ui-fabric-react';

export interface IHeaderPartProps {
}
const part: React.SFC<IHeaderPartProps> = (props) => {

    return (
        <div>
            <Spinner label={"Loading options..."} />
        </div>
    );
};

export { part as Loading };