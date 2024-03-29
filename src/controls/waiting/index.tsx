import * as React from 'react';
import { Spinner, Label, SpinnerSize, Stack, IStackProps } from 'office-ui-fabric-react';
import styles from "./waiting.module.scss";

export interface IHeaderPartProps {
    message: string;
}
const part: React.SFC<IHeaderPartProps> = (props) => {
    const rowProps: IStackProps = { horizontal: true, verticalAlign: 'center' };
    const tokens = {
        sectionStack: {
            childrenGap: 10
        },
        spinnerStack: {
            childrenGap: 20
        }
    };
    const isvisible = typeof props.message === "string" && props.message.length > 0;
    return (<div>
        {isvisible && <div className={styles.waiting}>
            <div className={styles.stackContainer}>
                <Stack {...rowProps} tokens={tokens.spinnerStack} className={styles.stack}>
                    <Label>{props.message}</Label>
                    <Spinner size={SpinnerSize.large} />
                </Stack>
            </div>
        </div>}
    </div>
    );
};

export { part as Waiting };