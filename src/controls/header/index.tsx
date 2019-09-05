import * as React from 'react';
import styles from './header.module.scss';

export enum EnumTextAlign {
    center = 0,
    left,
    right
}
export interface IHeaderPartProps {
    title: string;
    fontColor?: string;
    backgroundColor?: string;
    fontSize?: string;
    textAlign: EnumTextAlign;
}
const part: React.SFC<IHeaderPartProps> = (props) => {
    const webpartheaderCss: React.CSSProperties = {};
    const row: React.CSSProperties = {};
    const headerText: React.CSSProperties = {};
    const tempStyles = {
        webpartheaderCss,
        row,
        headerText
    };

    if (typeof props.backgroundColor === "string") {
        tempStyles.webpartheaderCss.backgroundColor = props.backgroundColor;
    }
    if (typeof props.fontColor === "string") {
        tempStyles.webpartheaderCss.color = props.fontColor;
    }
    if (typeof props.fontSize === "string") {
        tempStyles.headerText.fontSize = props.fontSize;
    }
    if (EnumTextAlign.left === props.textAlign) {
        tempStyles.headerText.float = 'left';
    }
    if (EnumTextAlign.right === props.textAlign) {
        tempStyles.headerText.float = 'right';
    }
    if (EnumTextAlign.center === props.textAlign) {
        tempStyles.headerText.float = 'none';
    }

    return (<div className={styles.webpartheader} style={tempStyles.webpartheaderCss}>
        <div className={styles.row} style={tempStyles.row}>
            {props.title.length > 0 &&
                <span className={styles.headerText} style={tempStyles.headerText}>{props.title}</span>
            }
            {props.children}
        </div>
    </div>);
};

export { part as Header };