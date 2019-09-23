import { combine, dateAdd, DateAddInterval } from "@pnp/common";
import { Logger, LogLevel } from "@pnp/logging";

export class McsUtil {
    public static isDefined(n: any): boolean {
        return typeof n !== "undefined" && n !== null;
    }

    public static isArray(n: any): boolean {
        return typeof n !== "undefined" && n !== null && Array.isArray(n);
    }

    public static isString(n: any): boolean {
        return typeof n === "string" && n.length > 0;
    }

    public static isNumeric(n: any): boolean {
        return typeof n !== "undefined" && n !== null && !isNaN(parseFloat(n)) && isFinite(n);
    }

    public static isUnsignedInt(n: any): boolean {
        return typeof n !== "undefined" && n !== null && !isNaN(parseInt(n, 10)) && isFinite(n) && n > -1;
    }

    public static toNumber(n: any): number {
        return McsUtil.isUnsignedInt(n) ? parseInt(n, 10) : 0;
    }

    public static isFunction(o: any): boolean {
        return toString.call(o) === "[object Function]";
    }

    public static trim(a: string): string {
        if (McsUtil.isString(a)) {
            return a.trim();
        }
        return a;
    }

    public static padNumber(n: number, length: number): string {
        let s: string = n.toString();
        while (s.length < (length || 2)) { s = "0" + s; }
        return s;
    }

    public static getString(value: number, length: number): string {
        let stringValue: string = value.toString();
        for (let i: number = stringValue.length; i < length; i++) {
            stringValue = "0" + stringValue;
        }
        return stringValue;
    }

    public static isNumberString(value: string): boolean {
        return /^\d+$/.test(value);
    }

    public static isGuid(stringToTest: string): boolean {
        if (McsUtil.isString(stringToTest)) {
            if (stringToTest[0] === "{") {
                stringToTest = stringToTest.substring(1, stringToTest.length - 1);
            }
            const regexGuid: RegExp = /^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/gi;
            return regexGuid.test(stringToTest);
        }
        return false;
    }

    public static getApiErrorMessage(err: any): string {
        if (McsUtil.isDefined(err)) {
            try {
                return err.data.responseBody["odata.error"].message.value;
            } catch (e) {
                return err.toString();
            }
        }
        return "";
    }

    public static combinePaths(...paths: string[]): string {
        return combine(...paths);
    }

    public static dateAdd(date: Date, interval: DateAddInterval, units: number): Date {
        return dateAdd(date, interval, units);
    }

    public static makeAbsUrl(url: string): string {
        if (url.length > 0 && "/" === url.substr(0, 1)) {
            url = window.location.protocol + "//" + window.location.host + url;
        }
        return url;
    }

    public static getRelativeUrl(url: string): string {
        let relativeUrl: string;
        const protocolIndex: number = url.indexOf("//");
        if (-1 !== protocolIndex) {
            const hostIndex: number = url.indexOf("/", protocolIndex + 2);
            if (-1 !== hostIndex) {
                relativeUrl = url.substr(hostIndex);
            } else {
                relativeUrl = "/";
            }
        } else {
            relativeUrl = url;
        }
        return relativeUrl;
    }

    public static parseHtmlEntities(value: string): string {
        if (McsUtil.isString(value)) {
            const includeRegExp: RegExp = new RegExp("&#([0-9]{1,3});", "gi");
            while (true) {
                // tslint:disable-next-line:prefer-const
                let regExpResult: RegExpExecArray | null = includeRegExp.exec(value);
                if (regExpResult) {
                    const fullMatch: string = regExpResult[0];
                    const tokenName: string = regExpResult[1];
                    const num: number = parseInt(tokenName, 10); // read num as normal number
                    value = value.replace(new RegExp(fullMatch, "gi"), String.fromCharCode(num));
                } else {
                    break;
                }
            }
        }
        return value;
    }

    // Change: 2018-9-24 - MW - Check length first.
    public static encodeHtmlEntity(value: string): string {
        const buf: string[] = [];
        if (McsUtil.isString(value)) {
            for (let i: number = value.length - 1; i >= 0; i--) {
                buf.unshift(["&#", value[i].charCodeAt(0), ";"].join(""));
            }
        }
        return buf.join("");
    }

    public static unescapeProperly(b: string): string {
        let a: string = null;
        try {
            a = decodeURIComponent(b);
        } catch (c) {
            // TODO manual parsing
            a = b;
        }
        return a;
    }

    public static getNextDocumentVersion(documentVersion: number, publish: boolean): number {
        let integerPart: number = Math.floor(documentVersion);
        let decimalPart: number = 0;
        if (!publish) {
            decimalPart = 0;
            integerPart = integerPart + 1;
        } else {
            const countDecimals: number = documentVersion.toString().split(".")[1].length || 0;
            if (countDecimals === 1) {
                decimalPart = Math.floor(documentVersion * 10) % 10;
                if (decimalPart === 9) {
                    decimalPart = decimalPart + 1;
                }
            }
            if (countDecimals === 2) {
                decimalPart = Math.floor(documentVersion * 100) % 100;
                if ((decimalPart + 1) % 10 === 0) {
                    decimalPart = decimalPart + 1;
                }
            }
            decimalPart = decimalPart + 1;
        }
        return parseFloat(`${integerPart}.${decimalPart}`);
    }

    public static getFileExtension(fileName: string): string {
        const tempFileName: string = McsUtil.isString(fileName) ? fileName : "";
        const fileExtensionRegex: RegExp = /^.*\.([^\.]*)$/;
        return tempFileName.replace(fileExtensionRegex, "$1").toLowerCase();
    }

    public static isWordDocument(fileName: string): boolean {
        const fileExtension: string = McsUtil.getFileExtension(fileName);
        return (/(doc|docx)$/i).test(fileExtension);
    }

    // tslint:disable:no-bitwise
    public static escape(f: string): string {
        let c: string = "";
        let b: number;
        let d: number = 0;
        const k: string = " \"%<>'&";
        if (typeof f === "undefined") { return ""; }
        for (d = 0; d < f.length; d++) {
            let a: number = f.charCodeAt(d);
            const e: string = f.charAt(d);
            if (e === "#" || e === "?") {
                c += f.substr(d);
                break;
            }
            if (e === "&") {
                c += e;
                continue;
            }
            if (a <= 127) {
                if (a >= 97 && a <= 122 || a >= 65 && a <= 90 || a >= 48 && a <= 57 || true && a >= 32 && a <= 95 && k.indexOf(e) < 0) {
                    c += e;
                }
                else if (a <= 15) {
                    c += "%0" + a.toString(16).toUpperCase();
                }
                else if (a <= 127) {
                    c += "%" + a.toString(16).toUpperCase();
                }
            } else if (a <= 2047) {
                b = 192 | a >> 6;
                c += "%" + b.toString(16).toUpperCase();
                b = 128 | a & 63;
                c += "%" + b.toString(16).toUpperCase();
            } else if ((a & 64512) !== 55296) {
                b = 224 | a >> 12;
                c += "%" + b.toString(16).toUpperCase();
                b = 128 | (a & 4032) >> 6;
                c += "%" + b.toString(16).toUpperCase();
                b = 128 | a & 63;
                c += "%" + b.toString(16).toUpperCase();
            } else if (d < f.length - 1) {
                a = (a & 1023) << 10;
                d++;
                const j: number = f.charCodeAt(d);
                a |= j & 1023;
                a += 65536;
                b = 240 | a >> 18;
                c += "%" + b.toString(16).toUpperCase();
                b = 128 | (a & 258048) >> 12;
                c += "%" + b.toString(16).toUpperCase();
                b = 128 | (a & 4032) >> 6;
                c += "%" + b.toString(16).toUpperCase();
                b = 128 | a & 63;
                c += "%" + b.toString(16).toUpperCase();
            }
        }
        return c;
    }

    public static unescape(b: string): string {
        let a: string = null;
        try {
            a = decodeURIComponent(b);
            // tslint:disable-next-line:no-empty
        } catch (c) {
        }
        return a;
    }

    // public static formatDate(datevalue: string | Date, format: string): string {
    //     if (McsUtil.isDefined(datevalue)) {
    //         if (datevalue instanceof Date && !isNaN(datevalue.valueOf())) {
    //             return datevalue.format(format);
    //         }
    //         if (McsUtil.isString(datevalue)) {
    //             const d: Date = new Date(datevalue);
    //             if (d instanceof Date && !isNaN(d.valueOf())) {
    //                 return d.format(format);
    //             }
    //         }
    //     }
    //     return "";
    // }

    public static chunkArray<T>(myArray: T[], chunkSize: number): T[][] {
        let index: number = 0;
        const arrayLength: number = myArray.length;
        const tempArray: T[][] = [];

        for (index = 0; index < arrayLength; index += chunkSize) {
            const myChunk: T[] = myArray.slice(index, index + chunkSize);
            tempArray.push(myChunk);
        }
        return tempArray;
    }

    public static binarySearch(sortedElements: any[], searchElement: string | number, propertyName: string): number {
        let minIndex: number = 0;
        let maxIndex: number = sortedElements.length - 1;
        let currentIndex: number;
        let currentElement: any;
        if (McsUtil.isDefined(searchElement)) {
            while (minIndex <= maxIndex) {
                // tslint:disable-next-line:no-bitwise
                currentIndex = (minIndex + maxIndex) / 2 | 0;
                currentElement = sortedElements[currentIndex];
                const tempValue: any = currentElement[propertyName];
                if (tempValue == searchElement) {
                    return currentIndex;
                } else {
                    if (tempValue < searchElement) {
                        minIndex = currentIndex + 1;
                    } else {
                        maxIndex = currentIndex - 1;
                    }
                }
            }
        }
        return -1;
    }

    public static handleLogError = (err: any, callback?: () => void) => {
        Logger.log({
            data: { err },
            level: LogLevel.Error,
            message: err
        });
        if (McsUtil.isFunction(callback)) {
            callback();
        }
    }

    public static convertUtcDateToLocalDate = (date: Date): Date => {
        var newDate = new Date(date.getTime() + date.getTimezoneOffset() * 60 * 1000);
        var offset = date.getTimezoneOffset() / 60;
        var hours = date.getHours();
        newDate.setHours(hours - offset);
        return newDate;
    }

    public static createDownloadLink(fileName: string, data: any): void {
        if (window.navigator.msSaveOrOpenBlob) {
            var fileData = [data];
            var blobObject = new Blob(fileData);
            window.navigator.msSaveOrOpenBlob(blobObject, fileName);

        } else {
            var blob = new Blob([data]);
            var link = <any>document.getElementById("noticePreview");
            link.href = window.URL.createObjectURL(blob);
            link.download = fileName;
            link.click();
            //}
        }
    }

}