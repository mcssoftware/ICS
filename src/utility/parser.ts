import { HttpRequestError } from "./HttpRequestError";
import { hOP, isFunc } from "@pnp/common";

export interface IODataParser<T> {
    hydrate?: (d: any) => T;
    parse(r: any): Promise<T>;
}

export abstract class ODataParserBase<T> implements IODataParser<T> {

    public parse(r: any): Promise<T> {

        return new Promise<T>((resolve, reject) => {
            if (this.handleError(r, reject)) {
                this.parseImpl(r, resolve, reject);
            }
        });
    }

    protected parseImpl(r: any, resolve: (value?: T | PromiseLike<T>) => void, reject: (reason?: Error) => void): void {
        if ((r.headers.has("Content-Length") && parseFloat(r.headers.get("Content-Length")!) === 0) || r.status === 204) {
            resolve({} as T);
        } else {

            // patch to handle cases of 200 response with no or whitespace only bodies (#487 & #545)
            r.text()
                .then((txt) => txt.replace(/\s/ig, "").length > 0 ? JSON.parse(txt) : {})
                .then((json) => resolve(this.parseODataJSON<T>(json)))
                .catch((e) => reject(e));
        }
    }

    /**
     * Handles a response with ok === false by parsing the body and creating a ProcessHttpClientResponseException
     * which is passed to the reject delegate. This method returns true if there is no error, otherwise false
     *
     * @param r Current response object
     * @param reject reject delegate for the surrounding promise
     */
    protected handleError(r: any, reject: (err?: Error) => void): boolean {
        if (!r.ok) {
            HttpRequestError.init(r).then(reject);
        }

        return r.ok;
    }

    /**
     * Normalizes the json response by removing the various nested levels
     *
     * @param json json object to parse
     */
    protected parseODataJSON<U>(json: any): U {
        let result: any = json;
        if (hOP(json, "d")) {
            if (hOP(json.d, "results")) {
                result = json.d.results;
            } else {
                result = json.d;
            }
        } else if (hOP(json, "value")) {
            result = json.value;
        }
        return result;
    }
}

// tslint:disable:max-classes-per-file
export class ODataDefaultParser extends ODataParserBase<any> {
}

export class TextParser extends ODataParserBase<string> {

    protected parseImpl(r: any, resolve: (value: any) => void): void {
        r.text().then(resolve);
    }
}

export class BlobParser extends ODataParserBase<Blob> {

    protected parseImpl(r: any, resolve: (value: any) => void): void {
        r.blob().then(resolve);
    }
}

export class JSONParser extends ODataParserBase<any> {

    protected parseImpl(r: any, resolve: (value: any) => void): void {
        r.json().then(resolve);
    }
}

export class BufferParser extends ODataParserBase<ArrayBuffer> {

    protected parseImpl(r: any, resolve: (value: any) => void): void {
        if (isFunc(r.arrayBuffer)) {
            r.arrayBuffer().then(resolve);
        } else {
            (r as any).buffer().then(resolve);
        }
    }
}