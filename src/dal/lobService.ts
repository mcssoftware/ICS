import { AadHttpClient, IHttpClientOptions, HttpClientResponse } from "@microsoft/sp-http";
import { ServiceScope } from "@microsoft/sp-core-library";
import { IODataParser, BlobParser, ODataDefaultParser } from "../utility/parser";
import { McsUtil } from "../utility/helper";
// import { McsUtil, ODataDefaultParser, IODataParser, BlobParser } from "mcs-lms-core";

export interface ILobService {
    getData(serviceScope: ServiceScope, url: string, parser?: IODataParser<any>): Promise<any>;
    getBlob(serviceScope: ServiceScope, url: string, parser?: BlobParser): Promise<any>;
    postData(serviceScope: ServiceScope, url: string, data: any, contentType?: string, parser?: IODataParser<any>, headers?: Headers): Promise<any>;
    putData(serviceScope: ServiceScope, url: string, data: any, contentType?: string, parser?: IODataParser<any>): Promise<any>;
}

const azureServiceUrl:string = 'https://WYOLEG.GOV/LsoOffice365Service';

class LobService implements ILobService {

    public getData(serviceScope: ServiceScope, url: string, parser: IODataParser<any> = new ODataDefaultParser()): Promise<any> {
        return new Promise((resolve, reject) => {
            // create an AadHttpClient object to consume the 3rd party API
            const aadClient: AadHttpClient = new AadHttpClient(
                serviceScope,
                azureServiceUrl
            );
            const requestHeaders: Headers = new Headers();
            requestHeaders.append("Accept", "application/json");

            const requestOptions: IHttpClientOptions = {
                headers: requestHeaders,
            };
            aadClient.get(
                url,
                AadHttpClient.configurations.v1,
                requestOptions
            ).then((httpResponse: HttpClientResponse) => {
                if (httpResponse.ok) {
                    return parser.parse(httpResponse as any);
                } else {
                    return Promise.reject(httpResponse.status);
                }
            }).then((value: any) => {
                resolve(value);
            }).catch((err) => McsUtil.handleLogError(err, () => { reject(err); }));
        });
    }

    public getBlob(serviceScope: ServiceScope, url: string, parser: BlobParser = new BlobParser()): Promise<any> {
        return new Promise((resolve, reject) => {
            // create an AadHttpClient object to consume the 3rd party API
            const aadClient: AadHttpClient = new AadHttpClient(
                serviceScope,
                azureServiceUrl
            );
            const requestHeaders: Headers = new Headers();
            requestHeaders.append("Accept", "application/json");

            const requestOptions: IHttpClientOptions = {
                headers: requestHeaders,
            };
            aadClient.get(
                url,
                AadHttpClient.configurations.v1,
                requestOptions
            ).then((httpResponse: HttpClientResponse) => {
                if (httpResponse.ok) {
                    return parser.parse(httpResponse as any);
                } else {
                    return Promise.reject(httpResponse.status);
                }
            }).then((value: any) => {
                resolve(value);
            }).catch((err) => McsUtil.handleLogError(err, () => { reject(err); }));
        });
    }

    public postData(serviceScope: ServiceScope, url: string, data: any, contentType: string = "application/json",
        parser: IODataParser<any> = new ODataDefaultParser(), headers?: Headers): Promise<any> {
        return new Promise((resolve, reject) => {
            // create an AadHttpClient object to consume the 3rd party API
            const aadClient: AadHttpClient = new AadHttpClient(
                serviceScope,
                azureServiceUrl
            );
            let requestHeaders: Headers = new Headers();
            requestHeaders.append("Accept", "application/json");
            requestHeaders.append("Content-Type", McsUtil.isString(contentType) ? contentType : "application/json");
            if (McsUtil.isDefined(headers)) {
                requestHeaders = headers;
            }
            const requestOptions: IHttpClientOptions = {
                headers: requestHeaders,
                body: JSON.stringify(data),
            };

            aadClient.post(
                url,
                AadHttpClient.configurations.v1,
                requestOptions
            ).then((httpResponse: HttpClientResponse) => {
                if (httpResponse.ok) {
                    return parser.parse(httpResponse as any);
                } else {
                    return Promise.reject(httpResponse.status);
                }
            }).then((value: any) => {
                resolve(value);
            }).catch((err) => McsUtil.handleLogError(err, () => {
                reject(err);
            }));
        });
    }

    public putData(serviceScope: ServiceScope, url: string, data: any, contentType: string = "application/json",
        parser: IODataParser<any> = new ODataDefaultParser()): Promise<any> {
        return new Promise((resolve, reject) => {
            // create an AadHttpClient object to consume the 3rd party API
            const aadClient: AadHttpClient = new AadHttpClient(
                serviceScope,
                azureServiceUrl
            );
            const requestHeaders: Headers = new Headers();
            requestHeaders.append("Accept", "application/json");
            requestHeaders.append("Content-Type", McsUtil.isString(contentType) ? contentType : "application/json");

            const requestOptions: IHttpClientOptions = {
                headers: requestHeaders,
                body: JSON.stringify(data),
                method: "PUT"
            };

            aadClient.fetch(
                url,
                AadHttpClient.configurations.v1,
                requestOptions
            ).then((httpResponse: HttpClientResponse) => {
                if (httpResponse.ok) {
                    return parser.parse(httpResponse as any);
                } else {
                    return Promise.reject(httpResponse.status);
                }
            }).then((value: any) => {
                resolve(value);
            }).catch((err) => McsUtil.handleLogError(err, () => { reject(err); }));
        });
    }
}

const lobService: ILobService = new LobService();
export default lobService;
