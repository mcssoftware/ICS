export class HttpRequestError extends Error {
    public isHttpRequestError: boolean = true;
    constructor(message: string, public response: any, public status: number = response.status, public statusText: string = response.statusText) {
        super(message);
    }
    public static init(r: any): Promise<HttpRequestError> {
        return r.clone().text().then((t) => {
            return new HttpRequestError(`Error making HttpClient request in queryable [${r.status}] ${r.statusText} ::> ${t}`, r.clone());
        });
    }
}
