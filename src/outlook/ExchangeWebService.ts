// Copyright (c) 2014-2017 Gautam Singh
// MIT License; Full text at - https://github.com/gautamsi/ews-javascript-api/blob/master/LICENSE
// Packages license notices pending

import * as ews from "../ExchangeWebService";

declare var Office: { context: { mailbox: { makeEwsRequestAsync: (data: string, callback: (asyncResult: { value: string, asyncContext?: any, status?: string, error?: Error }) => void) => void } } };


export class XHROutlook implements ews.IXHRApi {
	xhr(xhroptions: ews.IXHROptions, progressDelegate?: (progressData: ews.IXHRProgress) => void): Promise<XMLHttpRequest> {
		return <any>new Promise<XMLHttpRequest>((resolve, reject) => {
			Office.context.mailbox.makeEwsRequestAsync(xhroptions.data, (result) => {
				let res: XMLHttpRequest = <any>{
					status: 200,
					responseText: result.value,
					getAllResponseHeaders: () => [],
					getResponseHeader: (str: string) => ""
				}
				if (result.status === 'succeeded') {
					resolve(res)
				}
				else {
					(<any>res).message = result.error.message;
					(<any>res).status = 500;
					reject(res);
				}
			})
		});
	}

	xhrStream(xhroptions: ews.IXHROptions, progressDelegate: (progressData: ews.IXHRProgress) => void): Promise<XMLHttpRequest> {

		return <any>new Promise((resolve, reject) => {
			reject(new Error("not implemented/ not used"))
		});
	}

	disconnect() {

	}

	get apiName(): string {
		return "outlook";
	}

	constructor() {
	}
}

ews["XHROutlook"] = XHROutlook;

export function ConfigureForOutlook() {
	ews.ConfigurationApi.ConfigureXHR(new XHROutlook());
}

ews["ConfigureForOutlook"] = ConfigureForOutlook;

//ews.ConfigurationApi.ConfigureXHR(new XHROutlook());

window["EwsJS"] = ews;