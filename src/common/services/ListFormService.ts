import { ControlMode } from '../datatypes/ControlMode';
import { IFieldSchema, IField, IRowData } from './datatypes/RenderListData';
import { IListFormService } from './IListFormService';
import { Web, RenderListDataOptions } from '@pnp/sp';
import { String } from 'typescript-string-operations';

export class ListFormService implements IListFormService {

    /**
     * Gets the schema for all relevant fields for a specified SharePoint list form.
     *
     * @param webUrl The absolute Url to the SharePoint site.
     * @param listUrl The server-relative Url to the SharePoint list.
     * @param viewId The type of form (Display, New, Edit)
     * @returns Promise object represents the array of field schema for all relevant fields for this list form.
     */
    public async getViewFields(webUrl: string, listName: string, viewId: string): Promise < IField[] > {
        return new Promise < IField[] > (async(resolve, reject) => {
            let web = new Web(webUrl);
            const list = web
                .lists
                .getByTitle(listName);
            const {HtmlSchemaXml: ViewXml} = await list
                .getView(viewId)
                .get();
            await list
                .renderListDataAsStream({ViewXml, RenderOptions: RenderListDataOptions.ListSchema})
                .then((data) => {
                    resolve(data.Field);
                })
                .catch((error) => {
                    reject(error);
                });
        });
    }

    /**
     * Gets the schema for all relevant fields for a specified SharePoint list form.
     *
     * @param webUrl The absolute Url to the SharePoint site.
     * @param listUrl The server-relative Url to the SharePoint list.
     * @param viewId The type of form (Display, New, Edit)
     * @returns Promise object represents the array of field schema for all relevant fields for this list form.
     */
	public async getViewItems(
		webUrl: string,
		listName: string,
		viewId: string,
		overrideParameters?: any,
		paging?: string
	): Promise<IRowData> {
		return new Promise<IRowData>(async (resolve, reject) => {
			let web = new Web(webUrl);
			const list = web.lists.getByTitle(listName);
   const { HtmlSchemaXml: ViewXml } = await list.getView(viewId).get();
            
			await list
				.renderListDataAsStream(
					{
						ViewXml,
						AllowMultipleValueFilterForTaxonomyFields: true,
						RenderOptions: RenderListDataOptions.ListData,
						Paging: paging
					},
					overrideParameters
				)
				.then((data) => {
					resolve(data);
				})
				.catch((error) => {
					reject(error);
				});
		});
	}
    /**
     * Gets the schema for all relevant fields for a specified SharePoint list form.
     *
     * @param webUrl The absolute Url to the SharePoint site.
     * @param listName The name of the list.
     * @param formType The type of form (Display, New, Edit)
     * @returns Promise object represents the array of field schema for all relevant fields for this list form.
     */
    public getFieldSchemasForForm(webUrl: string, listName: string, formType: ControlMode): Promise < IFieldSchema[] > {
        return new Promise < IFieldSchema[] > ((resolve, reject) => {
            let web = new Web(webUrl);

            web
                .lists
                .getByTitle(listName)
                // tslint:disable-next-line:max-line-length
                .renderListDataAsStream({RenderOptions: RenderListDataOptions.ClientFormSchema, ViewXml: '<View><ViewFields><FieldRef Name="ID"/></ViewFields></View>'})
                .then((data) => {
                    const form = (formType === ControlMode.New)
                        ? data.ClientForms.New
                        : data.ClientForms.Edit;
                    resolve(form[Object.keys(form)[0]]);
                })
                .catch((error) => {
                    reject(this.getErrorMessage(webUrl, error));
                });

        });
    }

    /**
     * Returns an error message based on the specified error object
     * @param error : An error string/object
     */
    private getErrorMessage(webUrl: string, error: any): string {
        let errorMessage: string = error.statusText
            ? error.statusText
            : error.statusMessage
                ? error.statusMessage
                : error;
        const serverUrl = `{window.location.protocol}//{window.location.hostname}`;
        const webServerRelativeUrl = webUrl.replace(serverUrl, '');

        if (error.status === 403) {
            errorMessage = String.Format(
            // tslint:disable-next-line:max-line-length
            'You do not have access to the previously configured web url \'{0}\'. Either leav' +
                    'e the WebPart properties as is or select another web url.',
            webServerRelativeUrl);
        } else if (error.status === 404) {
            errorMessage = String.Format(
            // tslint:disable-next-line:max-line-length
            'The previously configured web url \'{0}\' is not found anymore. Either leave the' +
                    ' WebPart properties as is or select another web url.',
            webServerRelativeUrl);
        }
        return errorMessage;
    }

}
