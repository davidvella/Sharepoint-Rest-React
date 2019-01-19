import { IDocumentService } from './IDocumentService';
import { Web, FileAddResult, ChunkedFileUploadProgressData, ListItemFormUpdateValue, CamlQuery } from '@pnp/sp';
import { IFieldSchema } from './datatypes/RenderListData';
import { String } from 'typescript-string-operations';

export class DocumentService implements IDocumentService {
    /**
     * Gets a file by server relative url
     *
     * @param fileRelativeUrl The server relative path to the file (including /sites/ if applicable)
     * @returns The new File and the raw response.
     */
	public async getFileFromListAsBlob(webUrl: string, fileRelativeUrl: string): Promise<Blob> {
		return new Promise<Blob>(async (resolve, reject) => {
			let web = new Web(webUrl);
			await web
				.getFileByServerRelativeUrl(fileRelativeUrl)
				.getBlob()
				.then((res) => {
					resolve(res);
				})
				.catch((error) => {
					reject(error);
				});
		});
	}

    /**
     * Uploads a file. Not supported for batching
     * @param webUrl The absolute Url to the SharePoint site.
     * @param listName The name of the list the file will be uploaded to.
     * @param url The folder-relative url of the file.
     * @param content The Blob file content to add
     * @param progress A callback function which can be used to track the progress of the upload
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
     * @param chunkSize The size of each file slice, in bytes (default: 10485760)
     * @returns The new File and the raw response.
     */
	public async addFileBlob(
		webUrl: string,
		listName: string,
		url: string,
		content: Blob,
		progress?: (data: ChunkedFileUploadProgressData) => void,
		shouldOverwrite?: boolean,
		chunkSize?: number
	): Promise<FileAddResult> {
		return new Promise<FileAddResult>(async (resolve, reject) => {
			let web = new Web(webUrl);
			// large upload
			await web.lists
				.getByTitle(listName)
				.rootFolder.files.addChunked(url, content, progress, shouldOverwrite, chunkSize)
				.then((res) => {
					resolve(res);
				})
				.catch((error) => {
					reject(this.getErrorMessage(webUrl, error));
				});
		});
	}

    /**
     * Saves the given data to the specified SharePoint list item.
     *
     * @param webUrl The absolute Url to the SharePoint site.
     * @param listUrl The server-relative Url to the SharePoint list.
     * @param stringId The UniqueId of the list item to be updated.
     * @param fieldsSchema The array of field schema for all relevant fields of this list.
     * @param data An object containing all the field values to update.
     * @param newDocumentUpdate true if the list item is a document being updated after upload; otherwise false.
     * @returns Promise object represents the updated or erroneous form field values.
     */
	public async updateItem(
		webUrl: string,
		listName: string,
		stringId: string,
		fieldsSchema: IFieldSchema[],
		data: any,
		newDocumentUpdate?: boolean
	): Promise<ListItemFormUpdateValue[]> {
		return new Promise<ListItemFormUpdateValue[]>(async (resolve, reject) => {
			const xml = String.Format(
				// tslint:disable-next-line:max-line-length
				'<View Scope=\'RecursiveAll\'><Query><Where><Eq><FieldRef Name=\'UniqueId\' /><Va' +
					'lue Type=\'Guid\'>{0}</Value></Eq></Where></Query></View>',
				stringId
			);
			const values = this.GetFormValues(fieldsSchema, data);
			const camlQuery: CamlQuery = {
				ViewXml: xml
			};
			let web = new Web(webUrl);
			const list = web.lists.getByTitle(listName);
			await list.getItemsByCAMLQuery(camlQuery).then((res: any[]) => {
				list.items
					.getById(res[0].ID)
					.validateUpdateListItem(values, newDocumentUpdate)
					.then((respData) => {
						resolve(respData);
					})
					.catch((error) => {
						reject(this.getErrorMessage(webUrl, error));
					});
			});
		});
	}

	private GetFormValues(
		fieldsSchema: IFieldSchema[],
		data: any
	): Array<{
		FieldName: string;
		FieldValue: any;
		HasException: boolean;
		ErrorMessage: string;
	}> {
		return fieldsSchema
			.filter((field) => !field.ReadOnlyField && field.InternalName in data && data[field.InternalName] !== null)
			.map((field) => {
				return {
					ErrorMessage: null,
					FieldName: field.InternalName,
					FieldValue: data[field.InternalName],
					HasException: false
				};
			});
	}

	private getErrorMessage(webUrl: string, error: any): string {
		let errorMessage: string = error.statusText
			? error.statusText
			: error.statusMessage ? error.statusMessage : error;
		const serverUrl = `{window.location.protocol}//{window.location.hostname}`;
		const webServerRelativeUrl = webUrl.replace(serverUrl, '');

		if (error.status === 403) {
			errorMessage = String.Format(
				// tslint:disable-next-line:max-line-length
				'You do not have access to the previously configured web url \'{0}\'. Either leav' +
					'e the WebPart properties as is or select another web url.',
				webServerRelativeUrl
			);
		} else if (error.status === 404) {
			errorMessage = String.Format(
				// tslint:disable-next-line:max-line-length
				'The previously configured web url \'{0}\' is not found anymore. Either leave the' +
					' WebPart properties as is or select another web url.',
				webServerRelativeUrl
			);
		}
		return errorMessage;
	}
}
