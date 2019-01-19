import { FileAddResult, ChunkedFileUploadProgressData } from '@pnp/sp';
import { IFieldSchema } from './datatypes/RenderListData';
import { ListItemFormUpdateValue } from '@pnp/sp';

export interface IDocumentService {
	getFileFromListAsBlob: (webUrl: string, fileRelativeUrl: string) => Promise<Blob>;
	addFileBlob: (
		webUrl: string,
		listName: string,
		url: string,
		content: Blob,
		progress?: (data: ChunkedFileUploadProgressData) => void,
		shouldOverWrite?: boolean,
		chunkSize?: number
	) => Promise<FileAddResult>;
	updateItem: (
		webUrl: string,
		listName: string,
		stringId: string,
		fieldsSchema: IFieldSchema[],
		data: any,
		newDocumentUpdate?: boolean
	) => Promise<ListItemFormUpdateValue[]>;
}
