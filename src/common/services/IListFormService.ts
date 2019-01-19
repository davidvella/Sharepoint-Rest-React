import { ControlMode } from '../datatypes/ControlMode';
import { IFieldSchema, IRowData, IField } from './datatypes/RenderListData';

export interface IListFormService {
	getFieldSchemasForForm: (webUrl: string, listName: string, formType: ControlMode) => Promise<IFieldSchema[]>;
	getViewFields: (webUrl: string, listName: string, viewId: string) => Promise<IField[]>;
	getViewItems: (
		webUrl: string,
		listName: string,
		viewId: string,
		overrideParameters?: any,
		paging?: string
	) => Promise<IRowData>;
}
