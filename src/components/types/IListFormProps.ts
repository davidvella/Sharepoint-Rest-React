import { ControlMode } from '../../common/datatypes/ControlMode';
import { IFieldConfiguration } from './IFieldConfiguration';
import { IFieldSchema } from '../../common/services/datatypes/RenderListData';

export interface IListFormProps {
	title: string;
	description?: string;
	webUrl: string;
	listName: string;
	formType: ControlMode;
	fields?: IFieldConfiguration[];
	showUnsupportedFields?: boolean;
	onUpdateFields?: (fieldsSchema: IFieldSchema[], data: any) => void;
}
