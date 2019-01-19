import { IField } from '../../common/services/datatypes/RenderListData';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';

export interface IListViewState {
	viewSchema?: IField[];
	columns?: IColumn[];
	isLoadingFields: boolean;
	isLoadingData: boolean;
	errors: string[];
	items: any[];
	NextHref?: string;
}
