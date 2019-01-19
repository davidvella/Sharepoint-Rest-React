import { IField } from '../common/services/datatypes/RenderListData';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { FileTypeIcon } from './fileTypeIcon/FileTypeIcon';
import { IconType } from './fileTypeIcon/IFileTypeIcon';
import React = require('react');

export abstract class SPViewFields {
	public static GenerateColumn(fieldSchema: IField, columnClickHandler: any): IColumn {
		let column: IColumn;

		switch (fieldSchema.FieldType) {
			case 'Computed':
				column = this.GenerateColumnComputed(fieldSchema, columnClickHandler);
				break;
			case 'DateTime':
				column = this.GenerateColumnDate(fieldSchema, columnClickHandler);
				break;
			case 'User':
				column = this.GenerateColumnUser(fieldSchema, columnClickHandler);
				break;
			case 'Note':
				column = this.GenerateColumnString(fieldSchema, columnClickHandler);
				break;
			case 'Text':
				column = this.GenerateColumnString(fieldSchema, columnClickHandler);
				break;
			case 'Choice':
				column = this.GenerateColumnString(fieldSchema, columnClickHandler);
				break;
			case 'Number':
				column = this.GenerateColumnNumber(fieldSchema, columnClickHandler);
				break;
			case 'Currency':
				column = this.GenerateColumnCurrency(fieldSchema, columnClickHandler);
				break;
			case 'Lookup':
				column = this.GenerateColumnLookup(fieldSchema, columnClickHandler);
				break;
			case 'Boolean':
				column = this.GenerateColumnString(fieldSchema, columnClickHandler);
				break;
			case 'URL':
				column = this.GenerateColumnUrl(fieldSchema, columnClickHandler);
				break;
			case 'OutcomeChoice':
				column = this.GenerateColumnString(fieldSchema, columnClickHandler);
				break;
			case 'TaxonomyFieldType':
				column = this.GenerateColumnString(fieldSchema, columnClickHandler);
				break;
			default:
				column = this.GenerateColumnString(fieldSchema, columnClickHandler);
				break;
		}

		return column;
	}

	public static GenerateColumnText(fieldSchema: IField, columnClickHandler: any): IColumn {
		return {
			key: fieldSchema.ID,
			name: fieldSchema.DisplayName,
			fieldName: fieldSchema.RealFieldName,
			minWidth: 70,
			maxWidth: 350,
			isRowHeader: true,
			isResizable: true,
			isSorted: false,
			onColumnClick: columnClickHandler,
			isSortedDescending: false,
			sortAscendingAriaLabel: 'Sorted A to Z',
			sortDescendingAriaLabel: 'Sorted Z to A',
			data: 'string',
			isPadded: true
		};
	}

	public static GenerateColumnString(fieldSchema: IField, columnClickHandler: any): IColumn {
		return {
			key: fieldSchema.ID,
			name: fieldSchema.DisplayName,
			fieldName: fieldSchema.RealFieldName,
			minWidth: 70,
			maxWidth: 120,
			isRowHeader: true,
			onColumnClick: columnClickHandler,
			isResizable: true,
			isSorted: false,
			isSortedDescending: false,
			data: 'string',
			isPadded: true,
			onRender: (item: any) => {
				const fieldName = fieldSchema.RealFieldName;
				return <span>{item[fieldName]}</span>;
			}
		};
	}

	public static GenerateColumnComputed(fieldSchema: IField, columnClickHandler: any): IColumn {
		if (fieldSchema.RealFieldName === 'DocIcon') {
			return {
				key: fieldSchema.ID,
				name: fieldSchema.DisplayName,
				fieldName: fieldSchema.RealFieldName,
				iconName: 'Page',
				isIconOnly: true,
				minWidth: 16,
				maxWidth: 16,
				onColumnClick: columnClickHandler,
				onRender: (item: any) => {
					const docUrl: string = item.FileRef;
					return <FileTypeIcon type={IconType.image} path={docUrl} />;
				}
			};
		} else {
			return {
				key: fieldSchema.ID,
				name: fieldSchema.DisplayName,
				fieldName: fieldSchema.RealFieldName,
				minWidth: 100,
				maxWidth: 150,
				isRowHeader: true,
				onColumnClick: columnClickHandler,
				isResizable: true,
				isSorted: false,
				isSortedDescending: false,
				data: 'string',
				isPadded: true,
				onRender: (item: any) => {
					const fieldName = fieldSchema.RealFieldName;
					return <span>{item[fieldName]}</span>;
				}
			};
		}
	}

	public static GenerateColumnLookup(fieldSchema: IField, columnClickHandler: any): IColumn {
		return {
			key: fieldSchema.ID,
			name: fieldSchema.DisplayName,
			fieldName: fieldSchema.RealFieldName,
			minWidth: 70,
			maxWidth: 120,
			isRowHeader: true,
			isResizable: true,
			onColumnClick: columnClickHandler,
			isSorted: false,
			isSortedDescending: false,
			data: 'string',
			isPadded: true,
			onRender: (item: any) => {
				return <span>{''}</span>;
			}
		};
	}

	public static GenerateColumnUrl(fieldSchema: IField, columnClickHandler: any): IColumn {
		return {
			key: fieldSchema.ID,
			name: fieldSchema.DisplayName,
			fieldName: fieldSchema.RealFieldName,
			minWidth: 70,
			maxWidth: 120,
			isRowHeader: true,
			onColumnClick: columnClickHandler,
			isResizable: true,
			isSorted: false,
			isSortedDescending: false,
			data: 'Url',
			isPadded: true,
			onRender: (item: any) => {
				const fieldName = fieldSchema.RealFieldName;
				return <Link href={`${item[fieldName]}`}>{item[fieldName]}</Link>;
			}
		};
	}

	public static GenerateColumnUser(fieldSchema: IField, columnClickHandler: any): IColumn {
		return {
			key: fieldSchema.ID,
			name: fieldSchema.DisplayName,
			fieldName: fieldSchema.RealFieldName,
			minWidth: 70,
			maxWidth: 90,
			isRowHeader: true,
			isResizable: true,
			onColumnClick: columnClickHandler,
			isSorted: false,
			isSortedDescending: false,
			data: 'string',
			isPadded: true,
			onRender: (item: any) => {
				const fieldName = fieldSchema.RealFieldName;
				const user = item[fieldName];
				return <span>{user[0].title}</span>;
			}
		};
	}

	public static GenerateColumnCurrency(fieldSchema: IField, columnClickHandler: any): IColumn {
		return {
			key: fieldSchema.ID,
			name: fieldSchema.DisplayName,
			fieldName: fieldSchema.RealFieldName,
			minWidth: 70,
			maxWidth: 90,
			isRowHeader: true,
			isResizable: true,
			onColumnClick: columnClickHandler,
			isSorted: false,
			isSortedDescending: false,
			data: 'string',
			isPadded: true,
			onRender: (item: any) => {
				const fieldName = fieldSchema.RealFieldName;
				return <span>{item[fieldName]}</span>;
			}
		};
	}

	public static GenerateColumnDate(fieldSchema: IField, columnClickHandler: any): IColumn {
		return {
			key: fieldSchema.ID,
			name: fieldSchema.DisplayName,
			fieldName: fieldSchema.RealFieldName,
			minWidth: 70,
			maxWidth: 90,
			isRowHeader: true,
			isResizable: true,
			onColumnClick: columnClickHandler,
			isSorted: false,
			isSortedDescending: false,
			data: 'number',
			isPadded: true,
			onRender: (item: any) => {
				const fieldName = fieldSchema.RealFieldName;
				return <span>{item[fieldName]}</span>;
			}
		};
	}

	public static GenerateColumnNumber(fieldSchema: IField, columnClickHandler: any): IColumn {
		return {
			key: fieldSchema.ID,
			name: fieldSchema.DisplayName,
			fieldName: fieldSchema.RealFieldName,
			minWidth: 70,
			maxWidth: 90,
			onColumnClick: columnClickHandler,
			isResizable: true,
			isCollapsable: true,
			data: 'number',
			onRender: (item: any) => {
				const fieldName = fieldSchema.RealFieldName;
				return <span>{item[fieldName]}</span>;
			}
		};
	}
}
