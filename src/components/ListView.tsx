import * as React from 'react';

import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { IListViewProps } from './types/IListViewProps';
import { IListViewState } from './types/IListViewState';

import {
	DetailsList,
	DetailsListLayoutMode,
	IColumn,
	DetailsRow,
	IDetailsRowProps,
	ConstrainMode,
	SelectionMode,
	Selection
} from 'office-ui-fabric-react/lib/DetailsList';
import { IRowData } from '../common/services/datatypes/RenderListData';
import { IListFormService } from '../common/services/IListFormService';
import { ListFormService } from '../common/services/ListFormService';

import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { SPViewFields } from './SPViewFields';

/*************************************************************************************
 * React Component to render a SharePoint list on any page.
 *************************************************************************************/
export class ListView extends React.Component<IListViewProps, IListViewState> {
	readonly _selection: Selection;
	private listFormService: IListFormService;
	private _isMounted: boolean;
	constructor(props: IListViewProps) {
		super(props);

		// Binding the functions
		this._selection = new Selection({
			onSelectionChanged: () => {
				this.props.selection(this._selection.getSelection());
			}
		});

		// set initial state
		this.state = {
			isLoadingFields: false,
			isLoadingData: false,
			errors: [],
			items: []
		};

		this.listFormService = new ListFormService();
	}

  /**
   * Default React component render method
   */
	public render() {
		return (
			<div>
				{this.renderErrors()}
				{!this.props.listName ? (
					<MessageBar messageBarType={MessageBarType.warning}>
						Please configure a list for this component first.
					</MessageBar>
				) : (
					''
				)}
				{this.state.isLoadingFields ? (
					<Spinner size={SpinnerSize.large} label={'Loading the view...'} />
				) : (
					<div>{this.renderList()}</div>
				)}
			</div>
		);
	}

	/**
	 * Lifecycle hook when component is mounted
	 */
	public componentDidMount(): void {
		this._isMounted = true;
		if (this.props.listName != null && this.props.viewId != null) {
			this.readSchema(this.props.listName, this.props.viewId);
		}
	}

	public componentWillUnmount(): void {
		this._isMounted = false;
	}

  /**
   * Lifecycle hook when component did update after state or property changes
   * @param nextProps
   */
	public componentWillReceiveProps(nextProps: IListViewProps): void {
		if (this.props.listName !== nextProps.listName || this.props.viewId !== nextProps.viewId) {
			this.readSchema(nextProps.listName, nextProps.viewId);
		}
		if (this.props.searchTerms !== nextProps.searchTerms) {
			const overrideParameters: any = {
				InplaceSearchQuery: nextProps.searchTerms
			};
			this.readListData(overrideParameters);
		}
	}

	private renderList() {
		const { viewSchema, columns, items } = this.state;
		// const _columns: IColumn[] = [];
		return (
			<div>
				{viewSchema && viewSchema.length > 0 ? (
					<div>
						<DetailsList
							items={items}
							compact={this.props.compact}
							columns={columns}
							setKey="ListViewControl"
							layoutMode={DetailsListLayoutMode.justified}
							isHeaderVisible={true}
							selectionPreservedOnEmptyClick={true}
							enterModalSelectionOnTouch={true}
							onRenderRow={this._onRenderRow}
							selectionMode={this.props.selectionMode || SelectionMode.none}
							selection={this._selection}
							onRenderMissingItem={this._onLoadNextPage}
						/>{' '}
						{this.state.isLoadingData && (
							<Spinner className="loadingSpinner" label="Loading..." hidden={!this.state.isLoadingData} />
						)}
					</div>
				) : (
					<MessageBar messageBarType={MessageBarType.warning}>No fields available!</MessageBar>
				)}
			</div>
		);
	}

	private renderErrors() {
		return this.state.errors.length > 0 ? (
			<div>
				{this.state.errors.map((item, idx) => (
					<MessageBar
						messageBarType={MessageBarType.error}
						isMultiline={true}
						onDismiss={(_ev) => this.clearError(idx)}
					>
						{item}
					</MessageBar>
				))}
			</div>
		) : null;
	}

	@autobind
	private async readSchema(listName: string, viewId: string): Promise<void> {
		try {
			if (!listName) {
				if (this._isMounted) {
					this.setState({
						...this.state,
						isLoadingFields: false,
						errors: [ 'Please configure pass a list as a parameter.' ]
					});
				}
				return;
			}
			this.setState({
				...this.state,
				isLoadingFields: true
			});
			const viewSchema = await this.listFormService.getViewFields(this.props.webUrl, listName, viewId);
			const columns: IColumn[] = [];
			viewSchema.forEach((element) => {
				columns.push(SPViewFields.GenerateColumn(element, this._onColumnClick));
			});
			if (this._isMounted) {
				this.setState({
					...this.state,
					isLoadingFields: false,
					viewSchema,
					columns
				});
				this.readListData();
			}
		} catch (error) {
			const errorText = `Error loading for list with name: ${listName}: ${error}`;
			if (this._isMounted) {
				this.setState({
					...this.state,
					isLoadingFields: false,
					errors: [ ...this.state.errors, errorText ]
				});
			}
		}
	}

	@autobind
	private async readListData(overrideParameters?: any): Promise<void> {
		try {
			if (this._isMounted) {
				this.setState({
					...this.state,
					isLoadingData: true
				});
			}
			const items: IRowData = await this.listFormService.getViewItems(
				this.props.webUrl,
				this.props.listName,
				this.props.viewId,
				overrideParameters
			);
			const rows: any[] = items.Row;
			if (items.NextHref) {
				rows.push(null);
			}
			if (this._isMounted) {
				this.setState({
					...this.state,
					isLoadingData: false,
					items: rows,
					NextHref: items.NextHref
				});
			}
		} catch (error) {
			const errorText = `Error loading for data list with name: ${this.props.listName}: ${error}`;
			if (this._isMounted) {
				this.setState({
					...this.state,
					isLoadingData: false,
					errors: [ ...this.state.errors, errorText ]
				});
			}
		}
	}

	private _onLoadNextPage = (index: number): null => {
		let { items, NextHref } = this.state;
		this.setState({
			...this.state,
			isLoadingData: true
		});
		this.listFormService
			.getViewItems(this.props.webUrl, this.props.listName, this.props.viewId, null, NextHref.substring(1))
			.then((res) => {
				const tempRows = res.Row;
				if (res.NextHref) {
					tempRows.push(null);
				}
				items = items.slice(0, items.length - 1).concat(tempRows);
				if (this._isMounted) {
					this.setState({
						...this.state,
						isLoadingData: false,
						items,
						NextHref: res.NextHref
					});
				}
			})
			.catch((error) => {
				const errorText = `Error loading for data list with name: ${this.props.listName}: ${error}`;
				this.setState({
					...this.state,
					isLoadingData: false,
					errors: [ ...this.state.errors, errorText ]
				});
			});
		return null;
	}

  /**
   * Check if sorting needs to be set to the column
   * @param ev
   * @param column
   */
	private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
		const { columns } = this.state;
		const newColumns: IColumn[] = columns.slice();
		const currColumn: IColumn = newColumns.filter((currCol: IColumn, idx: number) => {
			return column.key === currCol.key;
		})[0];
		newColumns.forEach((newCol: IColumn) => {
			if (newCol === currColumn) {
				currColumn.isSortedDescending = !currColumn.isSortedDescending;
				currColumn.isSorted = true;
			} else {
				newCol.isSorted = false;
				newCol.isSortedDescending = true;
			}
		});
		const overrideParameters: any = {
			SortField: currColumn.fieldName,
			SortDir: currColumn.isSortedDescending ? 'Desc' : 'Asc'
		};
		this.readListData(overrideParameters);
		this.setState({ columns: newColumns });
	}

	private clearError(idx: number) {
		this.setState((prevState, _props) => {
			return {
				...prevState,
				errors: prevState.errors.splice(idx, 1)
			};
		});
	}

	private _onRenderRow = (props: IDetailsRowProps): JSX.Element => {
		return <DetailsRow {...props} aria-busy={false} />;
	}
}
