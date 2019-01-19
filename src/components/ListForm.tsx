import * as React from 'react';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { IFieldConfiguration } from './types/IFieldConfiguration';
import { IListFormProps } from './types/IListFormProps';
import { IListFormState } from './types/IListFormState';
import { ControlMode } from '../common/datatypes/ControlMode';

import { IListFormService } from '../common/services/IListFormService';
import { ListFormService } from '../common/services/ListFormService';

import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { css } from 'office-ui-fabric-react/lib/Utilities';

import SPFormField from './SPFormField';

import './ListForm.module.scss';

const styles = {
	listForm: 'listForm_7906cbc9',
	title: 'title_7906cbc9',
	description: 'description_7906cbc9',
	formFieldsContainer: 'formFieldsContainer_7906cbc9',
	isDataLoading: 'isDataLoading_7906cbc9',
	formButtonsContainer: 'formButtonsContainer_7906cbc9',
	addFieldToolbox: 'addFieldToolbox_7906cbc9',
	addFieldToolboxPlusButton: 'addFieldToolboxPlusButton_7906cbc9'
};

/*************************************************************************************
 * React Component to render a SharePoint list form on any page.
 * The list form can be configured to be either a new form for adding a new list item,
 * an edit form for changing an existing list item or a display form for showing the
 * fields of an existing list item.
 * In design mode the fields to render can be moved, added and deleted.
 *************************************************************************************/
export class ListForm extends React.Component<IListFormProps, IListFormState> {
	private listFormService: IListFormService;

	constructor(props: IListFormProps) {
		super(props);

		// set initial state
		this.state = {
			isLoadingSchema: false,
			isLoadingData: false,
			isSaving: false,
			data: {},
			originalData: {},
			errors: [],
			notifications: [],
			fieldErrors: {}
		};

		this.listFormService = new ListFormService();
	}

	public render() {
		return (
			<div className={styles.listForm}>
				<div className={css(styles.title, 'ms-font-xl')}>{this.props.title}</div>
				{this.props.description && <div className={styles.description}>{this.props.description}</div>}
				{this.renderNotifications()}
				{this.renderErrors()}
				{!this.props.listName ? (
					<MessageBar messageBarType={MessageBarType.warning}>
						Please configure a list for this component first.
					</MessageBar>
				) : (
					''
				)}
				{this.state.isLoadingSchema ? (
					<Spinner size={SpinnerSize.large} label={'Loading the form...'} />
				) : (
					this.state.fieldsSchema && (
						<div>
							<div
								className={css(
									styles.formFieldsContainer,
									this.state.isLoadingData ? styles.isDataLoading : null
								)}
							>
								{this.renderFields()}
							</div>
						</div>
					)
				)}
			</div>
		);
	}

	public componentDidMount(): void {
		this.readSchema(this.props.listName, this.props.formType);
	}

	public componentWillReceiveProps(nextProps: IListFormProps): void {
		if (this.props.listName !== nextProps.listName || this.props.formType !== nextProps.formType) {
			this.readSchema(nextProps.listName, nextProps.formType);
		}
	}

	private renderNotifications() {
		if (this.state.notifications.length === 0) {
			return null;
		}
		setTimeout(() => {
			this.setState({
				...this.state,
				notifications: []
			});
		},         4000);
		return (
			<div>
				{this.state.notifications.map((item, _idx) => (
					<MessageBar messageBarType={MessageBarType.success}>{item}</MessageBar>
				))}
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

	private renderFields() {
		const { fieldsSchema, data, fieldErrors } = this.state;
		const fields = this.getFields();
		return fields && fields.length > 0 ? (
			<div className="ard-formFieldsContainer">
				{fields.map((field, _idx) => {
					const fieldSchemas = fieldsSchema!.filter((f) => f.InternalName === field.fieldName);
					if (fieldSchemas.length > 0) {
						const fieldSchema = fieldSchemas[0];
						const value = data[field.fieldName];
						let extraData;
						if (data.hasOwnProperty(field.fieldName + '.')) {
							extraData = data[field.fieldName + '.'];
						} else {
							extraData = Object.keys(data)
								.filter((propName) => propName.indexOf(field.fieldName + '.') === 0)
								.reduce((newData, pn) => {
									newData[pn.substring(field.fieldName.length + 1)] = data[pn];
									return newData;
								},      {});
						}
						const errorMessage = fieldErrors[field.fieldName];
						const fieldComponent = SPFormField({
							fieldSchema: fieldSchema,
							controlMode: this.props.formType,
							value: value,
							extraData: extraData,
							errorMessage: errorMessage,
							hideIfFieldUnsupported: !this.props.showUnsupportedFields,
							valueChanged: (val) => this.valueChanged(field.fieldName, val),
							webUrl: this.props.webUrl
						});
						return fieldComponent;
					}
					return null;
				})}
			</div>
		) : (
			<MessageBar messageBarType={MessageBarType.warning}>No fields available!</MessageBar>
		);
	}

	@autobind
	private async readSchema(listName: string, formType: ControlMode): Promise<void> {
		try {
			if (!listName) {
				this.setState({
					...this.state,
					isLoadingSchema: false,
					errors: [ 'Please configure a list in the web part\'s editor first.' ]
				});
				return;
			}
			this.setState({
				...this.state,
				isLoadingSchema: true
			});
			const fieldsSchema = await this.listFormService.getFieldSchemasForForm(
				this.props.webUrl,
				listName,
				formType
			);
			this.setState({
				...this.state,
				isLoadingSchema: false,
				fieldsSchema
			});
		} catch (error) {
			const errorText = `Error loading for list with name: ${listName}: ${error}`;
			this.setState({
				...this.state,
				isLoadingSchema: false,
				errors: [ ...this.state.errors, errorText ]
			});
		}
	}

	@autobind
	private valueChanged(fieldName: string, newValue: any) {
		this.setState(
			(prevState, _props) => {
				return {
					...prevState,
					data: {
						...prevState.data,
						[fieldName]: newValue
					},
					fieldErrors: {
						...prevState.fieldErrors,
						[fieldName]:
							prevState.fieldsSchema!.filter((item) => item.InternalName === fieldName)[0].Required &&
							!newValue
								? 'Please enter a value!'
								: ''
					}
				};
			},
			() => {
				this.props.onUpdateFields(this.state.fieldsSchema, this.state.data);
			}
		);
	}

	private clearError(idx: number) {
		this.setState((prevState, _props) => {
			return {
				...prevState,
				errors: prevState.errors.splice(idx, 1)
			};
		});
	}

	private getFields(): IFieldConfiguration[] | undefined {
		let fields: IFieldConfiguration[] | undefined = this.props.fields;
		if (!fields && this.state.fieldsSchema) {
			fields = this.state.fieldsSchema.map((field) => ({
				key: field.InternalName,
				fieldName: field.InternalName
			}));
		}
		return fields;
	}
}
