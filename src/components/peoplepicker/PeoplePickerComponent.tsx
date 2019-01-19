import * as React from 'react';
import { IPeoplePickerProps, IPeoplePickerState, IPeoplePickerUserItem } from './IPeoplePicker';
import { TooltipHost, DirectionalHint } from 'office-ui-fabric-react/lib/Tooltip';
import { NormalPeoplePicker } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePicker';
import { MessageBar } from 'office-ui-fabric-react/lib/MessageBar';
import '../PeoplePickerComponent.module.scss';
import { assign } from 'office-ui-fabric-react/lib/Utilities';
import { IUsers } from './IUsers';
import { Label } from 'office-ui-fabric-react/lib/components/Label';
import { IBasePickerSuggestionsProps } from 'office-ui-fabric-react/lib/components/pickers/BasePicker.types';
import { IPersonaProps } from 'office-ui-fabric-react/lib/components/Persona/Persona.types';
import { MessageBarType } from 'office-ui-fabric-react/lib/components/MessageBar';
import { ValidationState } from 'office-ui-fabric-react/lib/components/pickers/BasePicker.types';
import { Icon } from 'office-ui-fabric-react/lib/components/Icon';
import { isEqual } from '@microsoft/sp-lodash-subset';
import { Web, SiteUserProps } from '@pnp/sp';
 
/**
 * PeoplePicker component
 */
export class PeoplePicker extends React.Component<IPeoplePickerProps, IPeoplePickerState> {
	constructor(props: IPeoplePickerProps) {
		super(props);

		this.state = {
			selectedPersons: [],
			mostRecentlyUsedPersons: [],
			currentSelectedPersons: [],
			allPersons: [],
			currentPicker: 0,
			peoplePartTitle: '',
			peoplePartTooltip: '',
			isLoading: false,
			showmessageerror: false
		};
	}

    /**
     * componentWillMount lifecycle hook
     */
	public componentWillMount(): void {
		// online mode Load the users
		this._thisLoadUsers();
	}

  /**
   * componentDidUpdate lifecycle hook
   */
	public componentDidUpdate(prevProps: IPeoplePickerProps, prevState: IPeoplePickerState): void {
		// If defaultSelectedUsers has changed then bind again
		if (
			!isEqual(this.props.defaultSelectedUsers, prevProps.defaultSelectedUsers) ||
			!isEqual(this.state.allPersons, prevState.allPersons)
		) {
			// Check if we have results to get from, if not provide a empty array to filter
			// on
			let userValuesArray: Array<IPeoplePickerUserItem> =
				this.state.allPersons.length !== 0 ? this.state.allPersons : new Array<IPeoplePickerUserItem>();

			// Set Default selected persons
			let defaultUsers: any = [];
			let defaultPeopleList: IPersonaProps[] = [];
			if (this.props.defaultSelectedUsers) {
				defaultUsers = this.getDefaultUsers(userValuesArray, this.props.defaultSelectedUsers);
				for (const persona of defaultUsers) {
					let selectedPeople: IPersonaProps = {};
					assign(selectedPeople, persona);
					defaultPeopleList.push(selectedPeople);
				}
			}

			this.setState({
				selectedPersons: defaultPeopleList.length !== 0 ? defaultPeopleList : [],
				showmessageerror: this.props.isRequired && defaultPeopleList.length === 0
			});
		}
	}

  /**
   * Default React component render method
   */
	public render(): React.ReactElement<IPeoplePickerProps> {
		const suggestionProps: IBasePickerSuggestionsProps = {
			suggestionsHeaderText: 'Suggested People',
			noResultsFoundText: 'No results found',
			loadingText: 'Loading',
			resultsMaximumNumber: this.props.suggestionsLimit ? this.props.suggestionsLimit : 5
		};

		const peoplepicker = (
			<div
				id="people"
				className={`defaultClass ${this.props.peoplePickerWPclassName
					? this.props.peoplePickerWPclassName
					: ''}`}
			>
				<Label>{this.props.titleText || 'Pick the users(s)'}</Label>

				<NormalPeoplePicker
					pickerSuggestionsProps={suggestionProps}
					onResolveSuggestions={this._onPersonFilterChanged}
					onEmptyInputFocus={this._returnMostRecentlyUsedPerson}
					getTextFromItem={(peoplePersonaMenu: IPersonaProps) => peoplePersonaMenu.text}
					className={`'ms-PeoplePicker' ${this.props.peoplePickerCntrlclassName
						? this.props.peoplePickerCntrlclassName
						: ''}`}
					key={'normal'}
					onValidateInput={this._validateInputPeople}
					removeButtonAriaLabel={'Remove'}
					inputProps={{
						'aria-label': 'People Picker'
					}}
					selectedItems={this.state.selectedPersons}
					itemLimit={this.props.personSelectionLimit || 1}
					disabled={this.props.disabled}
					onChange={this._onPersonItemsChange}
				/>
			</div>
		);

		return (
			<div>
				{this.props.showtooltip ? (
					<TooltipHost
						content={this.props.tooltipMessage || 'People Picker'}
						id="pntp"
						calloutProps={{
							gapSpace: 0
						}}
						directionalHint={this.props.tooltipDirectional || DirectionalHint.leftTopEdge}
					>
						{peoplepicker}
					</TooltipHost>
				) : (
					<div>{peoplepicker}</div>
				)}

				{this.props.isRequired &&
				this.state.showmessageerror && (
					<p
						className={`ms-TextField-errorMessage errorMessage ${this.props.errorMessageClassName
							? this.props.errorMessageClassName
							: ''}`}
					>
						<Icon iconName="Error" className={'errorIcon'} />
						<span data-automation-id="error-message">
							{this.props.errorMessage ? this.props.errorMessage : 'Required Field'}
						</span>
					</p>
				)}
			</div>
		);
	}
  /**
   * Retrieve the users
   */
	private async _thisLoadUsers(): Promise<void> {
		let web = new Web(this.props.webAbsoluteUrl);

		let users: SiteUserProps[];

		// filter for principal Type
		var filterVal: string = '';
		if (this.props.principleTypes) {
			filterVal = `${this.props.principleTypes
				.map((principalType) => `(PrincipalType eq ${principalType})`)
				.join(' or ')}`;
		}

		// filter for showHiddenInUI
		if (this.props.showHiddenInUI) {
			filterVal = filterVal
				? `${filterVal} and (IsHiddenInUI eq ${this.props.showHiddenInUI})`
				: `?$filter=IsHiddenInUI eq ${this.props.showHiddenInUI}`;
		}

		try {
			if (this.props.groupName) {
				users = await web.siteUsers.filter(filterVal).getByLoginName(this.props.groupName).get();
			} else {
				users = await web.siteUsers.filter(filterVal).get();
			}
			// Check if items were retrieved
			if (users && users.length > 0) {
				let userValuesArray: Array<IPeoplePickerUserItem> = new Array<IPeoplePickerUserItem>();

				// Loop over all the retrieved items
				for (let i = 0; i < users.length; i++) {
					const item = users[i];
					if (!item.IsHiddenInUI || (this.props.showHiddenInUI && item.IsHiddenInUI)) {
						// Check if the the type must be returned
						if (
							!this.props.principleTypes ||
							this.props.principleTypes.indexOf(item.PrincipalType) !== -1
						) {
							userValuesArray.push({
								id: item.Id.toString(),
								imageUrl: '',
								imageInitials: '',
								text: item.Title, // name
								secondaryText: item.Email, // email
								tertiaryText: '', // status
								optionalText: '', // anything
								loginName: item.LoginName
							});
						}
					}
				}

				let personaList: IPersonaProps[] = [];
				for (const persona of userValuesArray) {
					let personaWithMenu: IPersonaProps = {};
					assign(personaWithMenu, persona);
					personaList.push(personaWithMenu);
				}

				// Update the current state
				this.setState({
					allPersons: userValuesArray,
					peoplePersonaMenu: personaList,
					mostRecentlyUsedPersons: personaList.slice(0, 5)
				});
			}
		} catch (e) {
			console.error('Error occured while fetching the users and setting selected users.' + e);
		}
	}

  /**
   * On persona item changed event
   */
	private _onPersonItemsChange = (items: any[]) => {
		const { selectedItems } = this.props;

		this.setState({
			selectedPersons: items,
			showmessageerror: items.length > 0 ? false : true
		});

		if (selectedItems) {
			selectedItems(items);
		}
	}

  /**
   * Validates the user input
   *
   * @param input
   */
	private _validateInputPeople = (input: string) => {
		if (input.indexOf('@') !== -1) {
			return ValidationState.valid;
		} else if (input.length > 1) {
			return ValidationState.warning;
		} else {
			return ValidationState.invalid;
		}
	}

  /**
   * Returns the most recently used person
   *
   * @param currentPersonas
   */
	private _returnMostRecentlyUsedPerson = (currentPersonas: IPersonaProps[]): IPersonaProps[] => {
		let { mostRecentlyUsedPersons } = this.state;
		return this._removeDuplicates(mostRecentlyUsedPersons, currentPersonas);
	}

  /**
   * On filter changed event
   *
   * @param filterText
   * @param currentPersonas
   * @param limitResults
   */
	private _onPersonFilterChanged = (
		filterText: string,
		currentPersonas: IPersonaProps[],
		limitResults?: number
	): IPersonaProps[] => {
		if (filterText) {
			let filteredPersonas: IPersonaProps[] = this._filterPersons(filterText);
			filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
			filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
			return filteredPersonas;
		} else {
			return [];
		}
	}

  /**
   * Filter persons based on Name and Email (starting with and contains)
   *
   * @param filterText
   */
	private _filterPersons(filterText: string): IPersonaProps[] {
		return this.state.peoplePersonaMenu.filter(
			(item) =>
				this._doesTextStartWith(item.text as string, filterText) ||
				this._doesTextContains(item.text as string, filterText) ||
				this._doesTextStartWith(item.secondaryText as string, filterText) ||
				this._doesTextContains(item.secondaryText as string, filterText)
		);
	}

  /**
   * Removes duplicates
   *
   * @param personas
   * @param possibleDupes
   */
	private _removeDuplicates = (personas: IPersonaProps[], possibleDupes: IPersonaProps[]): IPersonaProps[] => {
		return personas.filter((persona) => !this._listContainsPersona(persona, possibleDupes));
	}

  /**
   * Checks if text starts with
   *
   * @param text
   * @param filterText
   */
	private _doesTextStartWith(text: string, filterText: string): boolean {
		return text && text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
	}

/**
 * Checks if text contains
 *
 * @param text
 * @param filterText
 */
	private _doesTextContains(text: string, filterText: string): boolean {
		return text && text.toLowerCase().indexOf(filterText.toLowerCase()) > 0;
	}

  /**
   * Checks if list contains the person
   *
   * @param persona
   * @param personas
   */
	private _listContainsPersona = (persona: IPersonaProps, personas: IPersonaProps[]): boolean => {
		if (!personas || !personas.length || personas.length === 0) {
			return false;
		}
		return personas.filter((item) => item.text === persona.text).length > 0;
	}

  /**
   * Gets the default users based on the provided email address.
   * Adds emails that are not found with a random generated User Id
   *
   * @param userValuesArray
   * @param selectedUsers
   */
	private getDefaultUsers(userValuesArray: any[], selectedUsers: string[]): any {
		let defaultuserValuesArray: any[] = [];
		for (let i = 0; i < selectedUsers.length; i++) {
			const obj = {
				valToCompare: selectedUsers[i]
			};
			const length = defaultuserValuesArray.length;
			defaultuserValuesArray =
				defaultuserValuesArray.length !== 0
					? defaultuserValuesArray.concat(userValuesArray.filter(this.filterUsers, obj))
					: userValuesArray.filter(this.filterUsers, obj);
			if (length === defaultuserValuesArray.length) {
				const defaultUnknownUser = [
					{
						id: 1000 + i, // just a random number
						imageUrl: '',
						imageInitials: '',
						text: selectedUsers[i], // Name
						secondaryText: selectedUsers[i], // Role
						tertiaryText: '', // status
						optionalText: '' // stgring
					}
				];
				defaultuserValuesArray =
					defaultuserValuesArray.length !== 0
						? defaultuserValuesArray.concat(defaultUnknownUser)
						: defaultUnknownUser;
			}
		}
		return defaultuserValuesArray;
	}

  /**
   * Filters Users based on email
   */
	private filterUsers = function(value: any, index: number, ar: any[]) {
		if (value.secondaryText.toLowerCase().indexOf(this.valToCompare.toLowerCase()) !== -1) {
			return value;
		}
	};
}
