import { IPickerTerms, IPickerTerm } from './ITermPicker';
import {
	Session,
	ITermStoreData,
	ITermStore,
	ITermSet,
	ILabelMatchInfo,
	StringMatchOption,
	ITermData,
	ITerm,
	ITermSetData
} from '@pnp/sp-taxonomy';

/**
 * PropertyFieldTermPickerHost properties interface
 */
export interface ITaxonomyPickerProps {
  /**
   * Url to retrieve taxonomy settings
   */
	webUrl: string;
  /**
   * Property field label displayed on top
   */
	label: string;
  /**
   * TermSet Picker Panel title
   */
	panelTitle: string;
  /**
   * Defines if the user can select only one or many term sets. Default value is false.
   */
	allowMultipleSelections?: boolean;
  /**
   * Defines the selected by default term sets.
   */
	initialValues?: IPickerTerms;
  /**
   * Limit the terms that can be picked by the Term Set name or ID
   */
	termsetNameOrID: string;
  /**
   * Id of a child term in the termset where to be able to selected and search the terms from
   */
	anchorId?: string;
  /**
   * Specify if the term set itself is selectable in the tree view
   */
	isTermSetSelectable?: boolean;
  /**
   * Specify which terms should be disabled in the term set so that they cannot be selected
   */
	disabledTermIds?: string[];
  /**
   * Specify if you want to disable the child terms when their parent is disabled
   */
	disableChildrenOfDisabledParents?: boolean;
  /**
   * Whether the property pane field is enabled or not.
   */
	disabled?: boolean;
  /**
   * The method is used to get the validation error message and determine whether the input value is valid or not.
   *
   *   When it returns string:
   *   - If valid, it returns empty string.
   *   - If invalid, it returns the error message string and the text field will
   *     show a red border and show an error message below the text field.
   *
   *   When it returns Promise<string>:
   *   - The resolved value is display as error message.
   *   - The rejected, the value is thrown away.
   *
   */
	onGetErrorMessage?: (value: IPickerTerms) => string | Promise<string>;

  /**
   * onChange Event
   */
	onChange?: (newValue?: IPickerTerms) => void;
}

/**
 * PropertyFieldTermPickerHost state interface
 */
export interface ITaxonomyPickerState {
	termSetAndTerms?: ITermSet & ITermSetData;
	errorMessage?: string;
	openPanel?: boolean;
	loaded?: boolean;
	activeNodes?: IPickerTerms;
}

export interface ITermChanges {
	changedCallback: (term: ITerm, checked: boolean) => void;
	activeNodes?: IPickerTerms;
	disabledTermIds?: string[];
	disableChildrenOfDisabledParents?: boolean;
}

export interface ITermParentProps extends ITermChanges {
	termset: ITermSet & ITermSetData;
	multiSelection: boolean;
	anchorId?: string;
	autoExpand: () => void;
}

export interface ITermParentState {
	loaded?: boolean;
	expanded?: boolean;
}

export interface ITermProps extends ITermChanges {
	termset: string;
	term: IPickerTerm;
	multiSelection: boolean;
	disabled: boolean;
}

export interface ITermState {
	selected?: boolean;
}
