import * as React from 'react';
import { BasePicker, IBasePickerProps, IPickerItemProps } from 'office-ui-fabric-react/lib/Pickers';
import { IPickerTerm, IPickerTerms } from './ITermPicker';
import SPTermStorePickerService from '../../common/services/SPTermStorePickerService';
import '../TaxonomyPicker.module.scss';
import { ITaxonomyPickerProps } from './ITaxonomyPicker';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export class TermBasePicker extends BasePicker<IPickerTerm, IBasePickerProps<IPickerTerm>> {}

export interface ITermPickerState {
	terms: IPickerTerms;
}

export interface ITermPickerProps {
	termPickerHostProps: ITaxonomyPickerProps;
	disabled: boolean;
	value: IPickerTerms;
	allowMultipleSelections: boolean;
	disabledTermIds?: string[];
	disableChildrenOfDisabledParents?: boolean;

	onChanged: (items: IPickerTerm[]) => void;
}

export default class TermPicker extends React.Component<ITermPickerProps, ITermPickerState> {
	private allTerms: IPickerTerms = null;

  /**
   * Constructor method
   */
	constructor(props: any) {
		super(props);
		this.onRenderItem = this.onRenderItem.bind(this);
		this.onRenderSuggestionsItem = this.onRenderSuggestionsItem.bind(this);
		this.onFilterChanged = this.onFilterChanged.bind(this);
		this.onGetTextFromItem = this.onGetTextFromItem.bind(this);

		this.state = {
			terms: this.props.value
		};
	}

  /**
   * componentWillReceiveProps method
   */
	public componentWillReceiveProps(nextProps: ITermPickerProps) {
		// check to see if props is different to avoid re-rendering
		let newKeys = nextProps.value.map((a) => a.Id);
		let currentKeys = this.state.terms.map((a) => a.Id);
		if (newKeys.sort().join(',') !== currentKeys.sort().join(',')) {
			this.setState({ terms: nextProps.value });
		}
	}

  /**
   * Render method
   */
	public render(): JSX.Element {
		return (
			<div>
				<TermBasePicker
					disabled={this.props.disabled}
					onResolveSuggestions={this.onFilterChanged}
					onRenderSuggestionsItem={this.onRenderSuggestionsItem}
					getTextFromItem={this.onGetTextFromItem}
					onRenderItem={this.onRenderItem}
					defaultSelectedItems={this.props.value}
					selectedItems={this.state.terms}
					onChange={this.props.onChanged}
					itemLimit={!this.props.allowMultipleSelections ? 1 : undefined}
					className={'termBasePicker'}
				/>
			</div>
		);
	}

  /**
   * Renders the item in the picker
   */
	protected onRenderItem(term: IPickerItemProps<IPickerTerm>) {
		return (
			<div
				className={'pickedTermRoot'}
				key={term.index}
				data-selection-index={term.index}
				data-is-focusable={!term.disabled && true}
			>
				<span className={'pickedTermText'}>{term.item.Name}</span>
				{!term.disabled && (
					<span className={'pickedTermCloseIcon'} onClick={term.onRemoveItem}>
						<Icon iconName="Cancel" />
					</span>
				)}
			</div>
		);
	}

  /**
   * Renders the suggestions in the picker
   */
	protected onRenderSuggestionsItem(term: IPickerTerm) {
		let termParent: string;
		let termTitle: string;
		term.sourceTerm.get().then((res) => {
			termParent = res.Name;
			termTitle = `${term.Name} [${res.Name}]`;
			if (term.PathOfTerm.indexOf(';') !== -1) {
				let splitPath = term.PathOfTerm.split(';');
				termParent = splitPath[splitPath.length - 2];
				splitPath.pop();
				termTitle = `${term.Name} [${res.Name}:${splitPath.join(':')}]`;
			}
		});
		return (
			<div className={'termSuggestion'} title={termTitle}>
				<div>{term.Name}</div>
				<div className={'termSuggestionSubTitle'}>
					{' '}
					{'in'}
					{termParent ? termParent : 'Term Set'}
				</div>
			</div>
		);
	}

  /**
   * When Filter Changes a new search for suggestions
   */
	private async onFilterChanged(filterText: string, tagList: IPickerTerm[]): Promise<IPickerTerm[]> {
		if (filterText !== '') {
			let termsService = new SPTermStorePickerService();
			let terms: IPickerTerm[] = await termsService.searchTermsByTermSet(
				this.props.termPickerHostProps.webUrl,
				this.props.termPickerHostProps.termsetNameOrID,
				filterText
			);
			// Filter out the terms which are already set
			const filteredTerms = [];
			const { disabledTermIds, disableChildrenOfDisabledParents } = this.props;
			for (const term of terms) {
				let canBePicked = true;

				// Check if term is not disabled
				if (disabledTermIds && disabledTermIds.length > 0) {
					// Check if current term need to be disabled
					if (disabledTermIds.indexOf(term.Id) !== -1) {
						canBePicked = false;
					} else {
						// Check if child terms need to be disabled
						if (disableChildrenOfDisabledParents) {
							// Check if terms were already retrieved
							if (!this.allTerms) {
								this.allTerms = await termsService.getAllTerms(
									this.props.termPickerHostProps.webUrl,
									this.props.termPickerHostProps.termsetNameOrID
								);
							}

							// Check if there are terms retrieved
							if (this.allTerms && this.allTerms.length > 0) {
								// Find the disabled parents
								const disabledParents = this.allTerms.filter(
									(t) => disabledTermIds.indexOf(t.Id) !== -1
								);
								// Check if disabled parents were found
								if (disabledParents && disabledParents.length > 0) {
									// Check if the current term lives underneath a disabled parent
									const findTerm = disabledParents.filter(
										(pt) => term.PathOfTerm.indexOf(pt.PathOfTerm) !== -1
									);
									if (findTerm && findTerm.length > 0) {
										canBePicked = false;
									}
								}
							}
						}
					}
				}

				if (canBePicked) {
					// Only retrieve the terms which are not yet tagged
					if (tagList.filter((tag) => tag.Id === term.Id).length === 0) {
						filteredTerms.push(term);
					}
				}
			}
			return filteredTerms;
		} else {
			return Promise.resolve([]);
		}
	}

  /**
   * gets the text from an item
   */
	private onGetTextFromItem(item: any): any {
		return item.name;
	}
}
