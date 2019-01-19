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
import { String } from 'typescript-string-operations';

/**
 * Service implementation to manage term stores in SharePoint
 */
export default class SPTermStorePickerService {
  /**
   * Gets the collection of term stores in the current SharePoint env
   */
	public async getTermStores(webUrl: string): Promise<(ITermStoreData & ITermStore)[]> {
		return new Promise<(ITermStoreData & ITermStore)[]>(async (resolve, reject) => {
			const taxonomy = new Session(webUrl);
			taxonomy.termStores
				.get()
				.then((res) => {
					resolve(res);
				})
				.catch((error) => {
					reject(this.getErrorMessage(webUrl, error));
				});
		});
	}

  /**
   * Gets the current term set
   */
	public async getTermSet(webUrl: string, termSetNameOrID: string): Promise<ITermSetData & ITermSet> {
		return new Promise<ITermSetData & ITermSet>(async (resolve, reject) => {
			let termStores = await this.getTermStores(webUrl);
			let termStore = termStores[0];
			termStore
				.getTermSetById(termSetNameOrID)
				.get()
				.then((res) => {
					resolve(res);
				})
				.catch((error) => {
					reject(this.getErrorMessage(webUrl, error));
				});
		});
	}

  /**
   * Retrieve all terms for the given term set
   * @param termset
   */
	public async getAllTerms(webUrl: string, termSetNameOrID: string): Promise<(ITermData & ITerm)[]> {
		return new Promise<(ITermData & ITerm)[]>(async (resolve, reject) => {
			const termSet = await this.getTermSet(webUrl, termSetNameOrID);
			termSet.terms
				.get()
				.then((res) => {
					resolve(res);
				})
				.catch((error) => {
					reject(this.getErrorMessage(webUrl, error));
				});
		});
	}

  /**
   * Retrieve all terms that starts with the searchText
   * @param searchText
   */
	public async searchTermsByName(webUrl: string, searchText: string): Promise<(ITermData & ITerm)[]> {
		return new Promise<(ITermData & ITerm)[]>(async (resolve, reject) => {
			const termStores = await this.getTermStores(webUrl);
			const termStore = termStores[0];
			const labelMatchInfo: ILabelMatchInfo = {
				StringMatchOption: StringMatchOption.StartsWith,
				TermLabel: searchText
			};
			termStore
				.getTerms(labelMatchInfo)
				.get()
				.then((res) => {
					resolve(res);
				})
				.catch((error) => {
					reject(this.getErrorMessage(webUrl, error));
				});
		});
	}

  /**
   * Searches terms for the given term set
   * @param searchText
   * @param termsetId
   */
	public async searchTermsByTermSet(
		webUrl: string,
		termSetNameOrID: string,
		searchText: string
	): Promise<(ITermData & ITerm)[]> {
		return new Promise<(ITermData & ITerm)[]>(async (resolve, reject) => {
			const terms = await this.getAllTerms(webUrl, termSetNameOrID);
			let returnTerms: (ITermData & ITerm)[] = [];
			terms.forEach((term) => {
				if (term.Name.toLowerCase().indexOf(searchText.toLowerCase()) !== -1) {
					returnTerms.push(term);
				}
			});
			resolve(returnTerms);
		});
	}

	private getErrorMessage(webUrl: string, error: any): string {
		let errorMessage: string = error.statusText
			? error.statusText
			: error.statusMessage ? error.statusMessage : error;
		const serverUrl = `{window.location.protocol}//{window.location.hostname}`;
		const webServerRelativeUrl = webUrl.replace(serverUrl, '');

		if (error.status === 403) {
			errorMessage = String.Format(
				// tslint:disable-next-line:max-line-length
				'You do not have access to the previously configured web url \'{0}\'. Either leav' +
					'e the WebPart properties as is or select another web url.',
				webServerRelativeUrl
			);
		} else if (error.status === 404) {
			errorMessage = String.Format(
				// tslint:disable-next-line:max-line-length
				'The previously configured web url \'{0}\' is not found anymore. Either leave the' +
					' WebPart properties as is or select another web url.',
				webServerRelativeUrl
			);
		}
		return errorMessage;
	}
}
