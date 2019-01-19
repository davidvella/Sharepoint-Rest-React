import { SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';

export interface IListViewProps {
  /**
   * Boolean value to indicate if the component should render in compact mode.
   * Set to false by default
   */
	webUrl: string;
  /**
   * Boolean value to indicate if the component should render in compact mode.
   * Set to false by default
   */
	listName: string;
  /**
   * Boolean value to indicate if the component should render in compact mode.
   * Set to false by default
   */
	viewId: string;
  /**
   * Boolean value to indicate if the component should render in compact mode.
   * Set to false by default
   */
	compact?: boolean;
  /**
   * Specify the item selection mode.
   * By default this is set to none.
   */
	selectionMode?: SelectionMode;
  /**
   * Boolean value to indicate if the component should render in compact mode.
   * Set to false by default
   */
	searchTerms?: string;
  /**
   * Selection event that passes the selected item(s)
   */
	selection?: (items: any[]) => void;
}
