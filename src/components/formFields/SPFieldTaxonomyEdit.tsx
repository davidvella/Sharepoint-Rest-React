import * as React from 'react';
import { ISPFormFieldProps } from '../SPFormField';
import '../SPFormField.scss';
import { TaxonomyPicker } from '../taxonomyPicker/TaxonomyPicker';
import { IPickerTerms } from '../taxonomyPicker/ITermPicker';
import { StringBuilder } from 'typescript-string-operations';

const SPFieldTaxonomyEdit: React.FunctionComponent<ISPFormFieldProps> = (props) => {
	function extractGuid(value: string): string {
		var regex = /([a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12})/i;
		// the RegEx will match the first occurrence of the pattern
		var match = regex.exec(value);

		// result is an array containing: [0] the entire string that was matched by our
		// RegEx [1] the first (only) group within our match, specified by the () within
		// our pattern, which contains the GUID value

		return match ? match[1] : null;
	}

	return (
		<TaxonomyPicker
			termsetNameOrID={props.fieldSchema.TermSetId}
			panelTitle="Select Term"
			label="Select a Term"
			onChange={(terms: IPickerTerms) => {
				var builder = new StringBuilder();
				for (let term of terms) {
					builder.Append(term.Name + '|' + extractGuid(term.Id) + ';');
				}
				props.valueChanged(builder.ToString());
			}}
			allowMultipleSelections={props.fieldSchema.AllowMultipleValues}
			webUrl={props.webUrl}
		/>
	);
};

export default SPFieldTaxonomyEdit;
