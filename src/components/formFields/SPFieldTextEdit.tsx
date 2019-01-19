import * as React from 'react';
import { ISPFormFieldProps } from '../SPFormField';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { String } from 'typescript-string-operations';

const SPFieldTextEdit: React.FunctionComponent<ISPFormFieldProps> = (props) => {
	// We need to set value to empty string when null or undefined to force
	// TextField still be used like a controlled component
	const value = props.value ? props.value : '';
	const illegalFileNames = [ '/', '\\', '<', '>', ':', '*', '"', '?', '|', '.' ];
	return (
		<TextField
			className="ard-TextFormField"
			name={props.fieldSchema.InternalName}
			value={value}
			onChanged={(e) => {
				props.valueChanged(e);
			}}
			placeholder={'Enter value here'}
			multiline={props.fieldSchema.FieldType === 'Note'}
			errorMessage={
				props.fieldSchema.Required && String.IsNullOrWhiteSpace(value) ? (
					'You can\'t leave this blank.'
				) : props.fieldSchema.InternalName === 'FileLeafRef' &&
				illegalFileNames.some((substring) => value.indexOf(substring) !== -1) ? (
					'File names can\'t begin or end with a period, or contain any of these characters: / \\  < > : * " ? |.'
				) : null
			}
		/>
	);
};

export default SPFieldTextEdit;
