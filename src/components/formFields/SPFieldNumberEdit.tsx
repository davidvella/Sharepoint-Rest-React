import * as React from 'react';
import { ISPFormFieldProps } from '../SPFormField';
import NumberFormField from './NumberFormField';

const SPFieldNumberEdit: React.FunctionComponent<ISPFormFieldProps> = (props) => {
	return (
		<NumberFormField
			className="ard-numberFormField"
			value={props.value}
			valueChanged={props.valueChanged}
			placeholder={'Enter value here'}
			underlined={true}
		/>
	);
};

export default SPFieldNumberEdit;
