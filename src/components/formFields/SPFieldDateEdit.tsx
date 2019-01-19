import * as React from 'react';
import * as moment from 'moment';
import { css } from 'office-ui-fabric-react/lib/Utilities';
import { Locales } from '../../common/Locales';
import { ISPFormFieldProps } from '../SPFormField';
import DateFormField from './DateFormField';
import '../SPFormField.scss';

const styles = {
	dropDownFormField: 'dropDownFormField_e5e89a2f',
	dateFormField: 'dateFormField_e5e89a2f',
	unsupportedFieldMessage: 'unsupportedFieldMessage_e5e89a2f'
};

const SPFieldDateEdit: React.FunctionComponent<ISPFormFieldProps> = (props) => {
	const locale = Locales[props.fieldSchema.LocaleId];
	return (
		<DateFormField
			{...(props.value && moment(props.value).isValid() ? { value: moment(props.value).toDate() } : {})}
			className={css(styles.dateFormField, 'ard-dateFormField')}
			placeholder={'Enter a date'}
			isRequired={props.fieldSchema.Required}
			ariaLabel={props.fieldSchema.Title}
			locale={Locales[locale]}
			firstDayOfWeek={props.fieldSchema.FirstDayOfWeek}
			allowTextInput={true}
			onSelectDate={(date) => props.valueChanged(date!.toLocaleDateString(locale))}
		/>
	);
};

export default SPFieldDateEdit;
