import * as React from 'react';
import { ISPFormFieldProps } from '../SPFormField';
import '../SPFormField.scss';
import { PeoplePicker } from '../peoplepicker/PeoplePickerComponent';
import { PrincipalType } from '../peoplepicker/PrincipalType';

const SPUserEdit: React.FunctionComponent<ISPFormFieldProps> = (props) => {
	const personSelectionLimit = props.fieldSchema.AllowMultipleValues ? 15 : null;
	return (
		<PeoplePicker
			isRequired={props.fieldSchema.Required}
			webAbsoluteUrl={props.webUrl}
			principleTypes={[ PrincipalType.SecurityGroup, PrincipalType.SharePointGroup, PrincipalType.User ]}
			selectedItems={(users: any[]) => {
				var conArray: string[] = [];
				for (let user of users) {
					conArray.push('{\'Key\':\'' + user.loginName + '\'}');
				}
				props.valueChanged('[' + conArray.join() + ']');
			}}
			key={'peopleFieldId'}
			titleText={''}
			personSelectionLimit={personSelectionLimit}
			showtooltip={true}
			showHiddenInUI={false}
		/>
	);
};

export default SPUserEdit;
