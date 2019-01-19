import * as React from 'react';
import { ISPFormFieldProps } from '../SPFormField';
import { Link } from 'office-ui-fabric-react/lib/Link';

const SPFieldUserDisplay: React.FunctionComponent<ISPFormFieldProps> = (props) => {
	if (props.value && props.value.length > 0) {
		const baseUrl = `${props.fieldSchema.ListFormUrl}?PageType=4&ListId=${props.fieldSchema.UserInfoListId}`;
		return (
			<div>
				{props.value.map((val: any) => (
					<div>
						<Link href={`${baseUrl}&ID=${val.id}`}>{val.title}</Link>
					</div>
				))}
			</div>
		);
	} else {
		return <div />;
	}
};

export default SPFieldUserDisplay;
