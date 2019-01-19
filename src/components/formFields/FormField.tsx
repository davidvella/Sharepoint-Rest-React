import * as React from 'react';

import { Label } from 'office-ui-fabric-react/lib/Label';
import { css, DelayedRender } from 'office-ui-fabric-react/lib/Utilities';

import { ControlMode } from '../../common/datatypes/ControlMode';

import '../FormField.module.scss';

const ardStyles = {
	formField: 'formField_d444b146',
	label: 'label_d444b146',
	controlContainerDisplay: 'controlContainerDisplay_d444b146'
};

export interface IFormFieldProps {
	className?: string;
	controlMode: ControlMode;
	label?: string;
	description?: string;
	required?: boolean;
	disabled?: boolean;
	active?: boolean;
	value: any;
	errorMessage?: string;
	webUrl: string;
	valueChanged(newValue: any): void;
}

const FormField: React.FunctionComponent<IFormFieldProps> = (props) => {
	const { children, className, description, disabled, label, required, active, errorMessage } = props;
	const formFieldClassName = css('ard-formField', ardStyles.formField, className, {
		['is-required ']: required,
		['is-disabled ']: disabled,
		['is-active ']: active
	});
	const isDescriptionAvailable = Boolean(props.description || props.errorMessage);

	return (
		<div className={css(formFieldClassName, 'od-ClientFormFields-field')}>
			<div className={'ard-FormField-wrapper'}>
				{label && (
					<Label className={ardStyles.label} htmlFor={this._id} required={props.required}>
						{label}
					</Label>
				)}
				<div className={css('ard-FormField-fieldGroup', ardStyles.controlContainerDisplay)}>{children}</div>
			</div>
			{isDescriptionAvailable && (
				<span>{description && <span className={'ard-FormField-description'}>{description}</span>}</span>
			)}
		</div>
	);
};

export default FormField;
