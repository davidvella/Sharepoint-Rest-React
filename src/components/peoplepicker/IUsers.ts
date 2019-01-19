export interface IUsers {
	'@odata.context': string;
	value: IValue[];
}

export interface IValue {
	'@odata.type': string;
	'@odata.id': string;
	'@odata.editLink': string;
	Id: number;
	IsHiddenInUI: boolean;
	LoginName: string;
	Title: string;
	PrincipalType: number;
	Email: string;
	IsEmailAuthenticationGuestUser: boolean;
	IsShareByEmailGuestUser: boolean;
	IsSiteAdmin: boolean;
	UserId?: IUserId;
}

export interface IUserId {
	NameId: string;
	NameIdIssuer: string;
}
