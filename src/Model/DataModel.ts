export interface ISPUsers {
    value: ISPUser[];
}

export interface ISPUser {
    businessPhones: Array<string>;
    displayName: string;
    givenName: string;
    jobTitle: string;
    mail: string;
    mobilePhone: string;
    officeLocation: string;
    preferredLanguage: string;
    surname: string;
    userPrincipalName: string;
    id: string;
    photo: string;
    context:any;
}

export interface ISPState {
    User: ISPUser;
    UserCollection: ISPUsers

}