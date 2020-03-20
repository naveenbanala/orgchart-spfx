export const BaseGraphUrl = "https://graph.microsoft.com"
export const GetManagerInfo = `${BaseGraphUrl}/v1.0/contacts/{id}/manager`
export const MyProperties = `${BaseGraphUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`
//export const AllUsers = `${BaseGraphUrl}/v1.0/users?$select=displayName,id,homepage,city,country,department,jobTitle,mail,mobile,postalCode,state,streetAddress,surname,telephoneNumber,thumbnailPhoto,manager,directReports`
export const AllUsers = `${BaseGraphUrl}/v1.0/users?$select=businessPhones,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,surname,preferredLanguage,surname,userPrincipalName,id,department`
export const UserPhoto = `${BaseGraphUrl}` + "/v1.0/users/{id}/photo/$value"
export const UserManager = `${BaseGraphUrl}` + "/v1.0/users/{id}/manager?$select=businessPhones,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,surname,preferredLanguage,surname,userPrincipalName,id,department"
export const UserDirectReports = `${BaseGraphUrl}` + "/v1.0/users/{id}/directReports?$select=businessPhones,displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,surname,preferredLanguage,surname,userPrincipalName,id,department"
export const UserPeers = `${BaseGraphUrl}` + "/v1.0/users/{id}/manager"
export const MyManager = `${BaseGraphUrl}` + "/v1.0/me/manager"
export const MyDirectReports = `${BaseGraphUrl}` + "/v1.0/me/directReports"
export const UserTypes = {
    peer: "peer",
    manager: "manager",
    directreports: "directreports"
}
export const MyDetails = `${BaseGraphUrl}` + "/v1.0/me"
