import { ISPUsers, ISPUser } from '../Model/DataModel';
import {
    SPHttpClient,
    SPHttpClientResponse,
    MSGraphClient,
    AadHttpClient
} from '@microsoft/sp-http';
import * as constants from './Constants'


export default class ServiceStore {

    public static getUser(cxt, id): Promise<ISPUser> {
        return cxt.aadHttpClientFactory
            .getClient(constants.BaseGraphUrl.replace('{id}', id))
            .then((client: AadHttpClient) => {
                return client.get(constants.GetManagerInfo, AadHttpClient.configurations.v1);
            }).then(resp => {
                resp.json()
            });
    }

    public static getAllUsers(context): Promise<ISPUsers> {
        return context.aadHttpClientFactory
            .getClient(constants.BaseGraphUrl)
            .then((client: AadHttpClient) => {
                return client
                    .get(
                        constants.AllUsers,
                        AadHttpClient.configurations.v1
                    );
            })
            .then(response => {
                return response.json();
            });
    }

    public static getUserphoto(context, id): Promise<any> {
        return context.aadHttpClientFactory
            .getClient(constants.BaseGraphUrl)
            .then((client: AadHttpClient) => {
                return client
                    .get(
                        constants.UserPhoto.replace("{id}", id),
                        AadHttpClient.configurations.v1
                    );
            })
            .then(response => {
                return response;
            });
    }

    public static getUserInfo(context, url): Promise<any> {
        return context.aadHttpClientFactory
            .getClient(constants.BaseGraphUrl)
            .then((client: AadHttpClient) => {
                return client
                    .get(
                        url,
                        AadHttpClient.configurations.v1
                    );
            })
            .then(response => {
                return response.json();
            });
    }

    public static getCurrentUser(cxt): Promise<ISPUser> {
        return cxt.aadHttpClientFactory
            .getClient(constants.BaseGraphUrl)
            .then((client: AadHttpClient) => {
                return client.get(constants.MyProperties, AadHttpClient.configurations.v1);
            }).then(resp => {
                resp.json()
            });
    }

    public static getCurrentUserProps(context): Promise<any> {
        return context.spHttpClient.get(context.pageContext.web.absoluteUrl + `/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            });
    }

    public static getUserProps(context, login): Promise<any> {
        return context.spHttpClient.get(context.pageContext.web.absoluteUrl + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='" + login.replace("#", "%23") + "'", SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            });
    }
}