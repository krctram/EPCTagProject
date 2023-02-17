import { ContextService } from "../../services/ContextService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as moment from 'moment';
import { User } from "../models/types/User";
import { Enums } from "../../globals/Enums";

export default class MapPNPResult extends ContextService {

    constructor(AppContext: WebPartContext, Lcid: number) {
        super(AppContext);
    }

    // Mapping results based on provided type
    public static async map(items: any, type: Enums.ItemResultType): Promise<any> {
        let result: any;
        switch (type) {
            case Enums.ItemResultType.User:
                result = this.mapUser(items);
                break;
            case Enums.ItemResultType.UserProfile:
                result = this.mapUserProfile(items);
                break;
            case Enums.ItemResultType.Users:
                result = this.mapUsers(items);
                break;
        }
        return result;
    }

    // Mapping multiple user
    public static mapUsers(userEntries: any): User[] {
        let result: User[] = [];
        if (userEntries instanceof Array) {
            userEntries.forEach(user => {
                result.push(this.mapUser(user));
            });
        }
        else {
            result.push(this.mapUser(userEntries));
        }

        return result;
    }

    // Mapping single user
    public static mapUser(user: any): User {
        if (user instanceof Array && user.length > 0) {
            user = user[0];
        }
        // Case : when it is null
        if (!user) {
            return undefined;
        }
        let result: User = new User();
        result.Email = user["Email"];
        result.Id = user["Id"];
        result.LoginName = user["LoginName"];
        result.Title = user["Title"];

        if(result.Email == null || result.Email == ""){
            result.Email = result.LoginName.split("|").pop();
        }
        return result;
    }

     // Mapping single user profile information
     public static mapUserProfile(user: any): User {
        
        if (user instanceof Array && user.length > 0) {
            user = user[0];
        }
        // Case : when it is null
        if (!user) {
            return undefined;
        }
        
        let result: User = new User();
        result.Email = user["Email"];
        result.LoginName = user["AccountName"];
        result.Title = user["DisplayName"];

        // Properties are stored in Key/Value pairs,
        // so parse into an object called userProperties
        var props = {};
        user.UserProfileProperties.results.forEach((prop) => {
            props[prop.Key] = prop.Value;
        });
        user.userProperties = props;
        result.FirstName = user.userProperties.FirstName;
        result.LastName = user.userProperties.LastName;
        
        if(result.Email == null || result.Email == ""){
            result.Email = result.LoginName.split("|").pop();
        }
        return result;
    }

    // Mapping boolean value
    private mapBoolean(itemValue: any): boolean {
        if (itemValue) {
            let result: boolean;
            result = (itemValue == "Yes" || itemValue.value == "1") ? true : false;
            return result;
        }
        return undefined;
    }

    // Mapping date field
    private mapDate(dateField: any): Date {
        if (dateField) {
            return (dateField);
        }
        return undefined;
    }

    // Mapping date field and return formatted date string
    private mapDateWithFormat(dateField: any): string {
        if (dateField) {
            return (moment(dateField).format('D MMMM YYYY'));
        }
        return "";
    }
}
















