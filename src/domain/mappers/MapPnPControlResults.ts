import { ContextService } from "../../services/ContextService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { User } from "../models/types/User";
import { Enums } from "../../globals/Enums";

export default class MapPnPControlResults extends ContextService {

    constructor(AppContext: WebPartContext, Lcid: number) {
        super(AppContext);
    }

    // Mapping results based on provided type
    public static map(item: any, type: Enums.ItemResultType): any[] {
        let result: any;
        switch (type) {
            case Enums.ItemResultType.User: result = this.mapUser(item);
                break;
        }
        return result;
    }

    //#region "Common Mappers"

    // Mapping multiple user
    private static mapUsers(userEntries: any): User[] {
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
    private static mapUser(user: any): User {
        // This in required, as in CAML it returns array even if it is single user
        if (user instanceof Array && user.length > 0) {
            user = user[0];
        }
        // Case : when it is null
        if (!user) {
            return new User();
        }
        let result: User = new User();
        result.Email = user["secondaryText"];
        result.Id = user["id"];
        result.LoginName = user["loginName"];
        result.Title = user["text"];

        if(result.Email == null || result.Email == ""){
            result.Email = result.LoginName.split("|").pop();
        }
        
        // Has Last Name, First Name pattern
        if(result.Title.indexOf(",") >= 0){
            let titleParts = result.Title.split(",");
            result.LastName = titleParts[0];
            result.FirstName = result.Title.replace(result.LastName + ",", "").trim();
        }
        // Can be First Name Last Name
        else{
            let titleParts = result.Title.split(" ");
            result.FirstName = titleParts[0];
            result.LastName = result.Title.replace(result.FirstName, "").trim();
        }
        return result;
    }

    //#endregion
}



