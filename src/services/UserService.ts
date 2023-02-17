import { ContextService } from "./ContextService";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import "@pnp/sp/profiles";
import { User } from "../domain/models/types/User";
import MapResult from "../domain/mappers/MapResult";
import { Enums } from "../globals/Enums";

export default class UserService extends ContextService {

    constructor(AppContext: any) {
        super(AppContext);
    }

    // Get details of current user
    public async GetCurrentUser(): Promise<User> {
        const userItem = await sp.web.currentUser.get();
        let result: User = await MapResult.map(userItem, Enums.MapperType.PNPResult, Enums.ItemResultType.User);
        return result;
    }

    // Get details of current user using Profile Server
    public async GetCurrentUserProfile() : Promise<User>{
        const userItem = await sp.profiles.myProperties.get();
        let result: User = await MapResult.map(userItem, Enums.MapperType.PNPResult, Enums.ItemResultType.UserProfile);
        result.Id = await (await sp.web.currentUser.get()).Id;
        //console.log(result);
        return result;
    }

    // Get details of specific user using Profile Server
    public async GetUserProfileDetailsByLoginName(loginName: string) : Promise<User>{
        const userItem = await sp.profiles.getPropertiesFor(loginName);
        let result: User = await MapResult.map(userItem, Enums.MapperType.PNPResult, Enums.ItemResultType.UserProfile);
        result.Id = await (await sp.web.siteUsers.getByLoginName(loginName).get()).Id;
        return result;
    }

    // Get details of user from login name
    public async GetUserByLoginName(loginName: string): Promise<User> {
        const userItem = await sp.web.siteUsers.getByLoginName(loginName).get();
        let result: User = await MapResult.map(userItem, Enums.MapperType.PNPResult, Enums.ItemResultType.User);
        return result;
    }

    // It will check whether current user is Site Collection Admin or not
    public async CheckCurrentUserIsAdmin() : Promise<boolean>{
        const currentUser = await sp.web.currentUser.get();
        return currentUser.IsSiteAdmin;
    }

    // This function will check whether current user is member of provided SP group or not
    public async CheckCurrentUserInSPGroup(groupName : string) : Promise<boolean>{
        const groups = await sp.web.currentUser.groups();
        let result: boolean = false;
        groups.forEach((groupInfo) => {
            if(groupInfo.Title == groupName){
                result = true;
            }
        });
        return result;
    }
}