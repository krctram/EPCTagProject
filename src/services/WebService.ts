import { ContextService } from "./ContextService";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";

export default class WebService extends ContextService {

    constructor(AppContext: any) {
        super(AppContext);
    }

    // Get details of current user
    public async GetAssociatedOwnerGroupName(): Promise<string> {
        let result = "";
        try {
            const ownerGroup = await sp.web.associatedOwnerGroup();
            result = ownerGroup.Title;
        }
        catch (ex) {
            // Handling the error, if any
        }
        return result;
    }
}