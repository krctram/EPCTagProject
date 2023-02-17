import { WebPartContext } from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";

export class ContextService {
    private context: WebPartContext;
    constructor(appContext: WebPartContext){
        this.context = appContext;
        sp.setup({
            ie11: true,
            spfxContext: this.context,
            sp: {
            headers: {
                Accept: "application/json;odata=verbose"
              },
              baseUrl: this.context.pageContext.site.absoluteUrl,
            },
        } );
    }
}