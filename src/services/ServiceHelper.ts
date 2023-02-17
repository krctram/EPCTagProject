import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/content-types";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import { Config } from "../globals/Config";


// Contains static methods which can be used wherver required in any of the service classes
export default class ServiceHelper {

    // Generating CAML query XML based on available data
    public static generateCAMLQueryXML(selectFields: string[],  orderByXML: string, camlFilterConditions: string = "", rowLimit: number = Config.List_ThresholdLimit): string {

        // For "View Fields"
        let viewFieldsXML = "";
        if (selectFields.length > 0) {
            viewFieldsXML = "<ViewFields>";
            selectFields.map((fieldName) => {
                viewFieldsXML += "<FieldRef Name='" + fieldName + "'/>";
            });
            viewFieldsXML += "</ViewFields>";
        }

        // For "Ordering"
        // Currently ordering by ID, but can be customizable on later stage
        if(orderByXML == undefined){
            orderByXML = "<OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy>";
        }

        // For "Row Limit"
        if (rowLimit > Config.List_ThresholdLimit) {
            // Limiting the rows to list view threshold
            rowLimit = Config.List_ThresholdLimit;
        }
        let rowLimitXML = "<RowLimit Paged='TRUE'>" + rowLimit + "</RowLimit>";

        // For "Filter Criteris"
        // Note: Currently user has to provide CAML conditions starting with "<Where>" clause
        // Ex: <Where><Eq>....</Eq>></Where>
        // TODO: Conditions string can be generated dynamically at later stage based on requirements
        let conditionsXML = camlFilterConditions;

        let viewXML = `<View><Query>${orderByXML}${conditionsXML}</Query>${viewFieldsXML}${rowLimitXML}</View>`;
        return viewXML;
    }

}