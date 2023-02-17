import { Enums } from "../../globals/Enums";
import MapCAMLResult from "./MapCAMLResult";
import MapPnPControlResults from "./MapPnPControlResults";
import MapPNPResult from "./MapPNPResult";

export default class MapResult {

    public static async map(items: any, mapperType: Enums.MapperType, itemResultType: Enums.ItemResultType) {

        // If no return type mentioned, then return the original values
        if (itemResultType == Enums.ItemResultType.None) {
            return items.Row;
        }

        switch (mapperType) {
            case Enums.MapperType.CAMLResult:
                return await MapCAMLResult.map(items, itemResultType);
                break;
            case Enums.MapperType.PnPControlResult:
                return await MapPnPControlResults.map(items, itemResultType);
                break;
            case Enums.MapperType.PNPResult:
                return await MapPNPResult.map(items, itemResultType);
                break;
        }
    }
}